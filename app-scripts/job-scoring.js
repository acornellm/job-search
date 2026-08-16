/**
 * job-scoring.gs — Google Apps Script (third file, after findJobLinks.gs and jobTriage.gs)
 *
 * Scores triaged jobs against your resume, a weighted rubric, and calibration
 * examples, then writes each one as a row on a "Job Score" sheet — your
 * working application tracker. Job Triage is read-only here; nothing is
 * written back to it.
 *
 * The three context sources:
 *   1. Resume  — a Google Doc you keep editing normally (canonical source).
 *   2. Rubric  — a "Scoring Rubric" sheet: must-haves, dealbreakers, and the
 *                equivalence rules that decide what counts as a match.
 *   3. Calibration — a "Calibration" sheet of postings you've already scored
 *                yourself, so the 1-100 scale means the same thing every run.
 *
 * All three are assembled into a byte-stable system prompt with a cache
 * breakpoint, so only the job details vary between calls. Cache reads run
 * ~0.1x base input price:
 * https://docs.claude.com/en/docs/build-with-claude/prompt-caching
 *
 * Depends on jobTriage.gs for: getApiKey_, getUiOrNull_, extractText_,
 * parseClaudeJson_, formatDate_, TRIAGE_SHEET_NAME, TRIAGE_DEFAULT_LIMIT.
 *
 * Setup:
 *   1. Add this file to the same Apps Script project.
 *   2. Run setupScoringSheets() — creates the Rubric and Calibration sheets
 *      with starter rows. Edit them; they are the whole point.
 *   3. Put your master resume in a Google Doc, copy its ID from the URL
 *      (docs.google.com/document/d/THIS_PART/edit), and either run
 *      setResumeDocId() or add RESUME_DOC_ID under Script Properties.
 *   4. Run previewScoringPrompt() to see exactly what the model will receive.
 *   5. Run scoreTriagedJobs().
 *
 * Usage:
 *   scoreTriagedJobs();                      // 20 new jobs -> Job Score rows
 *   scoreTriagedJobs({ limit: 25 });
 *   scoreTriagedJobs({ rescore: true });     // re-run scoring, keeps your Status
 *   scoreTriagedJobs({ limit: 1, dryRun: true });
 *   sortScoreSheet();                        // highest score to the top
 *
 * scoreTriagedJobs() and scoreOneUrl() both call sortScoreSheet() after
 * writing, so the sheet re-sorts itself automatically — no need to run it
 * by hand unless you've reordered rows some other way (e.g. by hand-editing
 * Score).
 *
 * Status, My Notes, and Applied Date belong to you — rescoring never
 * overwrites them.
 */

// --- Configuration --------------------------------------------------------

const SCORE_SHEET_NAME        = 'Job Score';
const RUBRIC_SHEET_NAME       = 'Scoring Rubric';
const CALIBRATION_SHEET_NAME  = 'Calibration';
const RESUME_DOC_ID_PROPERTY  = 'RESUME_DOC_ID';

// Status column is a dropdown you drive by hand.
const STATUS_OPTIONS = [
  'New', 'Interested', 'Applying', 'Applied', 'Skip',
];
const STATUS_DEFAULT = 'New';

// Sheet layout. Status and Score lead so the sheet is scannable at a glance.
// Top Keywords / Technical Skills are copied over from Job Triage at scoring
// time so the posting's substance is visible here without switching sheets.
const SCORE_SHEET_HEADERS = [
  'Status', 'Score', 'Verdict', 'Role Title', 'Company', 'Locations',
  'Salary Range', 'Posting Date', 'Top Keywords', 'Technical Skills',
  'Strengths', 'Gaps', 'Score Notes',
  'My Notes', 'Applied Date', 'Source', 'URL', 'Email Link', 'Scored At',
];

// Yours. Written once when the row is created, never touched again.
const USER_OWNED_COLUMNS = ['Status', 'My Notes', 'Applied Date'];

const SCORING_MODEL          = 'claude-sonnet-5';
const SCORING_MAX_TOKENS     = 1500;
const SCORING_DEFAULT_LIMIT  = 20;
const SCORING_SLEEP_MS       = 800;   // keeps calls inside the 5-min cache TTL
const SCORING_MAX_RUNTIME_MS = 5 * 60 * 1000;

// Sonnet 5 caches prefixes of 1,024 tokens and up; shorter ones are silently
// processed uncached. ~4 chars per token, so warn below ~4,100 characters.
const CACHE_MIN_CHARS = 4100;

const VERDICTS = ['Strong fit', 'Worth applying', 'Stretch', 'Pass'];

// --- Resume (source 3: the Google Doc) ------------------------------------

/** Store the resume Doc ID. Falls back to logging when there's no UI. */
function setResumeDocId() {
  const ui = getUiOrNull_();
  if (!ui) {
    Logger.log('Cannot prompt here. Add it manually: Project Settings -> ' +
               'Script Properties -> %s = <doc id>', RESUME_DOC_ID_PROPERTY);
    return;
  }
  const res = ui.prompt(
    'Master resume Doc',
    'Paste the Google Doc ID (the long string between /d/ and /edit in the URL).',
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return;
  const id = res.getResponseText().trim();
  if (!id) { ui.alert('Nothing entered — not changed.'); return; }

  PropertiesService.getScriptProperties().setProperty(RESUME_DOC_ID_PROPERTY, id);
  ui.alert('Saved. Run previewScoringPrompt() to confirm it reads.');
}

function getResumeDocId_() {
  const id = PropertiesService.getScriptProperties().getProperty(RESUME_DOC_ID_PROPERTY);
  if (!id) {
    throw new Error('No resume Doc ID stored. Run setResumeDocId(), or add ' +
                    RESUME_DOC_ID_PROPERTY + ' under Script Properties.');
  }
  return id;
}

/**
 * Read the master resume. Whitespace is normalized so an invisible edit in the
 * Doc doesn't silently invalidate the prompt cache.
 */
function loadProfileText_() {
  const doc = DocumentApp.openById(getResumeDocId_());
  const text = doc.getBody().getText();

  if (!text || text.trim().length < 200) {
    throw new Error('Resume Doc "' + doc.getName() + '" looks empty.');
  }
  return text.replace(/\r/g, '').replace(/[ \t]+/g, ' ')
             .replace(/\n{3,}/g, '\n\n').trim();
}

// --- Rubric (source 4: what should count as a match) ----------------------

/**
 * Parse the Scoring Rubric sheet.
 * Columns: Type | Criterion | Weight | Notes
 *   Type is one of: Must-have, Nice-to-have, Dealbreaker, Equivalence
 *   Weight is 1-5 for the first two, ignored for the rest.
 */
function loadRubric_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RUBRIC_SHEET_NAME);
  if (!sheet) throw new Error('No "' + RUBRIC_SHEET_NAME + '" sheet. Run setupScoringSheets().');
  if (sheet.getLastRow() < 2) throw new Error('Rubric sheet is empty — add some criteria.');

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const cType = header.indexOf('type');
  const cCrit = header.indexOf('criterion');
  const cWt   = header.indexOf('weight');
  const cNote = header.indexOf('notes');

  if (cType === -1 || cCrit === -1) {
    throw new Error('Rubric sheet needs at least "Type" and "Criterion" columns.');
  }

  const out = { mustHave: [], niceToHave: [], dealbreaker: [], equivalence: [] };

  for (let r = 1; r < values.length; r++) {
    const crit = String(values[r][cCrit] || '').trim();
    if (!crit) continue;

    const type = String(values[r][cType] || '').trim().toLowerCase().replace(/[\s_]/g, '-');
    const item = {
      criterion: crit,
      weight: cWt > -1 ? Number(values[r][cWt]) || 3 : 3,
      notes: cNote > -1 ? String(values[r][cNote] || '').trim() : '',
    };

    if (type.indexOf('dealbreaker') === 0) out.dealbreaker.push(item);
    else if (type.indexOf('equivalence') === 0) out.equivalence.push(item);
    else if (type.indexOf('nice') === 0) out.niceToHave.push(item);
    else out.mustHave.push(item);
  }
  return out;
}

/** Flatten the rubric into deterministic prompt text. */
function rubricToText_(rubric) {
  const section = function (title, items, showWeight) {
    if (!items.length) return title + ':\n  (none)';
    const lines = items.map(function (i) {
      const w = showWeight ? ' [weight ' + i.weight + '/5]' : '';
      const n = i.notes ? ' — ' + i.notes : '';
      return '  - ' + i.criterion + w + n;
    });
    return title + ':\n' + lines.join('\n');
  };

  return [
    section('MUST-HAVES (absence should pull the score down hard)', rubric.mustHave, true),
    '',
    section('NICE-TO-HAVES (presence lifts the score)', rubric.niceToHave, true),
    '',
    section('DEALBREAKERS (any one of these caps the verdict at Pass)', rubric.dealbreaker, false),
    '',
    section('EQUIVALENCE RULES (how to credit adjacent experience)', rubric.equivalence, false),
  ].join('\n');
}

// --- Calibration (source 5: anchoring the scale) --------------------------

/**
 * Parse the Calibration sheet.
 * Columns: Role Title | Company | Key Details | My Score | Reasoning
 */
function loadCalibration_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALIBRATION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const cTitle = header.indexOf('role title');
  const cComp  = header.indexOf('company');
  const cDet   = header.indexOf('key details');
  const cScore = header.indexOf('my score');
  const cWhy   = header.indexOf('reasoning');

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const score = Number(values[r][cScore]);
    const title = String(values[r][cTitle] || '').trim();
    if (!title || !score) continue;   // unscored rows are placeholders

    out.push({
      title: title,
      company: cComp > -1 ? String(values[r][cComp] || '').trim() : '',
      details: cDet > -1 ? String(values[r][cDet] || '').trim() : '',
      score: score,
      why: cWhy > -1 ? String(values[r][cWhy] || '').trim() : '',
    });
  }
  return out;
}

function calibrationToText_(examples) {
  if (!examples.length) {
    return 'CALIBRATION EXAMPLES:\n  (none yet — scores will drift between runs ' +
           'until at least three are added)';
  }
  const lines = examples.map(function (e) {
    return [
      '  - ' + e.title + (e.company ? ' at ' + e.company : ''),
      '    Posting: ' + (e.details || 'n/a'),
      '    My score: ' + e.score,
      '    Why: ' + (e.why || 'n/a'),
    ].join('\n');
  });
  return 'CALIBRATION EXAMPLES (scores I assigned myself — match this scale):\n' +
         lines.join('\n');
}

// --- Prompt assembly ------------------------------------------------------

const SCORING_INSTRUCTIONS = [
  'You score job postings for a single candidate. You reply with one JSON ',
  'object and nothing else — no markdown fences, no commentary.',
  '',
  'Scoring approach:',
  '- Judge capability transfer, not keyword overlap. A posting naming a tool ',
  '  the candidate has not used is a real gap only if the underlying ',
  '  capability is also absent. Apply the equivalence rules below before ',
  '  deciding something is missing.',
  '- Weight must-haves heavily and nice-to-haves lightly. Any dealbreaker ',
  '  caps the verdict at "Pass" regardless of the rest.',
  '- Anchor the number to the calibration examples. If a posting resembles ',
  '  one of them, its score should land close to that example.',
  '- Be blunt about gaps. A generous score that gets the candidate a rejection ',
  '  is worse than an honest low one.',
  '',
  'Score bands: 85-100 Strong fit / 70-84 Worth applying / 50-69 Stretch / ',
  '0-49 Pass. The verdict must match the band unless a dealbreaker forces Pass.',
].join('\n');

const SCORING_OUTPUT_SCHEMA = [
  'Return exactly this shape:',
  '{',
  '  "score": number,              // 0-100 integer',
  '  "verdict": string,            // one of: ' + VERDICTS.join(' | '),
  '  "strengths": [string],        // up to 4, specific to this posting',
  '  "gaps": [string],             // up to 4, most disqualifying first',
  '  "dealbreakers": [string],     // [] if none triggered',
  '  "notes": string               // one sentence on the deciding factor',
  '}',
].join('\n');

/**
 * Build the system blocks. Everything here must be byte-identical across
 * calls or the cache never hits — no timestamps, no row counts, no ordering
 * that depends on which job is being scored.
 */
function buildScoringSystemBlocks_() {
  const context = [
    '=== CANDIDATE RESUME ===',
    loadProfileText_(),
    '',
    '=== SCORING RUBRIC ===',
    rubricToText_(loadRubric_()),
    '',
    '=== ' + calibrationToText_(loadCalibration_()),
  ].join('\n');

  return [
    { type: 'text', text: SCORING_INSTRUCTIONS },
    // Breakpoint goes on the last block that is identical across requests.
    { type: 'text', text: context, cache_control: { type: 'ephemeral' } },
  ];
}

/** The per-job half of the prompt — this is the only part that varies. */
function buildScoringUserPrompt_(job) {
  return [
    SCORING_OUTPUT_SCHEMA,
    '',
    'Score this posting:',
    'Role: ' + (job.role || 'unknown'),
    'Company: ' + (job.company || 'unknown'),
    'Locations: ' + (job.locations || 'unknown'),
    'Salary: ' + (job.salary || 'not stated'),
    'Posted: ' + (job.posted || 'unknown'),
    'Keywords from the posting: ' + (job.keywords || 'none extracted'),
    'Technical skills required: ' + (job.skills || 'none extracted'),
    'URL: ' + (job.url || ''),
  ].join('\n');
}

/**
 * Messages API call with a cached system prefix.
 * @return {Object} { text, usage }
 */
function callClaudeCached_(systemBlocks, userPrompt, opts) {
  opts = opts || {};
  const retries = opts.retries === undefined ? 3 : opts.retries;

  const payload = {
    model: SCORING_MODEL,
    max_tokens: opts.maxTokens || SCORING_MAX_TOKENS,
    system: systemBlocks,
    messages: [{ role: 'user', content: userPrompt }],
  };

  let lastErr = '';

  for (let attempt = 0; attempt <= retries; attempt++) {
    if (attempt > 0) Utilities.sleep(Math.pow(2, attempt) * 1000);

    const res = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': getApiKey_(),
        'anthropic-version': ANTHROPIC_VERSION,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code === 200) {
      const data = JSON.parse(body);
      return { text: extractText_(data), usage: data.usage || {} };
    }

    lastErr = 'HTTP ' + code + ': ' + body.slice(0, 400);
    if (code !== 429 && code < 500) break;
    Logger.log('Retrying after %s', lastErr);
  }

  throw new Error('Claude API call failed — ' + lastErr);
}

// --- Job Score sheet ------------------------------------------------------

/** Header name -> 1-based column index. */
function headerMap_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach(function (h, i) {
    const name = String(h).trim();
    if (name) map[name] = i + 1;
  });
  return map;
}

/** Get the Job Score sheet, creating it with headers and a Status dropdown. */
function getOrCreateScoreSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCORE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SCORE_SHEET_NAME);
    sheet.appendRow(SCORE_SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(3);   // Status / Score / Verdict stay visible
    sheet.getRange(1, 1, 1, SCORE_SHEET_HEADERS.length).setFontWeight('bold');
    applyStatusValidation_(sheet);
    Logger.log('Created "%s".', SCORE_SHEET_NAME);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(SCORE_SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SCORE_SHEET_HEADERS.length).setFontWeight('bold');
    applyStatusValidation_(sheet);
  } else {
    ensureScoreSheetColumns_(sheet);
  }
  return sheet;
}

/**
 * Add any SCORE_SHEET_HEADERS columns missing from an existing sheet — e.g.
 * "Top Keywords" / "Technical Skills" added after the sheet already had
 * rows. Appended at the end, same idempotent pattern as
 * ensureTailoringColumns_() in job-tailoring.gs: never reorders or touches
 * existing columns, so it's safe to call on every run.
 */
function ensureScoreSheetColumns_(sheet) {
  const width = sheet.getLastColumn();
  const header = width > 0
    ? sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); })
    : [];
  const missing = SCORE_SHEET_HEADERS.filter(function (h) { return header.indexOf(h) === -1; });
  if (!missing.length) return;

  sheet.getRange(1, width + 1, 1, missing.length).setValues([missing]);
  sheet.getRange(1, width + 1, 1, missing.length).setFontWeight('bold');
  Logger.log('Added column(s) to "%s": %s', SCORE_SHEET_NAME, missing.join(', '));
}

/** Dropdown on the Status column. Invalid values are allowed but flagged. */
function applyStatusValidation_(sheet) {
  const col = headerMap_(sheet)['Status'];
  if (!col) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(true)
    .setHelpText('Where this application stands.')
    .build();

  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

/** URL -> row number for everything already on the Job Score sheet. */
function readScoredUrls_(sheet) {
  const map = {};
  if (sheet.getLastRow() < 2) return map;

  const col = headerMap_(sheet)['URL'];
  const urls = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();

  urls.forEach(function (r, i) {
    const u = String(r[0] || '').trim();
    if (u) map[u] = i + 2;
  });
  return map;
}

/** Read scorable jobs out of Job Triage. Triage is never modified. */
function readTriagedJobs_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRIAGE_SHEET_NAME);
  if (!sheet) throw new Error('No "' + TRIAGE_SHEET_NAME + '" sheet — run triageJobLinks() first.');
  if (sheet.getLastRow() < 2) return [];

  const col = headerMap_(sheet);
  const values = sheet.getDataRange().getValues();
  const get = function (row, name) {
    return col[name] ? String(row[col[name] - 1] || '') : '';
  };

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (get(row, 'Status').trim().toUpperCase() !== 'OK') continue;   // no data on skipped/errored rows

    const url = get(row, 'URL').trim();
    if (!url) continue;

    out.push({
      url: url,
      role: get(row, 'Role Title'),
      company: get(row, 'Company'),
      locations: get(row, 'Locations'),
      salary: get(row, 'Salary Range'),
      posted: formatDate_(col['Posting Date'] ? row[col['Posting Date'] - 1] : ''),
      keywords: get(row, 'Top Keywords'),
      skills: get(row, 'Technical Skills'),
      source: get(row, 'Source'),
      permalink: get(row, 'Email Link'),
    });
  }
  return out;
}

// --- Main routine ---------------------------------------------------------

/**
 * Score jobs from Job Triage into the Job Score sheet.
 *
 * @param {Object} [opts]
 * @param {number}  [opts.limit=20]  Max jobs to score this run.
 * @param {boolean} [opts.rescore]   Re-score rows already on Job Score.
 *                                   Status, My Notes, Applied Date are kept.
 * @param {boolean} [opts.dryRun]    Log results without writing.
 * @return {Object} { scored, added, updated, errors, remaining }
 *
 * Sorts Job Score by Score, highest first, after any write — see
 * sortScoreSheet(). Skipped on a dry run, since nothing was written.
 */
function scoreTriagedJobs(opts) {
  opts = opts || {};
  const limit = opts.limit === undefined ? SCORING_DEFAULT_LIMIT : opts.limit;
  const started = Date.now();

  const sheet = getOrCreateScoreSheet_();
  const col = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const already = readScoredUrls_(sheet);
  const jobs = readTriagedJobs_();

  const pending = jobs.filter(function (j) {
    return opts.rescore ? true : !already[j.url];
  });

  Logger.log('%s triaged job(s), %s pending, limit %s', jobs.length, pending.length, limit);
  if (!pending.length) return { scored: 0, added: 0, updated: 0, errors: 0, remaining: 0 };

  // Build the cached prefix once per run, not once per job.
  const systemBlocks = buildScoringSystemBlocks_();
  warnIfUncacheable_(systemBlocks);

  const batch = pending.slice(0, limit);
  const newRows = [];
  let errors = 0;
  let scored = 0;
  let updated = 0;

  for (let i = 0; i < batch.length; i++) {
    const job = batch[i];

    if (Date.now() - started > SCORING_MAX_RUNTIME_MS) {
      Logger.log('Runtime budget reached — stopping after %s of %s.', i, batch.length);
      break;
    }
    if (i > 0) Utilities.sleep(SCORING_SLEEP_MS);

    let result;
    try {
      const res = callClaudeCached_(systemBlocks, buildScoringUserPrompt_(job));
      logCacheUsage_(res.usage, i);
      result = normalizeScore_(parseClaudeJson_(res.text));
      scored++;
    } catch (e) {
      errors++;
      Logger.log('  %s failed: %s', job.url, e.message);
      result = { score: '', verdict: 'ERROR', strengths: '', gaps: '',
                 notes: e.message.slice(0, 300) };
    }

    Logger.log('[%s/%s] %s — %s (%s)', i + 1, batch.length,
               job.role || job.url, result.score, result.verdict);

    if (opts.dryRun) continue;

    const existingRow = already[job.url];
    if (existingRow) {
      updateScoreRow_(sheet, col, existingRow, job, result);
      updated++;
    } else {
      newRows.push(buildScoreRow_(job, result, col, width));
    }
  }

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, width)
      .setValues(newRows);
  }

  // Bring the highest score back to the top after any write. New rows land
  // at the bottom and updated scores can move either way, so this runs
  // whenever something actually changed.
  if (!opts.dryRun && (newRows.length || updated)) sortScoreSheet();

  const summary = {
    scored: scored,
    added: newRows.length,
    updated: updated,
    errors: errors,
    remaining: Math.max(0, pending.length - batch.length),
  };
  Logger.log('Done: %s', JSON.stringify(summary));
  return summary;
}

/**
 * Assemble a full new row, placed by column name via `col` rather than a
 * fixed index — a migrated sheet can have columns in a different order
 * than SCORE_SHEET_HEADERS (new ones land appended at the end), so this
 * must not assume position. `width` is the row's length, i.e. the sheet's
 * current column count.
 */
function buildScoreRow_(job, result, col, width) {
  const byName = {
    'Status': STATUS_DEFAULT,
    'Score': result.score,
    'Verdict': result.verdict,
    'Role Title': job.role,
    'Company': job.company,
    'Locations': job.locations,
    'Salary Range': job.salary,
    'Posting Date': job.posted,
    'Top Keywords': job.keywords,
    'Technical Skills': job.skills,
    'Strengths': result.strengths,
    'Gaps': result.gaps,
    'Score Notes': result.notes,
    'My Notes': '',
    'Applied Date': '',
    'Source': job.source,
    'URL': job.url,
    'Email Link': job.permalink,
    'Scored At': new Date(),
  };

  const row = new Array(width).fill('');
  Object.keys(byName).forEach(function (name) {
    if (col[name]) row[col[name] - 1] = byName[name];
  });
  return row;
}

/**
 * Update an existing row's machine-owned cells only. Anything in
 * USER_OWNED_COLUMNS is left exactly as you left it.
 */
function updateScoreRow_(sheet, col, rowNum, job, result) {
  const updates = {
    'Score': result.score,
    'Verdict': result.verdict,
    'Role Title': job.role,
    'Company': job.company,
    'Locations': job.locations,
    'Salary Range': job.salary,
    'Posting Date': job.posted,
    'Top Keywords': job.keywords,
    'Technical Skills': job.skills,
    'Strengths': result.strengths,
    'Gaps': result.gaps,
    'Score Notes': result.notes,
    'Source': job.source,
    'Email Link': job.permalink,
    'Scored At': new Date(),
  };

  Object.keys(updates).forEach(function (name) {
    if (USER_OWNED_COLUMNS.indexOf(name) > -1) return;
    if (!col[name]) return;
    sheet.getRange(rowNum, col[name]).setValue(updates[name]);
  });
}

/** Coerce the model's JSON into sheet-ready values. */
function normalizeScore_(data) {
  const score = Math.max(0, Math.min(100, Math.round(Number(data.score) || 0)));

  let verdict = String(data.verdict || '').trim();
  if (VERDICTS.indexOf(verdict) === -1) {
    verdict = score >= 85 ? VERDICTS[0] : score >= 70 ? VERDICTS[1] :
              score >= 50 ? VERDICTS[2] : VERDICTS[3];
  }
  // A dealbreaker overrides whatever the model said.
  const breakers = Array.isArray(data.dealbreakers) ? data.dealbreakers.filter(String) : [];
  if (breakers.length) verdict = 'Pass';

  const join = function (v) {
    return Array.isArray(v) ? v.map(String).join('; ') : String(v || '');
  };

  return {
    score: score,
    verdict: verdict,
    strengths: join(data.strengths),
    gaps: join(data.gaps) + (breakers.length ? ' [DEALBREAKER: ' + join(breakers) + ']' : ''),
    notes: String(data.notes || '').slice(0, 500),
  };
}

function logCacheUsage_(usage, i) {
  if (!usage) return;
  const write = usage.cache_creation_input_tokens || 0;
  const read = usage.cache_read_input_tokens || 0;
  if (i === 0) {
    Logger.log('  cache write %s tokens, read %s', write, read);
  } else if (read === 0 && write === 0) {
    Logger.log('  no cache activity — prefix may be under the minimum length');
  }
}

/** Caching fails silently below the minimum, so say something up front. */
function warnIfUncacheable_(systemBlocks) {
  const chars = systemBlocks.reduce(function (n, b) { return n + b.text.length; }, 0);
  if (chars < CACHE_MIN_CHARS) {
    Logger.log('Context is ~%s chars — likely under the cacheable minimum for %s. ' +
               'It will still work, just at full input price each call.',
               chars, SCORING_MODEL);
  } else {
    Logger.log('Context ~%s chars, cacheable.', chars);
  }
}

// --- Sheet setup and housekeeping ----------------------------------------

/** Create the Rubric and Calibration sheets with starter rows. */
function setupScoringSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(RUBRIC_SHEET_NAME)) {
    const s = ss.insertSheet(RUBRIC_SHEET_NAME);
    s.appendRow(['Type', 'Criterion', 'Weight', 'Notes']);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 4).setFontWeight('bold');
    s.getRange(2, 1, SCORING_RUBRIC_SEED.length, 4).setValues(SCORING_RUBRIC_SEED);
    s.autoResizeColumns(1, 4);
    Logger.log('Created "%s" with %s starter rows — edit them.',
               RUBRIC_SHEET_NAME, SCORING_RUBRIC_SEED.length);
  } else {
    Logger.log('"%s" already exists — left alone.', RUBRIC_SHEET_NAME);
  }

  if (!ss.getSheetByName(CALIBRATION_SHEET_NAME)) {
    const s = ss.insertSheet(CALIBRATION_SHEET_NAME);
    s.appendRow(['Role Title', 'Company', 'Key Details', 'My Score', 'Reasoning']);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 5).setFontWeight('bold');
    s.getRange(2, 1, CALIBRATION_SEED.length, 5).setValues(CALIBRATION_SEED);
    s.autoResizeColumns(1, 5);
    Logger.log('Created "%s" — replace the placeholders with real postings you rate.',
               CALIBRATION_SHEET_NAME);
  } else {
    Logger.log('"%s" already exists — left alone.', CALIBRATION_SHEET_NAME);
  }

  getOrCreateScoreSheet_();
}

/** Sort Job Score by score, highest first. Run whenever, it's non-destructive. */
function sortScoreSheet() {
  const sheet = getOrCreateScoreSheet_();
  if (sheet.getLastRow() < 3) return;

  const col = headerMap_(sheet);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .sort([{ column: col['Score'], ascending: false }]);
  Logger.log('Sorted by score.');
}

// Starter rubric. These are guesses based on a Director-level data engineering
// search out of a large-scale Oracle platform background — edit freely, the
// equivalence rows are the ones that actually change your scores.
const SCORING_RUBRIC_SEED = [
  ['Must-have', 'Director level or above, owning a data engineering org', 5, 'Not a hands-on IC or team-lead role'],
  ['Must-have', 'Manages managers, or a team of 15+ engineers', 4, 'Scope should not be a step down'],
  ['Must-have', 'Data platform / data engineering domain', 5, 'Not general application engineering'],
  ['Nice-to-have', 'Modern data stack in the environment', 3, 'Snowflake, dbt, Airbyte, Fivetran, Databricks'],
  ['Nice-to-have', 'Distributed / large-scale processing', 4, 'Spark, Hadoop, Kafka, Cassandra'],
  ['Nice-to-have', 'Multi-region or globally distributed teams', 3, ''],
  ['Nice-to-have', 'Data governance and privacy ownership', 3, 'GDPR, CCPA, cataloging, lineage'],
  ['Nice-to-have', 'Platform cost / TCO accountability', 2, ''],
  ['Nice-to-have', 'BI and ML workflow enablement', 2, ''],
  ['Dealbreaker', 'Requires relocation outside Colorado with no remote option', 0, ''],
  ['Dealbreaker', 'Individual contributor role with a director title', 0, 'Title inflation at small startups'],
  ['Equivalence', 'Spark / Hadoop platform ownership counts as ~70% credit toward Databricks or Snowflake requirements', 0, 'Same architectural problems, different vendor'],
  ['Equivalence', 'Cassandra and large-scale OLTP counts toward modern warehouse modeling', 0, ''],
  ['Equivalence', 'OCI counts as cloud platform depth alongside AWS, GCP, Azure', 0, 'Do not penalize for the specific cloud'],
  ['Equivalence', 'Building internal ELT tooling counts toward dbt / Airbyte experience', 0, 'The pattern transfers; the tool is learnable'],
  ['Equivalence', 'Named tools absent from the resume are gaps only when no adjacent capability exists', 0, ''],
];

// Placeholders. Replace with real postings and your own numbers — three
// honest examples are worth more than sixteen invented ones.
const CALIBRATION_SEED = [
  ['Director, Data Engineering', '(example - replace)',
   'Remote, 40-person data org, Spark + Snowflake, $230-270k', 90,
   'Right scope, right domain, stack overlaps heavily'],
  ['Senior Manager, Data Platform', '(example - replace)',
   'Hybrid Denver, 12 engineers, mostly dbt/Airflow, $190-210k', 68,
   'Scope is a half step down and comp is light, but domain fits'],
  ['Director, Software Engineering', '(example - replace)',
   'Onsite Austin, mobile apps org, no data mandate', 30,
   'Right level and title, wrong domain, relocation required'],
];

// --- Inspection -----------------------------------------------------------

/** Print the exact system prompt without spending an API call. */
function previewScoringPrompt() {
  const blocks = buildScoringSystemBlocks_();
  blocks.forEach(function (b, i) {
    Logger.log('--- system block %s (%s chars%s) ---\n%s',
               i + 1, b.text.length, b.cache_control ? ', cached' : '', b.text);
  });
  warnIfUncacheable_(blocks);

  Logger.log('--- sample user prompt ---\n%s', buildScoringUserPrompt_({
    role: 'Director, Data Engineering', company: 'Example Co',
    locations: 'Remote', salary: '$220,000 - $260,000', posted: '2026-08-01',
    keywords: 'data platform, leadership, ELT', skills: 'Snowflake, dbt, Airflow',
    url: 'https://example.com/jobs/123',
  }));
}

/** Re-score one URL and log the raw response — the loop for tuning the rubric. */
function scoreOneUrl(url) {
  if (!url) { Logger.log('Pass a job URL, e.g. scoreOneUrl("https://...")'); return; }

  const job = readTriagedJobs_().filter(function (j) { return j.url === url; })[0];
  if (!job) { Logger.log('URL not found on the Job Triage sheet with Status OK.'); return; }

  const res = callClaudeCached_(buildScoringSystemBlocks_(), buildScoringUserPrompt_(job));
  Logger.log('Raw response:\n%s', res.text);

  const result = normalizeScore_(parseClaudeJson_(res.text));
  const sheet = getOrCreateScoreSheet_();
  const col = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const existing = readScoredUrls_(sheet)[url];

  if (existing) updateScoreRow_(sheet, col, existing, job, result);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, width)
        .setValues([buildScoreRow_(job, result, col, width)]);

  sortScoreSheet();
  return result;
}

// --- Menu -----------------------------------------------------------------

/**
 * Job Score menu. Called from onOpen() in jobTriage.gs — Apps Script permits
 * only one onOpen per project, so this is a builder rather than its own hook.
 */
function jobScoreMenu_(ui) {
  ui.createMenu('Job Score')
    .addItem('Score next batch', 'menuScoreBatch')
    .addItem('Score all pending', 'menuScoreAll')
    .addItem('Re-score everything', 'menuRescoreAll')
    .addSeparator()
    .addItem('Sort by score', 'sortScoreSheet')
    .addSeparator()
    .addItem('Set resume Doc ID', 'setResumeDocId')
    .addItem('Create rubric + calibration sheets', 'setupScoringSheets')
    .addItem('Preview scoring prompt (see logs)', 'previewScoringPrompt')
    .addSeparator()
    .addItem('Find + triage + score', 'findTriageAndScore')
    .addToUi();
}

/** Menu items can't take arguments, so each option gets a thin wrapper. */
function menuScoreBatch() {
  toastResult_('Scored', scoreTriagedJobs({ limit: SCORING_DEFAULT_LIMIT }));
}

function menuScoreAll() {
  toastResult_('Scored', scoreTriagedJobs({ limit: 100 }));
}

function menuRescoreAll() {
  const ui = getUiOrNull_();
  if (ui) {
    const ok = ui.alert('Re-score every row?',
      'Scores, strengths, and gaps get refreshed. Your Status, My Notes, and ' +
      'Applied Date are left alone.', ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;
  }
  toastResult_('Re-scored', scoreTriagedJobs({ limit: 100, rescore: true }));
}

/** Toast the summary so menu runs give feedback without opening the logs. */
function toastResult_(verb, summary) {
  const msg = verb + ' ' + (summary.scored || 0) +
              ' — ' + (summary.added || 0) + ' added, ' +
              (summary.updated || 0) + ' updated, ' +
              (summary.errors || 0) + ' error(s), ' +
              (summary.remaining || 0) + ' left';
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Job Score', 8);
  } catch (e) {
    Logger.log('UI Not Available in this Context');
  }
  Logger.log(msg);
}

/** Find links, triage them, score them. Good candidate for a daily trigger. */
function findTriageAndScore() {
  const added = exportToSheet(TRIAGE_SOURCE_SHEET);
  Logger.log('Gmail: %s new link(s).', added);

  const t = triageJobLinks({ limit: TRIAGE_DEFAULT_LIMIT });
  Logger.log('Triage: %s', JSON.stringify(t));

  return scoreTriagedJobs({ limit: SCORING_DEFAULT_LIMIT });
}