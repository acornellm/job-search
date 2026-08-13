/**
 * job-tailoring.gs — Google Apps Script (fourth file)
 *
 * Takes the jobs you've marked Status = "Applying" on the Job Score sheet and
 * produces a per-job prep packet as a Google Doc:
 *
 *   1. Tailoring brief  — which resume bullets to rewrite and how, which
 *                         keywords are missing, what to reframe, what to leave.
 *   2. Cover letter     — a draft built from the brief, not from scratch.
 *   3. Outreach         — LinkedIn search URLs for the recruiter and the
 *                         likely hiring manager, plus who to target and why.
 *
 * The Doc URL is written back to a "Tailoring Doc" column on Job Score. This
 * runs on a handful of jobs a week, not the whole funnel, so it re-fetches the
 * full posting text rather than reusing the eight extracted triage fields, and
 * it spends real tokens per job. Keep the limit small.
 *
 * Depends on:
 *   job-triage.gs  — getApiKey_, getUiOrNull_, extractText_, parseClaudeJson_,
 *                   fetchPageText_, hostFromUrl_, TRIAGE_MAX_PAGE_CHARS
 *   job-scoring.gs — loadProfileText_, headerMap_, getOrCreateScoreSheet_,
 *                   SCORE_SHEET_NAME
 *
 * Setup:
 *   1. Add this file to the same project.
 *   2. Resume Doc ID must already be set (Job Score menu -> Set resume Doc ID).
 *   3. Optionally run setTailoringFolderId() to pick a Drive folder;
 *      otherwise a "Job Applications" folder is created at the Drive root.
 *   4. Mark a row Status = "Applying", then Job Prep -> Prep applying jobs.
 *
 * Usage:
 *   prepApplyingJobs();                        // up to 3 Applying rows
 *   prepApplyingJobs({ limit: 1 });
 *   prepApplyingJobs({ useWebSearch: true });  // let Claude look up the org
 *   prepOneUrl('https://...');                 // redo a single job
 *   previewTailoringPrompt('https://...');     // see the exact prompts, no API call
 */

// --- Configuration --------------------------------------------------------

const TAILORING_FOLDER_PROPERTY = 'TAILORING_FOLDER_ID';
const TAILORING_FOLDER_NAME     = 'Job Applications';

const TAILORING_STATUS_TRIGGER  = 'Applying';   // Job Score Status that queues a job
const TAILORING_COLUMNS         = ['Tailoring Doc', 'Prepped At'];

// These calls are long-output and low-volume — the opposite economics of
// scoring, so a bigger budget is worth it. Bump to 'claude-opus-5' if the
// briefs feel shallow; at three jobs a week the cost difference is trivial.
const TAILORING_MODEL       = 'claude-sonnet-5';
const TAILORING_MAX_TOKENS  = 4000;
const TAILORING_LIMIT       = 3;
const TAILORING_SLEEP_MS    = 1500;
const TAILORING_MAX_RUNTIME_MS = 5 * 60 * 1000;

// Server-side search, used only when opts.useWebSearch is on.
// https://docs.claude.com/en/docs/agents-and-tools/tool-use/web-search-tool
const CLAUDE_WEB_SEARCH_TOOL = { type: 'web_search_20250305', name: 'web_search', max_uses: 4 };

// --- Drive folder ---------------------------------------------------------

/** Pick the Drive folder for prep Docs. */
function setTailoringFolderId() {
  const ui = getUiOrNull_();
  if (!ui) {
    Logger.log('Cannot prompt here. Add it manually: Project Settings -> ' +
               'Script Properties -> %s = <folder id>', TAILORING_FOLDER_PROPERTY);
    return;
  }
  const res = ui.prompt('Prep Doc folder',
    'Paste a Drive folder ID (the part after /folders/ in the URL). ' +
    'Leave blank to use a "' + TAILORING_FOLDER_NAME + '" folder at the root.',
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return;
  const id = res.getResponseText().trim();

  if (!id) {
    PropertiesService.getScriptProperties().deleteProperty(TAILORING_FOLDER_PROPERTY);
    ui.alert('Cleared — will use the default folder.');
    return;
  }
  DriveApp.getFolderById(id);   // throws early if it isn't reachable
  PropertiesService.getScriptProperties().setProperty(TAILORING_FOLDER_PROPERTY, id);
  ui.alert('Saved.');
}

/** Stored folder, or a "Job Applications" folder created once at the root. */
function getTailoringFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(TAILORING_FOLDER_PROPERTY);

  if (id) {
    try { return DriveApp.getFolderById(id); }
    catch (e) { Logger.log('Stored folder unreachable, falling back: %s', e.message); }
  }

  const existing = DriveApp.getFoldersByName(TAILORING_FOLDER_NAME);
  const folder = existing.hasNext() ? existing.next()
                                    : DriveApp.createFolder(TAILORING_FOLDER_NAME);
  props.setProperty(TAILORING_FOLDER_PROPERTY, folder.getId());
  return folder;
}

// --- Job Score plumbing ---------------------------------------------------

/** Add the tailoring columns to Job Score if they aren't there. */
function ensureTailoringColumns_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  const missing = TAILORING_COLUMNS.filter(function (h) { return header.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, header.length + 1, 1, missing.length).setFontWeight('bold');
    Logger.log('Added column(s): %s', missing.join(', '));
  }
  return headerMap_(sheet);
}

/** Rows on Job Score whose Status is the trigger value. */
function readApplyingJobs_(sheet, col, opts) {
  opts = opts || {};
  const values = sheet.getDataRange().getValues();
  const get = function (row, name) {
    return col[name] ? String(row[col[name] - 1] || '').trim() : '';
  };

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (get(row, 'Status').toLowerCase() !== TAILORING_STATUS_TRIGGER.toLowerCase()) continue;

    // Already prepped rows are skipped unless asked for again.
    if (!opts.redo && get(row, 'Tailoring Doc')) continue;

    const url = get(row, 'URL');
    if (!url) continue;

    out.push({
      rowNum: r + 1,
      url: url,
      role: get(row, 'Role Title'),
      company: get(row, 'Company'),
      locations: get(row, 'Locations'),
      salary: get(row, 'Salary Range'),
      score: get(row, 'Score'),
      verdict: get(row, 'Verdict'),
      strengths: get(row, 'Strengths'),
      gaps: get(row, 'Gaps'),
      source: get(row, 'Source'),
    });
  }
  return out;
}

// --- Job description ------------------------------------------------------

/**
 * Get the full posting text. Falls back to the fields already extracted during
 * triage when the page can't be read — a thinner brief beats no brief.
 */
function loadJobDescription_(job) {
  let text = '';
  try {
    text = fetchPageText_(job.url);
  } catch (e) {
    Logger.log('  page fetch failed: %s', e.message);
  }

  if (text && text.length > TRIAGE_MAX_PAGE_CHARS) {
    text = text.slice(0, TRIAGE_MAX_PAGE_CHARS);
  }
  if (text && text.length > 600) return { text: text, full: true };

  return {
    full: false,
    text: [
      '(Full posting could not be fetched. Working from the extracted summary.)',
      'Role: ' + job.role,
      'Company: ' + job.company,
      'Locations: ' + job.locations,
      'Salary: ' + job.salary,
      'Noted strengths: ' + job.strengths,
      'Noted gaps: ' + job.gaps,
    ].join('\n'),
  };
}

// --- Prompts --------------------------------------------------------------

const BRIEF_SYSTEM = [
  'You advise one candidate on tailoring their materials to a specific job ',
  'posting. You are direct and concrete. You never invent experience the ',
  'candidate does not have, and you never suggest wording that implies it.',
  '',
  'The most useful thing you produce is the rewrite: take an actual bullet ',
  'from the resume, quote it, and give the replacement wording for this ',
  'posting. Vague advice like "emphasize leadership" is worthless — show the ',
  'before and the after.',
  '',
  'Distinguish three kinds of gap, and treat them differently:',
  '  - Reframe: the candidate has the capability under a different name. ',
  '    Give the wording that surfaces it honestly.',
  '  - Address: a real gap worth naming directly in the letter or the screen, ',
  '    paired with the adjacent experience that offsets it.',
  '  - Ignore: a nice-to-have not worth spending space on.',
  '',
  'Write in Markdown with ## section headings and - bullets. No preamble.',
].join('\n');

function buildBriefPrompt_(job, jd, resume) {
  return [
    '# Posting',
    'Role: ' + job.role,
    'Company: ' + job.company,
    'Location: ' + job.locations,
    'Salary: ' + (job.salary || 'not stated'),
    'Fit score from my earlier pass: ' + job.score + ' (' + job.verdict + ')',
    '',
    '## Job description',
    jd.text,
    '',
    '# My current resume',
    resume,
    '',
    '# What to produce',
    'Write a tailoring brief with exactly these sections:',
    '',
    '## Positioning',
    'Two or three sentences: the single strongest angle for this specific ',
    'posting, and the thing about my background most likely to worry them.',
    '',
    '## Resume rewrites',
    'Four to six bullets. Each one quotes my existing bullet, then gives the ',
    'rewritten version, then one line on what changed and why for this posting.',
    '',
    '## Keyword gaps',
    'Terms in the posting that an ATS would look for and my resume lacks. ',
    'For each: whether it is a reframe, an address, or an ignore, and if it is ',
    'a reframe, exactly where in the resume it belongs.',
    '',
    '## Screen call prep',
    'Three questions they are likely to open with given the gaps, and the ',
    'substance of how I should answer each.',
    '',
    '## What not to do',
    'Two or three specific things to avoid claiming or over-emphasizing here.',
  ].join('\n');
}

const LETTER_SYSTEM = [
  'You draft cover letters for one candidate. You write plainly and in their ',
  'voice — no throat-clearing, no "I am writing to express my interest", no ',
  'restating the job description back at them.',
  '',
  'Constraints: under 350 words, four paragraphs at most, no bullet lists. ',
  'Open with the specific reason this role and this company, not a generic ',
  'enthusiasm line. Every claim must be traceable to the resume. If the brief ',
  'flags a gap as one to address, address it in one clause and move on — do ',
  'not apologize for it or dwell.',
  '',
  'Return only the letter body. No header block, no date, no signature line.',
].join('\n');

function buildLetterPrompt_(job, brief, resume) {
  return [
    'Role: ' + job.role + ' at ' + job.company,
    '',
    '# Tailoring brief (already produced for this posting)',
    brief,
    '',
    '# My resume',
    resume,
    '',
    'Write the cover letter.',
  ].join('\n');
}

const OUTREACH_SYSTEM = [
  'You help one candidate figure out who to contact about a specific job ',
  'posting. You work only from public professional information — company ',
  'career pages, public LinkedIn profiles, press releases, engineering blogs.',
  '',
  'Be honest about uncertainty. If you cannot determine who the hiring manager ',
  'is, say so and describe the title to look for instead of guessing at a ',
  'name. A confident wrong name is worse than no name.',
  '',
  'Write in Markdown with ## headings. No preamble.',
].join('\n');

function buildOutreachPrompt_(job, useWebSearch) {
  return [
    'Role: ' + job.role,
    'Company: ' + job.company,
    'Posting: ' + job.url,
    'Location: ' + job.locations,
    '',
    useWebSearch
      ? 'Search for what is publicly known about this company\'s engineering ' +
        'and data organization and its recruiting team.'
      : 'Work from what you know about this company. Do not guess at names.',
    '',
    'Produce:',
    '',
    '## Who to target',
    'The likely hiring manager title for this role given the level and the ' +
    'org shape, and who they probably report to. Name people only if you are ' +
    'confident they currently hold the role.',
    '',
    '## Recruiting team',
    'How this company appears to recruit for senior engineering leadership — ' +
    'in-house talent team, agency, or the hiring manager directly — and what ' +
    'that means for who to reach.',
    '',
    '## Approach',
    'Whether to apply first or reach out first for this specific company, and ' +
    'the one-sentence hook that would work in a connection request.',
  ].join('\n');
}

// --- LinkedIn search URLs (deterministic, no API call) --------------------

/**
 * Build LinkedIn and Google search URLs for the people worth finding. These
 * are constructed strings, not scraped data — you click through and judge the
 * results yourself.
 */
function linkedInSearchUrls_(job) {
  const company = String(job.company || '').replace(/,?\s*(inc|llc|corp|ltd)\.?$/i, '').trim();
  if (!company) return [];

  const enc = encodeURIComponent;
  const peopleSearch = function (keywords) {
    return 'https://www.linkedin.com/search/results/people/?keywords=' + enc(keywords);
  };

  return [
    {
      label: 'Recruiters at ' + company,
      url: peopleSearch(company + ' (recruiter OR "talent acquisition" OR "technical recruiter")'),
    },
    {
      label: 'Engineering leadership at ' + company,
      url: peopleSearch(company + ' ("VP Engineering" OR "Head of Data" OR "Director of Data")'),
    },
    {
      label: 'Data engineering staff at ' + company + ' (warm-intro candidates)',
      url: peopleSearch(company + ' "data engineering"'),
    },
    {
      label: 'Google: profiles naming this role',
      url: 'https://www.google.com/search?q=' +
           enc('site:linkedin.com/in "' + company + '" ("data engineering" OR "talent acquisition")'),
    },
    {
      label: 'Company page — People tab shows mutual connections',
      url: 'https://www.linkedin.com/search/results/companies/?keywords=' + enc(company),
    },
  ];
}

// --- API call -------------------------------------------------------------

/** Messages API call sized for long-form output, optionally with web search. */
function callClaudeTailor_(system, userPrompt, opts) {
  opts = opts || {};
  const retries = opts.retries === undefined ? 2 : opts.retries;

  const payload = {
    model: TAILORING_MODEL,
    max_tokens: opts.maxTokens || TAILORING_MAX_TOKENS,
    system: system,
    messages: [{ role: 'user', content: userPrompt }],
  };
  if (opts.webSearch) payload.tools = [CLAUDE_WEB_SEARCH_TOOL];

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

    if (code === 200) return extractText_(JSON.parse(body));

    lastErr = 'HTTP ' + code + ': ' + body.slice(0, 400);
    if (code !== 429 && code < 500) break;
    Logger.log('Retrying after %s', lastErr);
  }
  throw new Error('Claude API call failed — ' + lastErr);
}

// --- Doc output -----------------------------------------------------------

/** Create the prep Doc and return { id, url }. */
function createTailoringDoc_(job, sections) {
  const name = [
    job.company || 'Unknown company',
    job.role || 'Role',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
  ].join(' — ');

  const doc = DocumentApp.create(name);
  const body = doc.getBody();
  body.clear();

  body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(
    'Fit score ' + (job.score || 'n/a') + ' (' + (job.verdict || 'n/a') + ')  ·  ' +
    (job.locations || 'location n/a') + '  ·  ' + (job.salary || 'salary not stated'));
  body.appendParagraph(job.url).setLinkUrl(job.url);
  body.appendHorizontalRule();

  renderMarkdownToDoc_(body, '# Tailoring brief\n\n' + sections.brief);
  body.appendPageBreak();
  renderMarkdownToDoc_(body, '# Cover letter draft\n\n' + sections.letter);
  body.appendPageBreak();
  renderMarkdownToDoc_(body, '# Outreach\n\n' + sections.outreach);

  if (sections.searchUrls.length) {
    body.appendParagraph('Search links')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    sections.searchUrls.forEach(function (s) {
      const p = body.appendListItem(s.label);
      p.setLinkUrl(s.url);
      p.setGlyphType(DocumentApp.GlyphType.BULLET);
    });
    body.appendParagraph('')
        .appendText('LinkedIn people search needs you to be signed in; these ' +
                    'links open the search, they do not run it for you.')
        .setItalic(true);
  }

  doc.saveAndClose();
  getTailoringFolder_().addFile(DriveApp.getFileById(doc.getId()));
  return { id: doc.getId(), url: doc.getUrl() };
}

/** Minimal Markdown renderer — headings, bullets, and **bold** runs. */
function renderMarkdownToDoc_(body, md) {
  String(md).split('\n').forEach(function (line) {
    const raw = line.replace(/\s+$/, '');

    if (!raw.trim()) { body.appendParagraph(''); return; }

    const h = raw.match(/^(#{1,4})\s+(.*)$/);
    if (h) {
      const levels = [
        DocumentApp.ParagraphHeading.HEADING1,
        DocumentApp.ParagraphHeading.HEADING2,
        DocumentApp.ParagraphHeading.HEADING3,
        DocumentApp.ParagraphHeading.HEADING4,
      ];
      const p = body.appendParagraph(stripBold_(h[2]));
      p.setHeading(levels[h[1].length - 1]);
      return;
    }

    const b = raw.match(/^\s*[-*]\s+(.*)$/);
    if (b) {
      const item = body.appendListItem(stripBold_(b[1]));
      item.setGlyphType(DocumentApp.GlyphType.BULLET);
      applyBold_(item, b[1]);
      return;
    }

    const p = body.appendParagraph(stripBold_(raw));
    applyBold_(p, raw);
  });
}

function stripBold_(s) {
  return String(s).replace(/\*\*/g, '');
}

/** Re-apply bold to the spans that had ** markers before they were stripped. */
function applyBold_(element, original) {
  const text = element.editAsText();
  const parts = String(original).split('**');
  if (parts.length < 3) return;

  let cursor = 0;
  for (let i = 0; i < parts.length; i++) {
    const len = parts[i].length;
    if (i % 2 === 1 && len > 0) {
      try { text.setBold(cursor, cursor + len - 1, true); } catch (e) { /* offsets drifted */ }
    }
    cursor += len;
  }
}

// --- Main routine ---------------------------------------------------------

/**
 * Prep every job marked Applying.
 *
 * @param {Object} [opts]
 * @param {number}  [opts.limit=3]      Max jobs per run — these calls are long.
 * @param {boolean} [opts.redo]         Re-prep rows that already have a Doc.
 * @param {boolean} [opts.useWebSearch] Let Claude search for org/recruiter info.
 * @param {boolean} [opts.skipOutreach] Brief and letter only.
 * @return {Object} { prepped, errors, remaining }
 */
function prepApplyingJobs(opts) {
  opts = opts || {};
  const limit = opts.limit === undefined ? TAILORING_LIMIT : opts.limit;
  const started = Date.now();

  const sheet = getOrCreateScoreSheet_();
  const col = ensureTailoringColumns_(sheet);
  const jobs = readApplyingJobs_(sheet, col, opts);

  Logger.log('%s job(s) marked %s and unprepped, limit %s',
             jobs.length, TAILORING_STATUS_TRIGGER, limit);
  if (!jobs.length) return { prepped: 0, errors: 0, remaining: 0 };

  const resume = loadProfileText_();
  const batch = jobs.slice(0, limit);
  let prepped = 0;
  let errors = 0;

  for (let i = 0; i < batch.length; i++) {
    const job = batch[i];

    if (Date.now() - started > TAILORING_MAX_RUNTIME_MS) {
      Logger.log('Runtime budget reached — stopping after %s of %s.', i, batch.length);
      break;
    }
    if (i > 0) Utilities.sleep(TAILORING_SLEEP_MS);

    Logger.log('[%s/%s] %s at %s', i + 1, batch.length, job.role, job.company);

    try {
      const jd = loadJobDescription_(job);
      if (!jd.full) Logger.log('  posting not fetchable — using the triage summary');

      // Sequential on purpose: the letter is written from the brief, which
      // produces a far better draft than writing both from the posting.
      const brief = callClaudeTailor_(BRIEF_SYSTEM, buildBriefPrompt_(job, jd, resume));
      const letter = callClaudeTailor_(LETTER_SYSTEM,
        buildLetterPrompt_(job, brief, resume), { maxTokens: 1200 });

      const outreach = opts.skipOutreach ? '(skipped)' :
        callClaudeTailor_(OUTREACH_SYSTEM, buildOutreachPrompt_(job, opts.useWebSearch),
          { webSearch: opts.useWebSearch, maxTokens: 1500 });

      const doc = createTailoringDoc_(job, {
        brief: brief,
        letter: letter,
        outreach: outreach,
        searchUrls: linkedInSearchUrls_(job),
      });

      if (col['Tailoring Doc']) {
        sheet.getRange(job.rowNum, col['Tailoring Doc'])
          .setValue(doc.url).setNote(job.role + ' — ' + job.company);
      }
      if (col['Prepped At']) sheet.getRange(job.rowNum, col['Prepped At']).setValue(new Date());

      Logger.log('  -> %s', doc.url);
      prepped++;
    } catch (e) {
      errors++;
      Logger.log('  failed: %s', e.message);
      if (col['Prepped At']) {
        sheet.getRange(job.rowNum, col['Prepped At']).setValue('ERROR: ' + e.message.slice(0, 200));
      }
    }
  }

  const summary = { prepped: prepped, errors: errors,
                    remaining: Math.max(0, jobs.length - batch.length) };
  Logger.log('Done: %s', JSON.stringify(summary));
  return summary;
}

/** Re-prep one job by URL, whatever its Status. */
function prepOneUrl(url) {
  if (!url) { Logger.log('Pass a job URL, e.g. prepOneUrl("https://...")'); return; }

  const sheet = getOrCreateScoreSheet_();
  const col = ensureTailoringColumns_(sheet);
  const values = sheet.getDataRange().getValues();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][col['URL'] - 1] || '').trim() !== url) continue;

    const original = values[r][col['Status'] - 1];
    sheet.getRange(r + 1, col['Status']).setValue(TAILORING_STATUS_TRIGGER);
    try {
      return prepApplyingJobs({ limit: 1, redo: true });
    } finally {
      sheet.getRange(r + 1, col['Status']).setValue(original);   // put Status back
    }
  }
  Logger.log('URL not found on the Job Score sheet.');
}

// --- Inspection -------------------------------------------------------------

/**
 * Print the exact prompts prepApplyingJobs() would send for one job, without
 * spending any API calls. The URL must already be on the Job Score sheet —
 * whatever's in Role Title, Company, Score, Strengths, Gaps, etc. right now
 * is what gets used, same as prepOneUrl() works from.
 *
 * The letter prompt can't be shown exactly: the real run builds it from the
 * brief the brief call generates, so a placeholder stands in for that text.
 */
function previewTailoringPrompt(url) {
  if (!url) { Logger.log('Pass a job URL, e.g. previewTailoringPrompt("https://...")'); return; }

  const sheet = getOrCreateScoreSheet_();
  const col = headerMap_(sheet);
  const values = sheet.getDataRange().getValues();

  const get = function (row, name) {
    return col[name] ? String(row[col[name] - 1] || '').trim() : '';
  };

  let job = null;
  for (let r = 1; r < values.length; r++) {
    if (get(values[r], 'URL') !== url) continue;
    job = {
      url: url,
      role: get(values[r], 'Role Title'),
      company: get(values[r], 'Company'),
      locations: get(values[r], 'Locations'),
      salary: get(values[r], 'Salary Range'),
      score: get(values[r], 'Score'),
      verdict: get(values[r], 'Verdict'),
      strengths: get(values[r], 'Strengths'),
      gaps: get(values[r], 'Gaps'),
      source: get(values[r], 'Source'),
    };
    break;
  }
  if (!job) { Logger.log('URL not found on the Job Score sheet.'); return; }

  const jd = loadJobDescription_(job);
  Logger.log('--- job description (%s, %s chars) ---\n%s',
             jd.full ? 'fetched live' : 'triage summary fallback', jd.text.length, jd.text);

  const resume = loadProfileText_();

  Logger.log('--- brief: system prompt ---\n%s', BRIEF_SYSTEM);
  Logger.log('--- brief: user prompt ---\n%s', buildBriefPrompt_(job, jd, resume));

  Logger.log('--- letter: system prompt ---\n%s', LETTER_SYSTEM);
  Logger.log('--- letter: user prompt (brief text below is a placeholder — the real ' +
             'run uses whatever the brief call actually generates) ---\n%s',
             buildLetterPrompt_(job, '(brief output goes here)', resume));

  Logger.log('--- outreach: system prompt ---\n%s', OUTREACH_SYSTEM);
  Logger.log('--- outreach: user prompt ---\n%s', buildOutreachPrompt_(job, false));
}

// --- Menu -----------------------------------------------------------------

/** Called from onOpen() in job-triage.gs. */
function jobTailoringMenu_(ui) {
  ui.createMenu('Job Prep')
    .addItem('Prep applying jobs', 'menuPrepApplying')
    .addItem('Prep applying jobs (with web search)', 'menuPrepApplyingSearch')
    .addItem('Re-prep everything marked Applying', 'menuRePrepApplying')
    .addSeparator()
    .addItem('Set prep Doc folder', 'setTailoringFolderId')
    .addItem('Open prep folder (see logs)', 'menuOpenPrepFolder')
    .addItem('Preview tailoring prompt (see logs)', 'menuPreviewTailoringPrompt')
    .addToUi();
}

function menuPrepApplying() {
  toastPrep_(prepApplyingJobs({ limit: TAILORING_LIMIT }));
}

function menuPrepApplyingSearch() {
  toastPrep_(prepApplyingJobs({ limit: TAILORING_LIMIT, useWebSearch: true }));
}

function menuRePrepApplying() {
  const ui = getUiOrNull_();
  if (ui) {
    const ok = ui.alert('Re-prep everything marked Applying?',
      'This creates a new Doc per job and overwrites the Tailoring Doc link. ' +
      'The old Docs stay in Drive.', ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;
  }
  toastPrep_(prepApplyingJobs({ limit: TAILORING_LIMIT, redo: true }));
}

function menuOpenPrepFolder() {
  const folder = getTailoringFolder_();
  Logger.log('Prep folder: %s', folder.getUrl());
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(folder.getName(), 'Prep folder (URL in logs)', 8);
  } catch (e) {
    Logger.log('UI Not Available in this Context');
  }
}

/** Menu items can't take arguments, so this prompts for the URL instead. */
function menuPreviewTailoringPrompt() {
  const ui = getUiOrNull_();
  if (!ui) {
    Logger.log('Cannot prompt here. Call previewTailoringPrompt("https://...") directly.');
    return;
  }
  const res = ui.prompt('Preview tailoring prompt',
    'Paste the job URL — it must already be on the Job Score sheet.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const url = res.getResponseText().trim();
  if (!url) { ui.alert('Nothing entered.'); return; }

  previewTailoringPrompt(url);
  ui.alert('Logged. View -> Executions (or Logs) to read the prompts.');
}

function toastPrep_(summary) {
  const msg = 'Prepped ' + summary.prepped + ', ' + summary.errors +
              ' error(s), ' + summary.remaining + ' left';
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Job Prep', 8);
  } catch (e) {
    Logger.log('UI Not Available in this Context');
  }
  Logger.log(msg);
}