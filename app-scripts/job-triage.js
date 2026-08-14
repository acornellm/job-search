/**
 * job-triage.gs — Google Apps Script (companion to job-alerts.gs)
 *
 * Reads the "Job Links" sheet produced by exportToSheet(), sends each unseen
 * job URL to the Claude API for structured extraction, and writes the results
 * to a "Job Triage" sheet in the same spreadsheet.
 *
 * Add this as a SECOND file in the same Apps Script project. Apps Script
 * shares one global scope across files, so every identifier here is prefixed
 * or suffixed to avoid colliding with job-alerts.gs.
 *
 * Setup:
 *   1. Add this file to the project alongside job-alerts.gs.
 *   2. Store your API key ONCE — either:
 *        a. Run setApiKey() and paste the key into the prompt, or
 *        b. Project Settings -> Script Properties -> add
 *           ANTHROPIC_API_KEY = sk-ant-...
 *      The key is never written into this file or into the spreadsheet.
 *   3. Run checkApiKey() to confirm it's readable.
 *   4. Run triageJobLinks() — processes 5 links by default.
 *
 * Usage:
 *   triageJobLinks();                         // 5 newest unprocessed links
 *   triageJobLinks({ limit: 20 });            // bigger batch
 *   triageJobLinks({ retryErrors: true });    // redo rows that failed before
 *   triageJobLinks({ limit: 1, dryRun: true });  // log only, no writes
 *
 * Hosts in TRIAGE_EXCLUDED_HOSTS are logged as SKIPPED without an API call.
 */

// --- Configuration --------------------------------------------------------

const TRIAGE_SOURCE_SHEET = 'Job Links';   // written by exportToSheet()
const TRIAGE_SHEET_NAME   = 'Job Triage';  // created on first run

// How many links to send to the API per execution. Apps Script kills a run at
// ~6 minutes on consumer accounts, and each link costs one API call plus a
// page fetch, so keep this modest and run it on a trigger instead.
const TRIAGE_DEFAULT_LIMIT = 5;

// Claude API. Model IDs are pinned snapshots — see
// https://docs.claude.com/en/docs/about-claude/models/overview
const ANTHROPIC_API_URL   = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_VERSION   = '2023-06-01';
const CLAUDE_MODEL        = 'claude-sonnet-5';  // good accuracy/cost for extraction
const CLAUDE_MAX_TOKENS   = 1500;               // fits 15 keywords + 15 skills
const API_KEY_PROPERTY    = 'ANTHROPIC_API_KEY';

// Fallback path: when a direct page fetch comes back empty or blocked, let
// Claude fetch the URL itself using the server-side web fetch tool.
// https://docs.claude.com/en/docs/agents-and-tools/tool-use/web-fetch-tool
const CLAUDE_WEB_FETCH_TOOL = { type: 'web_fetch_20250910', name: 'web_fetch', max_uses: 3 };
const CLAUDE_WEB_FETCH_BETA = 'web-fetch-2025-09-10';

const TRIAGE_MIN_PAGE_CHARS = 600;    // below this, the fetch probably failed
const TRIAGE_MAX_PAGE_CHARS = 40000;  // truncate before sending (~10k tokens)
const TRIAGE_SLEEP_MS       = 1200;   // pause between API calls
const TRIAGE_MAX_RUNTIME_MS = 5 * 60 * 1000;  // bail out before Apps Script does

// Hosts to skip without spending an API call. LinkedIn postings sit behind an
// auth wall, so neither the direct fetch nor the web fetch tool can read them —
// they'd burn tokens and land as errors. Matches the host and its subdomains.
// Paste a resolved URL into the "Manual URL" column on Job Links to bypass
// this for a specific posting — see readJobLinks_().
const TRIAGE_EXCLUDED_HOSTS = ['linkedin.com'];

const TRIAGE_HEADERS = [
  'Processed At', 'Role Title', 'Company', 'Posting Date', 'Locations',
  'Salary Range', 'Top Keywords', 'Technical Skills', 'URL', 'Email Date',
  'Source', 'Email Link', 'Status', 'Notes',
];

const TRIAGE_URL_COL    = 9;   // column I — used for the processed-URL lookup
const TRIAGE_STATUS_COL = 13;  // column M

// --- API key handling -----------------------------------------------------

/**
 * SpreadsheetApp.getUi() throws whenever there's no spreadsheet UI attached —
 * running from the editor on an unbound project, from a time-driven trigger,
 * or from the API. Returns null in those cases instead of blowing up.
 */
function getUiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('UI Not Available in this Context');
    return null;
  }
}

/**
 * Store the API key in Script Properties. Nothing is written to the file or
 * the sheet. Without a UI, this logs the manual instructions instead.
 */
function setApiKey() {
  const ui = getUiOrNull_();
  if (!ui) {
    Logger.log(
      'Cannot prompt for the key here. Add it manually: Project Settings -> ' +
      'Script Properties -> %s = sk-ant-...', API_KEY_PROPERTY);
    return;
  }

  const res = ui.prompt(
    'Anthropic API key',
    'Paste your key (starts with sk-ant-). It is stored in Script Properties, ' +
    'not in the spreadsheet.',
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key) {
    ui.alert('Nothing entered — key not changed.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(API_KEY_PROPERTY, key);
  ui.alert('Key saved. Run checkApiKey() to verify it works.');
}

/** Remove the stored key. */
function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(API_KEY_PROPERTY);
  Logger.log('API key cleared.');
}

/** Cheap round-trip to confirm the key is present and accepted. */
function checkApiKey() {
  const res = callClaude_('Reply with the single word: ok', { maxTokens: 16 });
  Logger.log('API responded: %s', res);
  return res;
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY);
  if (!key) {
    throw new Error(
      'No API key stored. Run setApiKey(), or add ' + API_KEY_PROPERTY +
      ' under Project Settings -> Script Properties.');
  }
  return key;
}

// --- Sheet access ---------------------------------------------------------

/** Get a sheet by name, creating it with headers only if it doesn't exist. */
function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);          // sheet exists but was emptied out
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Read the Job Links sheet into objects, keyed off the header row so column
 * order can shift without breaking this. Duplicate URLs collapse to the first.
 * Rows flagged SKIP are dropped entirely and never returned.
 * @return {Array<Object>} { url, manual, host, date, subject, from, permalink }
 */
function readJobLinks_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName || TRIAGE_SOURCE_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + (sheetName || TRIAGE_SOURCE_SHEET) +
      '" not found. Run exportToSheet() first.');
  }
  ensureManualUrlColumn_(sheet);
  ensureSkipColumn_(sheet);
  if (sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });

  const idx = function (label) { return header.indexOf(label.toLowerCase()); };
  const cUrl    = idx('URL');
  const cManual = idx(MANUAL_URL_HEADER);
  const cSkip   = idx(SKIP_HEADER);
  const cDate   = idx('Date');
  const cHost   = idx('Company/Host');
  const cSubj   = idx('Subject');
  const cFrom   = idx('From');
  const cLink   = idx('Email Link');

  if (cUrl === -1) throw new Error('No "URL" column found in the source sheet.');

  const seen = {};
  const out = [];
  let skipped = 0;

  for (let r = 1; r < values.length; r++) {
    // A row flagged SKIP never becomes a candidate — no fetch, no API call,
    // no Job Triage row. Checked before the Manual URL override so SKIP
    // always wins if someone sets both.
    if (cSkip > -1 && String(values[r][cSkip] || '').trim().toUpperCase() === SKIP_VALUE) {
      skipped++;
      continue;
    }

    const rawUrl = String(values[r][cUrl] || '').trim();
    const manualUrl = cManual > -1 ? String(values[r][cManual] || '').trim() : '';

    // A pasted Manual URL takes priority over whatever the email scrape
    // found — on any host, not just the ones on TRIAGE_EXCLUDED_HOSTS.
    // isExcludedHost_() re-derives the host from this value, so a manual
    // link gets past that check where the excluded original never could.
    const url = manualUrl || rawUrl;
    if (!url || seen[url]) continue;
    seen[url] = true;

    out.push({
      url: url,
      manual: !!manualUrl,
      host: manualUrl ? hostFromUrl_(manualUrl)
          : cHost > -1 ? String(values[r][cHost] || '') : hostFromUrl_(url),
      date: cDate > -1 ? values[r][cDate] : '',
      subject: cSubj > -1 ? String(values[r][cSubj] || '') : '',
      from: cFrom > -1 ? String(values[r][cFrom] || '') : '',
      permalink: cLink > -1 ? String(values[r][cLink] || '') : '',
    });
  }
  if (skipped) Logger.log('%s link(s) excluded by manual "%s" flag.', skipped, SKIP_HEADER);
  return out;
}

/**
 * Map of URLs already in the Job Triage sheet -> { row, status }.
 * This is what keeps a link from being sent to the API twice.
 */
function readProcessedUrls_(sheet) {
  const map = {};
  if (sheet.getLastRow() < 2) return map;

  const rows = sheet.getLastRow() - 1;
  const urls = sheet.getRange(2, TRIAGE_URL_COL, rows, 1).getValues();
  const stats = sheet.getRange(2, TRIAGE_STATUS_COL, rows, 1).getValues();

  for (let i = 0; i < rows; i++) {
    const u = String(urls[i][0] || '').trim();
    if (u) map[u] = { row: i + 2, status: String(stats[i][0] || '').trim().toUpperCase() };
  }
  return map;
}

// --- Page retrieval -------------------------------------------------------

function hostFromUrl_(url) {
  const m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1].toLowerCase().replace(/^www\./, '') : '';
}

/** True when the URL's host is on the no-API-call list. */
function isExcludedHost_(url) {
  const host = hostFromUrl_(url);
  if (!host) return false;
  return TRIAGE_EXCLUDED_HOSTS.some(function (h) {
    h = String(h).toLowerCase().replace(/^www\./, '');
    return host === h || host.endsWith('.' + h);
  });
}

/**
 * Pull readable text off a job posting page. Returns '' when the page is
 * blocked, non-HTML, or JavaScript-rendered with nothing in the initial HTML
 * (common on Workday). The caller falls back to Claude's web fetch tool.
 */
function fetchPageText_(url) {
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        // Some ATS hosts return a stub page to unrecognized clients.
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ' +
                      'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml',
      },
    });
  } catch (e) {
    Logger.log('Fetch threw for %s — %s', url, e.message);
    return '';
  }

  if (res.getResponseCode() >= 400) {
    Logger.log('Fetch %s returned HTTP %s', url, res.getResponseCode());
    return '';
  }

  const type = String(res.getHeaders()['Content-Type'] ||
                      res.getHeaders()['content-type'] || '');
  if (type && type.indexOf('text') === -1 && type.indexOf('html') === -1) return '';

  return htmlToText_(res.getContentText());
}

/** Strip markup down to plain text, keeping paragraph breaks. */
function htmlToText_(html) {
  if (!html) return '';

  const text = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<head[\s\S]*?<\/head>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<\/(p|div|li|tr|h[1-6]|section|article)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<li[^>]*>/gi, '\n• ')
    .replace(/<[^>]+>/g, ' ');

  return decodeEntitiesFull_(text)
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/\n\s*\n\s*\n+/g, '\n\n')
    .trim();
}

/** Named and numeric entity decoding — separate from job-alerts.gs's version. */
function decodeEntitiesFull_(s) {
  return String(s)
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#0?39;|&apos;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&mdash;/gi, '—')
    .replace(/&ndash;/gi, '–')
    .replace(/&#(\d+);/g, function (_, d) {
      const n = parseInt(d, 10);
      return (n > 31 && n < 1114111) ? String.fromCharCode(n) : ' ';
    });
}

// --- Claude API -----------------------------------------------------------

/**
 * One call to the Messages API. Returns the concatenated text blocks.
 *
 * @param {string} prompt
 * @param {Object} [opts]
 * @param {string} [opts.system]
 * @param {number} [opts.maxTokens]
 * @param {boolean} [opts.webFetch]  Attach the server-side web fetch tool.
 * @param {number} [opts.retries=3]
 * @return {string}
 */
function callClaude_(prompt, opts) {
  opts = opts || {};
  const retries = opts.retries === undefined ? 3 : opts.retries;

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: opts.maxTokens || CLAUDE_MAX_TOKENS,
    messages: [{ role: 'user', content: prompt }],
  };
  if (opts.system) payload.system = opts.system;
  if (opts.webFetch) payload.tools = [CLAUDE_WEB_FETCH_TOOL];

  const headers = {
    'x-api-key': getApiKey_(),
    'anthropic-version': ANTHROPIC_VERSION,
  };
  if (opts.webFetch) headers['anthropic-beta'] = CLAUDE_WEB_FETCH_BETA;

  let lastErr = '';

  for (let attempt = 0; attempt <= retries; attempt++) {
    if (attempt > 0) Utilities.sleep(Math.pow(2, attempt) * 1000);  // 2s, 4s, 8s

    const res = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code === 200) return extractText_(JSON.parse(body));

    lastErr = 'HTTP ' + code + ': ' + body.slice(0, 400);

    // 429 and 5xx are worth retrying; 400/401/403 are not.
    if (code !== 429 && code < 500) break;
    Logger.log('Retrying after %s', lastErr);
  }

  throw new Error('Claude API call failed — ' + lastErr);
}

/**
 * Responses can interleave text with server_tool_use and web_fetch_tool_result
 * blocks, so pick blocks out by type rather than by position.
 */
function extractText_(data) {
  if (!data || !data.content) return '';
  return data.content
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('\n')
    .trim();
}

/** Tolerate code fences or stray prose around the JSON object. */
function parseClaudeJson_(text) {
  if (!text) throw new Error('Empty response from the model.');

  let s = text.replace(/```json/gi, '').replace(/```/g, '').trim();

  const first = s.indexOf('{');
  const last = s.lastIndexOf('}');
  if (first === -1 || last === -1 || last <= first) {
    throw new Error('No JSON object in response: ' + s.slice(0, 200));
  }
  return JSON.parse(s.slice(first, last + 1));
}

// --- Extraction prompt ----------------------------------------------------

const TRIAGE_SYSTEM_PROMPT =
  'You extract structured facts from tech job postings. You reply with a single ' +
  'JSON object and nothing else — no markdown fences, no commentary, no ' +
  'preamble. You never invent details that are not in the posting.';

function buildExtractionPrompt_(link, pageText) {
  const schema = [
    'Return exactly this JSON shape:',
    '{',
    '  "role_title": string,            // exact title as posted',
    '  "company": string,               // hiring company, not the job board',
    '  "posting_date": string,          // "YYYY-MM-DD", or "N/A" if not stated',
    '  "locations": [string],           // "City, ST" or "Remote"; [] if none',
    '  "salary_range": string,          // as written, e.g. "$180,000 - $220,000"; "N/A" if absent',
    '  "keywords": [string],            // up to 15 terms characterizing this role',
    '  "technical_skills": [string]     // up to 15 technical skills/tools required',
    '}',
    '',
    'Rules:',
    '- Use the string "N/A" for any scalar field the posting does not state.',
    '- Do not guess salary. If no range appears, "N/A".',
    '- locations: one entry per distinct location. Use "Remote" for fully ' +
      'remote, "City, ST" (2-letter state) for US sites, "City, Country" otherwise.',
    '- keywords: up to 15, drawn from the posting\'s own wording — domain ' +
      'terms, methods, responsibilities, seniority and scope signals. Order ' +
      'them most to least central to the role. Return fewer than 15 rather ' +
      'than padding with generic filler ("team player", "fast-paced").',
    '- technical_skills: up to 15 concrete technologies, platforms, languages, ' +
      'or tools named in the posting. Order them most to least emphasized, ' +
      'and put anything the posting calls required ahead of anything it calls ' +
      'preferred. Do not infer tools the posting does not name. If none are ' +
      'named, return ["N/A"].',
    '- Keep the two lists distinct: technical_skills is named tooling, ' +
      'keywords is everything else.',
    '- If the content is not a job posting (login wall, expired listing, error ' +
      'page), return {"error": "not a job posting"} and nothing else.',
  ].join('\n');

  const context = [
    'Context from the email this link arrived in (use only if the posting ' +
    'itself is silent):',
    '- Source host: ' + (link.host || 'unknown'),
    '- Email subject: ' + (link.subject || 'unknown'),
    '- Email received: ' + formatDate_(link.date),
  ].join('\n');

  if (pageText) {
    return [
      schema, '', context, '',
      'Job posting content from ' + link.url + ':',
      '---', pageText, '---',
    ].join('\n');
  }

  // No local text — hand the URL to Claude and let the web fetch tool get it.
  return [
    schema, '', context, '',
    'Fetch ' + link.url + ' and extract the fields above from that posting.',
  ].join('\n');
}

// --- Main routine ---------------------------------------------------------

/**
 * Process unprocessed job links into the Job Triage sheet.
 *
 * @param {Object} [opts]
 * @param {number}  [opts.limit=5]         Max links to send to the API this run.
 * @param {string}  [opts.sourceSheet]     Defaults to TRIAGE_SOURCE_SHEET.
 * @param {string}  [opts.triageSheet]     Defaults to TRIAGE_SHEET_NAME.
 * @param {boolean} [opts.retryErrors]     Re-run rows whose Status is ERROR,
 *                                         updating them in place.
 * @param {boolean} [opts.useClaudeFetch=true] Fall back to the web fetch tool.
 * @param {boolean} [opts.dryRun]          Log results without writing.
 * @return {Object} { processed, written, skipped, errors, remaining }
 */
function triageJobLinks(opts) {
  opts = opts || {};
  const limit = opts.limit === undefined ? TRIAGE_DEFAULT_LIMIT : opts.limit;
  const useClaudeFetch = opts.useClaudeFetch !== false;
  const started = Date.now();

  const sheet = getOrCreateSheet_(opts.triageSheet || TRIAGE_SHEET_NAME, TRIAGE_HEADERS);
  const processed = readProcessedUrls_(sheet);
  const links = readJobLinks_(opts.sourceSheet);

  // A link is pending if it's absent from Job Triage, or present with an
  // ERROR status and we've been asked to retry.
  const pending = links.filter(function (l) {
    const seen = processed[l.url];
    if (!seen) return true;
    return opts.retryErrors && seen.status === 'ERROR';
  });

  Logger.log('%s link(s) in source, %s pending, limit %s', links.length, pending.length, limit);
  if (!pending.length) return { processed: 0, written: 0, skipped: 0, errors: 0, remaining: 0 };

  const batch = pending.slice(0, limit);
  const newRows = [];
  let errors = 0;

  for (let i = 0; i < batch.length; i++) {
    const link = batch[i];

    if (Date.now() - started > TRIAGE_MAX_RUNTIME_MS) {
      Logger.log('Runtime budget reached — stopping after %s of %s.', i, batch.length);
      break;
    }
    if (i > 0) Utilities.sleep(TRIAGE_SLEEP_MS);

    Logger.log('[%s/%s] %s', i + 1, batch.length, link.url);
    const row = extractOneLink_(link, useClaudeFetch);
    if (row[TRIAGE_HEADERS.indexOf('Status')] === 'ERROR') errors++;

    if (opts.dryRun) {
      Logger.log('DRY RUN — %s', JSON.stringify(row));
      continue;
    }

    const existing = processed[link.url];
    if (existing) {
      sheet.getRange(existing.row, 1, 1, TRIAGE_HEADERS.length).setValues([row]);
    } else {
      newRows.push(row);
    }
  }

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, TRIAGE_HEADERS.length)
      .setValues(newRows);
    sheet.autoResizeColumns(1, 4);
  }

  const summary = {
    processed: batch.length,
    written: newRows.length,
    skipped: links.length - pending.length,
    errors: errors,
    remaining: Math.max(0, pending.length - batch.length),
  };
  Logger.log('Done: %s', JSON.stringify(summary));
  return summary;
}

/** Fetch, extract, and flatten one link into a Job Triage row. */
function extractOneLink_(link, useClaudeFetch) {
  const now = new Date();
  let status = 'OK';
  let note = '';
  let data = {};

  // Excluded hosts never reach the API — no fetch, no tokens, no error row.
  if (isExcludedHost_(link.url)) {
    Logger.log('  excluded host, skipping without an API call');
    return buildTriageRow_(now, link, {}, 'SKIPPED', 'excluded host');
  }

  try {
    let pageText = fetchPageText_(link.url);

    if (pageText.length > TRIAGE_MAX_PAGE_CHARS) {
      pageText = pageText.slice(0, TRIAGE_MAX_PAGE_CHARS);
      note = 'page truncated';
    }

    let usedWebFetch = false;
    if (pageText.length < TRIAGE_MIN_PAGE_CHARS) {
      if (!useClaudeFetch) throw new Error('Page fetch returned too little text.');
      pageText = '';
      usedWebFetch = true;
      note = 'via web fetch tool';
    }

    const raw = callClaude_(buildExtractionPrompt_(link, pageText), {
      system: TRIAGE_SYSTEM_PROMPT,
      webFetch: usedWebFetch,
      maxTokens: CLAUDE_MAX_TOKENS,
    });

    data = parseClaudeJson_(raw);

    if (data.error) {
      status = 'SKIPPED';
      note = String(data.error);
      data = {};
    }
  } catch (e) {
    status = 'ERROR';
    note = e.message.slice(0, 500);
    Logger.log('  failed: %s', note);
  }

  return buildTriageRow_(now, link, data, status, note);
}

/** Assemble one row in TRIAGE_HEADERS order. */
function buildTriageRow_(now, link, data, status, note) {
  data = data || {};
  if (link.manual) note = note ? 'manual URL — ' + note : 'manual URL';
  return [
    now,                                   // Processed At
    scalar_(data.role_title),              // Role Title
    scalar_(data.company),                 // Company
    scalar_(data.posting_date),            // Posting Date
    list_(data.locations),                 // Locations
    scalar_(data.salary_range),            // Salary Range
    list_(data.keywords),                  // Top Keywords
    list_(data.technical_skills),          // Technical Skills
    link.url,                              // URL
    formatDate_(link.date),                // Email Date   (from the link data)
    link.host || hostFromUrl_(link.url),   // Source       (from the link data)
    link.permalink || '',                  // Email Link
    status,                                // Status
    note,                                  // Notes
  ];
}

/** Arrays -> comma-separated string, with N/A for empties. */
function list_(v) {
  if (Array.isArray(v)) {
    const clean = v.map(function (x) { return String(x).trim(); })
                   .filter(function (x) { return x && x.toUpperCase() !== 'N/A'; });
    return clean.length ? clean.join(', ') : 'N/A';
  }
  return scalar_(v);
}

function scalar_(v) {
  if (v === null || v === undefined || v === '') return 'N/A';
  return String(v).trim() || 'N/A';
}

function formatDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d);
}

// --- Runners --------------------------------------------------------------

/** Small batch — safe to attach to a time-driven trigger. */
function triageNextBatch() {
  return triageJobLinks({ limit: TRIAGE_DEFAULT_LIMIT });
}

/**
 * Pull new links from Gmail, then triage a batch. Chains the two files.
 * exportToSheet() comes from job-alerts.gs.
 */
function refreshAndTriage() {
  const added = exportToSheet(TRIAGE_SOURCE_SHEET);
  Logger.log('exportToSheet added %s link(s).', added);
  return triageJobLinks({ limit: TRIAGE_DEFAULT_LIMIT });
}

/** Optional menu. Delete this if you'd rather run from the editor. */
function onOpen() {
  const ui = getUiOrNull_();
  if (!ui) return;

  ui.createMenu('Job Triage')
    .addItem('Set Anthropic API key', 'setApiKey')
    .addItem('Check API key', 'checkApiKey')
    .addSeparator()
    .addItem('Pull new links from Gmail', 'exportToSheet')
    .addItem('Triage next batch', 'triageNextBatch')
    .addItem('Pull + triage', 'refreshAndTriage')
    .addToUi();

  // Apps Script allows exactly one onOpen across the whole project, so the
  // other files expose menu builders instead of defining their own. The
  // guards let any file be absent without breaking the rest.
  if (typeof jobScoreMenu_ === 'function') jobScoreMenu_(ui);
  if (typeof jobTailoringMenu_ === 'function') jobTailoringMenu_(ui);
  if (typeof jobTriggersMenu_ === 'function') jobTriggersMenu_(ui);
}