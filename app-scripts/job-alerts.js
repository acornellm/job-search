/**
 * job-alerts.gs — Google Apps Script
 *
 * Scans Gmail threads under a label and pulls out links that look like
 * job descriptions (ATS platforms, job boards, careers pages).
 *
 * Setup:
 *   1. https://script.google.com  ->  New project
 *   2. Paste this file in, save.
 *   3. Run `demo()` once and approve the Gmail permission prompt.
 *
 * Usage:
 *   const rows = findJobLinks();                                  // uses JOB_ALERT_LABEL
 *   const rows = findJobLinks(JOB_ALERT_LABEL, { newerThan: '30d', maxThreads: 200 });
 *   const rows = findJobLinks('Job Search/Recruiters');           // one-off override
 */

// The Gmail label to scan. Change this once and every function below follows.
// Nested labels use the full path, e.g. 'Jobs/JobAlerts'.
const JOB_ALERT_LABEL = 'Jobs/JobAlerts';

// Manual-entry column on Job Links. Paste a job posting URL here (any host —
// the company's own careers page, a resolved LinkedIn posting, whatever) and
// job-triage.gs uses it instead of the URL scraped from the email. See
// ensureManualUrlColumn_() and readJobLinks_() in job-triage.gs.
const MANUAL_URL_HEADER = 'Manual URL';

// Manual-entry column on Job Links. Type SKIP in this cell to keep a row out
// of triage entirely — no fetch, no API call, no Job Triage row at all. See
// ensureSkipColumn_() and readJobLinks_() in job-triage.gs.
const SKIP_HEADER = 'Skip';
const SKIP_VALUE = 'SKIP';

// Hosts that almost always mean "this is a job posting."
const JOB_HOSTS = [
  'greenhouse.io', 'boards.greenhouse.io', 'job-boards.greenhouse.io',
  'lever.co', 'jobs.lever.co',
  'ashbyhq.com', 'jobs.ashbyhq.com',
  'myworkdayjobs.com', 'workday.com', 'wd1.myworkdaysite.com',
  'smartrecruiters.com', 'icims.com', 'taleo.net', 'jobvite.com',
  'workable.com', 'breezy.hr', 'bamboohr.com', 'recruitee.com',
  'successfactors.com', 'dayforcehcm.com', 'paylocity.com',
  'oraclecloud.com', 'rippling.com', 'pinpointhq.com', 'teamtailor.com',
  'linkedin.com', 'indeed.com', 'ziprecruiter.com', 'glassdoor.com',
  'dice.com', 'builtin.com', 'wellfound.com', 'angel.co',
  'otta.com', 'hired.com', 'monster.com', 'simplyhired.com', 'hiringcafe.com'
];

// Path fragments that signal a posting even on a company's own domain.
const JOB_PATH_HINTS = ['/job', '/jobs', '/careers', '/career', '/opening', '/position', '/vacanc', '/apply', '/viewjob', '/req', '/job-detail'];

// Click-tracking / redirect wrappers whose real destination sits in a query param.
const REDIRECT_PARAMS = ['url', 'q', 'u', 'target', 'redirect', 'destination', 'link', 'r'];

// Image / asset links — logos, tracking pixels, email banners. Never a posting.
const IMAGE_EXT = /\.(png|jpe?g|gif|svg|webp|bmp|ico|tiff?)$/i;

// Analytics junk to strip so the same posting doesn't show up five times.
const NOISE_PARAMS = /^(utm_|gh_|ga_|mc_|hsa_|_hs|trk$|trkEmail$|refId$|midToken$|lipi$|eBP$|otpToken$|src$|source$|ref$|sourceType$)/i;

/**
 * @param {string} [labelName]  Gmail label. Defaults to JOB_ALERT_LABEL.
 * @param {Object} [opts]
 * @param {string} [opts.query]        Extra Gmail search terms, e.g. 'from:recruiter'.
 * @param {string} [opts.newerThan]    Gmail relative age, e.g. '30d', '6m'.
 * @param {number} [opts.maxThreads=100]
 * @param {string[]} [opts.extraHosts] Additional domains to treat as job links.
 * @param {boolean} [opts.includeAll]  Return every link, not just job-looking ones.
 * @param {boolean} [opts.resolveRedirects] Follow one hop on tracking links (slow).
 * @return {Array<Object>} { url, host, subject, from, date, threadId, permalink }
 */
function findJobLinks(labelName, opts) {
  opts = opts || {};
  const label = labelName || JOB_ALERT_LABEL;
  const max = opts.maxThreads || 100;
  const hosts = JOB_HOSTS.concat(opts.extraHosts || []);

  let q = 'label:"' + label + '"';
  if (opts.newerThan) q += ' newer_than:' + opts.newerThan;
  if (opts.query) q += ' ' + opts.query;

  const threads = GmailApp.search(q, 0, max);
  const seen = {};
  const rows = [];

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      const raw = extractUrls_(msg.getBody()).concat(extractUrls_(msg.getPlainBody()));

      raw.forEach(function (url) {
        let clean = normalizeUrl_(unwrapRedirect_(url));
        if (!clean) return;
        if (isImageUrl_(clean)) return;

        if (opts.resolveRedirects && !isJobLink_(clean, hosts)) {
          clean = normalizeUrl_(followOnce_(clean)) || clean;
          if (isImageUrl_(clean)) return;
        }
        if (!opts.includeAll && !isJobLink_(clean, hosts)) return;
        if (seen[clean]) return;
        seen[clean] = true;

        rows.push({
          url: clean,
          host: hostOf_(clean),
          subject: msg.getSubject(),
          from: msg.getFrom(),
          date: msg.getDate(),
          threadId: thread.getId(),
          permalink: thread.getPermalink(),
        });
      });
    });
  });

  return rows;
}

/** Pull URLs out of an HTML or plain-text body. */
function extractUrls_(body) {
  if (!body) return [];
  const out = [];

  // href="..." first — catches links whose anchor text is just "View job".
  const href = /href\s*=\s*["']([^"']+)["']/gi;
  let m;
  while ((m = href.exec(body)) !== null) out.push(decodeEntities_(m[1]));

  // Bare URLs sitting in the text.
  const bare = /https?:\/\/[^\s"'<>()\[\]]+/gi;
  while ((m = bare.exec(body)) !== null) out.push(decodeEntities_(m[0]));

  return out
    .map(function (u) { return u.replace(/[.,;:!)"'\]]+$/, '').trim(); })
    .filter(function (u) { return /^https?:\/\//i.test(u); });
}

function decodeEntities_(s) {
  return s.replace(/&amp;/g, '&').replace(/&#61;/g, '=')
          .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
          .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}

function hostOf_(url) {
  const m = url.match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1].toLowerCase().replace(/^www\./, '') : '';
}

/**
 * True for links that point at an image file. Checks the path only, so
 * `.../logo.png?width=200` is still caught, while a posting at
 * `.../jobs/1234?ref=image.png` is not falsely dropped.
 */
function isImageUrl_(url) {
  const path = url.split('#')[0].split('?')[0];
  if (IMAGE_EXT.test(path)) return true;
  // Data URIs and CID references occasionally survive the href scrape.
  return /^(data:image|cid:)/i.test(url);
}

function isJobLink_(url, hosts) {
  const host = hostOf_(url);
  const path = url.replace(/^https?:\/\/[^\/]+/i, '').toLowerCase();

  const hostHit = hosts.some(function (h) { return host === h || host.endsWith('.' + h); });
  if (!hostHit) {
    return JOB_PATH_HINTS.some(function (p) { return path.indexOf(p) === 0 || path.indexOf(p + '/') > -1 || path.indexOf(p) > -1; });
  }
  // On LinkedIn/Indeed, only the posting URLs count — not feed or profile links.
  if (host.indexOf('linkedin.com') > -1) return /\/jobs\/view|\/jobs\/collections|currentJobId=/.test(url);
  if (host.indexOf('indeed.com') > -1) return /viewjob|\/jobs|jk=/.test(url);
  if (host.indexOf('dice.com') > -1) return /job-detail|\/jobs|jk=/.test(url);
  if (host.indexOf('hiringcafe.com') > -1) return /job\//.test(url);
  if (host.indexOf('builtin.com') > -1) return /job\/|jk=/.test(url);
  return true;
}

/** Tracking wrappers hide the real URL in a query param — dig it out. */
function unwrapRedirect_(url) {
  for (let i = 0; i < REDIRECT_PARAMS.length; i++) {
    const re = new RegExp('[?&]' + REDIRECT_PARAMS[i] + '=(https?%3A%2F%2F[^&]+|https?:\\/\\/[^&]+)', 'i');
    const m = url.match(re);
    if (m) {
      try { return unwrapRedirect_(decodeURIComponent(m[1])); } catch (e) { return m[1]; }
    }
  }
  return url;
}

/** Strip analytics params and trailing slashes so duplicates collapse. */
function normalizeUrl_(url) {
  if (!url) return null;
  const parts = url.split('#')[0].split('?');
  const base = parts[0].replace(/\/+$/, '');
  if (!parts[1]) return base;

  const kept = parts[1].split('&').filter(function (kv) {
    const k = kv.split('=')[0];
    return k && !NOISE_PARAMS.test(k);
  });
  return kept.length ? base + '?' + kept.join('&') : base;
}

/** Resolve one redirect hop. Costs a fetch per link, so it's opt-in. */
function followOnce_(url) {
  try {
    const res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    const loc = res.getHeaders()['Location'] || res.getHeaders()['location'];
    return loc || url;
  } catch (e) {
    return url;
  }
}

// --- Example runners ------------------------------------------------------

function demo() {
  const rows = findJobLinks(JOB_ALERT_LABEL, { newerThan: '4h', resolveRedirects: true });
  rows.forEach(function (r) { Logger.log(r.date + '  ' + r.host + '  ' + r.url); });
  Logger.log('Found %s links', rows.length);
}

/**
 * Add the "Manual URL" column if it isn't there yet. Sheets created before
 * this feature existed only have the original six headers — this makes the
 * upgrade idempotent instead of requiring a by-hand sheet edit.
 */
function ensureManualUrlColumn_(sheet) {
  const width = sheet.getLastColumn();
  const header = width > 0
    ? sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); })
    : [];
  if (header.indexOf(MANUAL_URL_HEADER) > -1) return;

  const col = width + 1;
  sheet.getRange(1, col).setValue(MANUAL_URL_HEADER).setFontWeight('bold');
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNote('Paste a job posting URL here to override the URL scraped from ' +
             'the email. Takes priority for triage, on any host.');
}

/**
 * Add the "Skip" column if it isn't there yet. Same idempotent pattern as
 * ensureManualUrlColumn_() — safe to call on every export/read.
 */
function ensureSkipColumn_(sheet) {
  const width = sheet.getLastColumn();
  const header = width > 0
    ? sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); })
    : [];
  if (header.indexOf(SKIP_HEADER) > -1) return;

  const col = width + 1;
  sheet.getRange(1, col).setValue(SKIP_HEADER).setFontWeight('bold');
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNote('Type ' + SKIP_VALUE + ' to keep this link out of triage entirely ' +
             '— no fetch, no API call, no Job Triage row.');
}

/**
 * Dump results into a sheet — handy alongside an application tracker.
 * Reuses the sheet if it's already there, and only appends URLs that
 * aren't logged yet, so this is safe to run on a recurring trigger.
 */
function exportToSheet(sheetName, labelName, opts) {
  sheetName = sheetName || 'Job Links';
  const HEADERS = ['Date', 'Company/Host', 'URL', 'Subject', 'From', 'Email Link'];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);   // sheet exists but was emptied out
    sheet.setFrozenRows(1);
  }
  ensureManualUrlColumn_(sheet);
  ensureSkipColumn_(sheet);

  // Column C holds the URL — build a set of what's already there.
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0]) existing[String(r[0]).trim()] = true; });
  }

  const rows = findJobLinks(labelName || JOB_ALERT_LABEL, opts || { newerThan: '1d', resolveRedirects: true })
    .filter(function (r) { return !existing[r.url]; })
    .map(function (r) { return [r.date, r.host, r.url, r.subject, r.from, r.permalink]; });

  if (!rows.length) {
    Logger.log('No new links for "%s".', sheetName);
    return 0;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.autoResizeColumns(1, HEADERS.length);
  Logger.log('Added %s new link(s) to "%s".', rows.length, sheetName);
  return rows.length;
}