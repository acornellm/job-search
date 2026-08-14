/**
 * job-tracker.gs — Google Apps Script (fifth file)
 *
 * A "Job Tracker" sheet that follows an application after you've hit submit.
 * Job Score tracks candidates through the decision of whether to apply;
 * Job Tracker takes over once Status = "Applied" there and follows the
 * application through the post-submission pipeline (recruiter screen,
 * interviews, offer/rejection) — a different lifecycle with different
 * statuses, so it gets its own Status column rather than overloading
 * Job Score's.
 *
 * Depends on job-scoring.gs for: SCORE_SHEET_NAME, headerMap_.
 *
 * Setup:
 *   1. Add this file to the same Apps Script project as the other four.
 *   2. Run setupJobTracker() (or use the Job Tracker menu) to create the
 *      sheet with headers and the Status dropdown.
 *   3. Mark rows Status = "Applied" on Job Score, then run syncAppliedJobs()
 *      (or Job Tracker -> Sync applied jobs from Job Score).
 *
 * Usage:
 *   syncAppliedJobs();   // pull new Applied rows from Job Score
 *
 * Sync only ever appends. Once a job is on Job Tracker, every column is
 * yours — re-running sync never touches an existing row, so Status,
 * interview notes, follow-up dates, etc. are always safe to edit by hand.
 * Dedup key is Job Link (the posting URL), same as Job Score.
 */

// --- Configuration ----------------------------------------------------------

const JOB_TRACKER_SHEET_NAME = 'Job Tracker';

// A post-application pipeline — deliberately separate from Job Score's
// STATUS_OPTIONS (New/Interested/Applying/...), which stops at "Applied".
const JOB_TRACKER_STATUS_OPTIONS = [
  'Applied', 'Recruiter', 'Interviewing', 'Offer', 'Rejected', 'Ghosted', 'Closed',
];
const JOB_TRACKER_STATUS_DEFAULT = 'Applied';

// Beyond the requested columns (Title, Company, Status, Applied Date, Salary,
// Referral Name(s), Job Link, Tailoring Link, Notes), a few more that earn
// their keep once an application is in flight:
//   - Days Since Applied — formula off Applied Date; the fastest way to spot
//     rows that have gone quiet and belong in "Ghosted".
//   - Location, Score, Source — carried over from Job Score so you don't have
//     to flip sheets to remember what the row is or how well it scored.
//   - Recruiter/Contact + Contact Email — who to follow up with; separate
//     from Referral Name(s), which is the person who got you in the door.
//   - Interview Stage — free text ("Phone screen 8/20", "Onsite 8/28") so
//     the Status dropdown can stay coarse-grained.
//   - Next Follow-up / Last Activity — the two dates that actually drive
//     what to do today.
const JOB_TRACKER_HEADERS = [
  'Title', 'Company', 'Status', 'Applied Date', 'Days Since Applied',
  'Location', 'Salary', 'Score', 'Referral Name(s)', 'Recruiter/Contact',
  'Contact Email', 'Interview Stage', 'Next Follow-up', 'Last Activity',
  'Source', 'Job Link', 'Tailoring Link', 'Notes',
];

// --- Sheet setup --------------------------------------------------------

/** Get the Job Tracker sheet, creating it with headers + dropdown if absent. */
function getOrCreateTrackerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(JOB_TRACKER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(JOB_TRACKER_SHEET_NAME);
    sheet.appendRow(JOB_TRACKER_HEADERS);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(3);   // Title / Company / Status stay visible
    sheet.getRange(1, 1, 1, JOB_TRACKER_HEADERS.length).setFontWeight('bold');
    applyTrackerStatusValidation_(sheet);
    Logger.log('Created "%s".', JOB_TRACKER_SHEET_NAME);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(JOB_TRACKER_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, JOB_TRACKER_HEADERS.length).setFontWeight('bold');
    applyTrackerStatusValidation_(sheet);
  }
  return sheet;
}

/** Explicit menu/editor entry point — same effect as the first sync. */
function setupJobTracker() {
  getOrCreateTrackerSheet_();
}

/** Dropdown on the Status column. Invalid values are allowed but flagged. */
function applyTrackerStatusValidation_(sheet) {
  const col = headerMap_(sheet)['Status'];
  if (!col) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(JOB_TRACKER_STATUS_OPTIONS, true)
    .setAllowInvalid(true)
    .setHelpText('Where this application stands, post-submission.')
    .build();

  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

/** Job Link -> row number for everything already on Job Tracker. */
function readTrackerLinks_(sheet) {
  const map = {};
  if (sheet.getLastRow() < 2) return map;

  const col = headerMap_(sheet)['Job Link'];
  const links = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();

  links.forEach(function (r, i) {
    const u = String(r[0] || '').trim();
    if (u) map[u] = i + 2;
  });
  return map;
}

/** Rows on Job Score whose Status is exactly "Applied". */
function readAppliedFromScore_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCORE_SHEET_NAME);
  if (!sheet) throw new Error('No "' + SCORE_SHEET_NAME + '" sheet — run setupScoringSheets() first.');
  if (sheet.getLastRow() < 2) return [];

  const col = headerMap_(sheet);
  const values = sheet.getDataRange().getValues();
  const get = function (row, name) {
    return col[name] ? String(row[col[name] - 1] || '').trim() : '';
  };
  const raw = function (row, name) {
    return col[name] ? row[col[name] - 1] : '';
  };

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (get(row, 'Status') !== 'Applied') continue;

    const url = get(row, 'URL');
    if (!url) continue;

    out.push({
      url: url,
      title: get(row, 'Role Title'),
      company: get(row, 'Company'),
      location: get(row, 'Locations'),
      salary: get(row, 'Salary Range'),
      score: raw(row, 'Score'),
      appliedDate: raw(row, 'Applied Date'),
      source: get(row, 'Source'),
      tailoringLink: get(row, 'Tailoring Doc'),   // added by job-tailoring.gs; '' if absent
    });
  }
  return out;
}

// --- Sync ---------------------------------------------------------------

/**
 * Pull rows marked Applied on Job Score into Job Tracker. Append-only —
 * never touches a row that's already there, so your own tracking (Status,
 * interview notes, follow-up dates) is always safe.
 *
 * @return {Object} { added, skipped, total }
 */
function syncAppliedJobs() {
  const sheet = getOrCreateTrackerSheet_();
  const col = headerMap_(sheet);
  const existing = readTrackerLinks_(sheet);
  const applied = readAppliedFromScore_();

  let added = 0;
  applied.forEach(function (job) {
    if (existing[job.url]) return;   // already tracked

    const appliedDate = job.appliedDate || new Date();
    const row = JOB_TRACKER_HEADERS.map(function (h) {
      switch (h) {
        case 'Title':               return job.title;
        case 'Company':             return job.company;
        case 'Status':              return JOB_TRACKER_STATUS_DEFAULT;
        case 'Applied Date':        return appliedDate;
        case 'Days Since Applied':  return '';   // formula, set below
        case 'Location':            return job.location;
        case 'Salary':              return job.salary;
        case 'Score':               return job.score;
        case 'Last Activity':       return appliedDate;
        case 'Source':              return job.source;
        case 'Job Link':            return job.url;
        case 'Tailoring Link':      return job.tailoringLink;
        default:                    return '';   // user-owned: Referral Name(s),
                                                  // Recruiter/Contact, Contact Email,
                                                  // Interview Stage, Next Follow-up, Notes
      }
    });

    sheet.appendRow(row);
    const rowNum = sheet.getLastRow();
    if (col['Days Since Applied'] && col['Applied Date']) {
      const cell = trackerColLetter_(col['Applied Date']) + rowNum;
      sheet.getRange(rowNum, col['Days Since Applied'])
        .setFormula('=IF(' + cell + '="","",TODAY()-' + cell + ')');
    }
    existing[job.url] = rowNum;
    added++;
  });

  const summary = { added: added, skipped: applied.length - added, total: applied.length };
  Logger.log('Job Tracker sync: %s added, %s already tracked, %s Applied row(s) on Job Score.',
             summary.added, summary.skipped, summary.total);
  return summary;
}

/** Column number (1-based, <= 26) -> spreadsheet letter. */
function trackerColLetter_(col) {
  return String.fromCharCode(64 + col);
}

// --- Menu -----------------------------------------------------------------

/**
 * Job Tracker menu. Called from onOpen() in jobTriage.gs — Apps Script
 * permits only one onOpen per project, so this is a builder, not its own hook.
 */
function jobTrackerMenu_(ui) {
  ui.createMenu('Job Tracker')
    .addItem('Sync applied jobs from Job Score', 'menuSyncTracker')
    .addItem('Create Job Tracker sheet', 'setupJobTracker')
    .addToUi();
}

function menuSyncTracker() {
  const summary = syncAppliedJobs();
  const msg = 'Synced ' + summary.added + ' new — ' + summary.skipped +
              ' already tracked, ' + summary.total + ' total Applied on Job Score';
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Job Tracker', 8);
  } catch (e) {
    Logger.log('UI Not Available in this Context');
  }
  Logger.log(msg);
}
