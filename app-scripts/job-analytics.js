/**
 * job-analytics.gs — Google Apps Script (reads Job Links, Job Triage, Job
 * Score, and Job Tracker; writes summary stats to an "Analytics" sheet)
 *
 * Pure read + summarize. Never writes back to any source sheet — safe to run
 * any time, as often as you like, no Claude API calls.
 *
 * Usage:
 *   refreshAnalytics();   // recompute and rewrite the Analytics sheet
 *
 * Job Score's STATUS_OPTIONS stops at "Applied" — everything past that
 * (recruiter screens, interviews, offer/rejection) lives on Job Tracker's
 * own Status pipeline instead, so this file reports the pre-application
 * funnel from Job Score and the post-application funnel from Job Tracker
 * separately rather than conflating the two.
 *
 * Depends on job-triage.gs for: formatDate_, TRIAGE_SHEET_NAME,
 * TRIAGE_SOURCE_SHEET.
 * Depends on job-scoring.gs for: headerMap_, SCORE_SHEET_NAME, STATUS_OPTIONS,
 * STATUS_DEFAULT.
 * Depends on job-tracker.gs for: JOB_TRACKER_SHEET_NAME,
 * JOB_TRACKER_STATUS_OPTIONS, JOB_TRACKER_STATUS_DEFAULT.
 * Depends on job-alerts.gs for: SKIP_HEADER, SKIP_VALUE.
 */

const ANALYTICS_SHEET_NAME = 'Analytics';

// Job Tracker statuses that represent a still-open application — used to
// scope "days since applied" and "overdue follow-up" to rows still in play.
const ANALYTICS_TRACKER_ACTIVE_STATUSES = ['Applied', 'Recruiter', 'Interviewing', 'Offer'];

// --- Reading ----------------------------------------------------------

/** Header map + data rows (no header row) for a sheet, or empty if missing. */
function readSheetRows_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return { col: {}, rows: [] };
  const col = headerMap_(sheet);
  const values = sheet.getDataRange().getValues();
  return { col: col, rows: values.slice(1) };
}

function cell_(row, col, name) {
  return col[name] ? row[col[name] - 1] : '';
}

function textCell_(row, col, name) {
  return String(cell_(row, col, name) || '').trim();
}

function numCell_(row, col, name) {
  const v = cell_(row, col, name);
  return typeof v === 'number' && !isNaN(v) ? v : null;
}

function dateCell_(row, col, name) {
  const v = cell_(row, col, name);
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  if (typeof v === 'string' && v.trim()) {
    const d = new Date(v.trim());
    if (!isNaN(d)) return d;
  }
  return null;
}

// --- Small stats helpers ------------------------------------------------

function pct_(n, d) {
  if (!d) return 'n/a';
  return Math.round((n / d) * 1000) / 10 + '%';
}

/** Rounded mean, or 'n/a' for an empty list — never a bare falsy 0. */
function avg_(nums) {
  if (!nums.length) return 'n/a';
  const sum = nums.reduce(function (a, b) { return a + b; }, 0);
  return Math.round((sum / nums.length) * 10) / 10;
}

/** Monday of the week containing d, formatted yyyy-MM-dd. */
function weekStart_(d) {
  const day = d.getDay();               // 0 = Sunday
  const diff = (day === 0 ? -6 : 1) - day;
  const monday = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff);
  return formatDate_(monday);
}

/** Top N {key, count} pairs from a plain counts map, count desc. */
function topN_(counts, n) {
  return Object.keys(counts)
    .map(function (k) { return { key: k, count: counts[k] }; })
    .sort(function (a, b) { return b.count - a.count; })
    .slice(0, n);
}

function bump_(map, key) {
  if (!key) return;
  map[key] = (map[key] || 0) + 1;
}

// --- Data gathering -------------------------------------------------------

/**
 * Everything the report needs, read once.
 * @return {Object}
 */
function gatherAnalyticsData_() {
  const links = readSheetRows_(TRIAGE_SOURCE_SHEET);
  const triage = readSheetRows_(TRIAGE_SHEET_NAME);
  const score = readSheetRows_(SCORE_SHEET_NAME);
  const tracker = readSheetRows_(JOB_TRACKER_SHEET_NAME);

  // Job Links: count everything not flagged SKIP.
  let linksCaptured = 0;
  links.rows.forEach(function (row) {
    if (textCell_(row, links.col, SKIP_HEADER).toUpperCase() === SKIP_VALUE) return;
    linksCaptured++;
  });

  // Job Triage: OK / SKIPPED / ERROR counts.
  const triageStatusCounts = { OK: 0, SKIPPED: 0, ERROR: 0 };
  triage.rows.forEach(function (row) {
    const status = textCell_(row, triage.col, 'Status').toUpperCase();
    if (triageStatusCounts[status] === undefined) triageStatusCounts[status] = 0;
    triageStatusCounts[status]++;
  });

  // Job Score: per-row fields used across every section below.
  const jobs = score.rows.map(function (row) {
    return {
      status: textCell_(row, score.col, 'Status') || STATUS_DEFAULT,
      scoreNum: numCell_(row, score.col, 'Score'),
      verdict: textCell_(row, score.col, 'Verdict'),
      company: textCell_(row, score.col, 'Company'),
      source: textCell_(row, score.col, 'Source'),
      appliedDate: dateCell_(row, score.col, 'Applied Date'),
      scoredAt: dateCell_(row, score.col, 'Scored At'),
    };
  });

  // Job Tracker: per-row fields for the post-application sections below.
  const trackerRows = tracker.rows.map(function (row) {
    return {
      status: textCell_(row, tracker.col, 'Status') || JOB_TRACKER_STATUS_DEFAULT,
      scoreNum: numCell_(row, tracker.col, 'Score'),
      company: textCell_(row, tracker.col, 'Company'),
      source: textCell_(row, tracker.col, 'Source'),
      appliedDate: dateCell_(row, tracker.col, 'Applied Date'),
      nextFollowUp: dateCell_(row, tracker.col, 'Next Follow-up'),
      lastActivity: dateCell_(row, tracker.col, 'Last Activity'),
    };
  });

  return {
    linksCaptured: linksCaptured,
    triageStatusCounts: triageStatusCounts,
    jobs: jobs,
    trackerRows: trackerRows,
  };
}

// --- Report sections --------------------------------------------------

/** @return {Array<Array<*>>} rows to append, each row an array of cell values. */
function buildPipelineFunnelSection_(data) {
  const scored = data.jobs.length;
  const triagedOk = data.triageStatusCounts.OK || 0;

  const statusCounts = {};
  STATUS_OPTIONS.forEach(function (s) { statusCounts[s] = 0; });
  data.jobs.forEach(function (j) {
    if (statusCounts[j.status] === undefined) statusCounts[j.status] = 0;
    statusCounts[j.status]++;
  });

  const rows = [
    ['JOB SCORE — PIPELINE FUNNEL', '', ''],
    ['Stage', 'Count', '% of links captured'],
    ['Links captured (Gmail, not skipped)', data.linksCaptured, ''],
    ['Triaged — OK', triagedOk, pct_(triagedOk, data.linksCaptured)],
    ['Triaged — Skipped', data.triageStatusCounts.SKIPPED || 0, pct_(data.triageStatusCounts.SKIPPED || 0, data.linksCaptured)],
    ['Triaged — Error', data.triageStatusCounts.ERROR || 0, pct_(data.triageStatusCounts.ERROR || 0, data.linksCaptured)],
    ['Scored (Job Score rows)', scored, pct_(scored, data.linksCaptured)],
    ['', '', ''],
  ];

  STATUS_OPTIONS.forEach(function (s) {
    rows.push(['Status: ' + s, statusCounts[s], pct_(statusCounts[s], scored)]);
  });

  return rows;
}

function buildConversionSection_(data) {
  const scored = data.jobs.length;
  const applied = data.jobs.filter(function (j) {
    return j.appliedDate || j.status === 'Applied';
  }).length;
  const triagedOk = data.triageStatusCounts.OK || 0;

  return [
    ['JOB SCORE — CONVERSION RATES', ''],
    ['Metric', 'Value'],
    ['Triage yield (OK / links captured)', pct_(triagedOk, data.linksCaptured)],
    ['Application rate (applied / scored)', pct_(applied, scored)],
    ['', ''],
    ['What happens after you apply — recruiter contact, interviews, offers — ' +
     'is tracked on the Job Tracker sheet; see the sections below.', ''],
  ];
}

function buildScoreSection_(data) {
  const withScore = data.jobs.filter(function (j) { return j.scoreNum !== null; });
  const scoresOf = function (pred) {
    return withScore.filter(pred).map(function (j) { return j.scoreNum; });
  };

  const buckets = [
    ['90-100', function (n) { return n >= 90; }],
    ['80-89', function (n) { return n >= 80 && n < 90; }],
    ['70-79', function (n) { return n >= 70 && n < 80; }],
    ['60-69', function (n) { return n >= 60 && n < 70; }],
    ['Below 60', function (n) { return n < 60; }],
  ];

  const verdictCounts = {};
  data.jobs.forEach(function (j) { bump_(verdictCounts, j.verdict || '(none)'); });

  const rows = [
    ['JOB SCORE — SCORE INSIGHTS', ''],
    ['Metric', 'Value'],
    ['Average score — all scored jobs', avg_(scoresOf(function () { return true; }))],
    ['Average score — applied', avg_(scoresOf(function (j) {
      return j.appliedDate || j.status === 'Applied';
    }))],
    ['Average score — not yet applied (New/Interested/Applying)', avg_(scoresOf(function (j) {
      return ['New', 'Interested', 'Applying'].indexOf(j.status) !== -1;
    }))],
    ['', ''],
    ['Score distribution', 'Count'],
  ];

  buckets.forEach(function (b) {
    const count = withScore.filter(function (j) { return b[1](j.scoreNum); }).length;
    rows.push([b[0], count]);
  });

  rows.push(['', '']);
  rows.push(['Verdict breakdown', 'Count']);
  Object.keys(verdictCounts).sort(function (a, b) {
    return verdictCounts[b] - verdictCounts[a];
  }).forEach(function (v) {
    rows.push([v, verdictCounts[v]]);
  });

  return rows;
}

function buildTopListsSection_(data) {
  const companyCounts = {};
  const companyApplied = {};
  const sourceCounts = {};

  data.jobs.forEach(function (j) {
    if (j.company) {
      bump_(companyCounts, j.company);
      if (j.appliedDate || j.status === 'Applied') bump_(companyApplied, j.company);
    }
    if (j.source) bump_(sourceCounts, j.source);
  });

  const rows = [
    ['JOB SCORE — TOP COMPANIES (by scored postings)', '', ''],
    ['Company', 'Scored', 'Applied'],
  ];
  topN_(companyCounts, 10).forEach(function (c) {
    rows.push([c.key, c.count, companyApplied[c.key] || 0]);
  });

  rows.push(['', '', '']);
  rows.push(['JOB SCORE — TOP SOURCES', '', '']);
  rows.push(['Source', 'Count', '']);
  topN_(sourceCounts, 10).forEach(function (s) {
    rows.push([s.key, s.count, '']);
  });

  return rows;
}

function buildActivitySection_(data) {
  const scoredByWeek = {};
  const appliedByWeek = {};

  data.jobs.forEach(function (j) {
    if (j.scoredAt) bump_(scoredByWeek, weekStart_(j.scoredAt));
    if (j.appliedDate) bump_(appliedByWeek, weekStart_(j.appliedDate));
  });

  const weeks = [];
  const today = new Date();
  for (let i = 7; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - i * 7);
    weeks.push(weekStart_(d));
  }
  const uniqueWeeks = weeks.filter(function (w, i) { return weeks.indexOf(w) === i; });

  const rows = [
    ['JOB SCORE — ACTIVITY (last 8 weeks)', '', ''],
    ['Week of', 'Scored', 'Applied'],
  ];
  uniqueWeeks.forEach(function (w) {
    rows.push([w, scoredByWeek[w] || 0, appliedByWeek[w] || 0]);
  });

  return rows;
}

function buildTimingSection_(data) {
  const latencies = [];
  let lastApplied = null;
  let lastScored = null;

  data.jobs.forEach(function (j) {
    if (j.scoredAt && j.appliedDate) {
      const days = Math.round((j.appliedDate - j.scoredAt) / 86400000);
      if (days >= 0) latencies.push(days);
    }
    if (j.appliedDate && (!lastApplied || j.appliedDate > lastApplied)) lastApplied = j.appliedDate;
    if (j.scoredAt && (!lastScored || j.scoredAt > lastScored)) lastScored = j.scoredAt;
  });

  const daysSinceLastApplication = lastApplied
    ? Math.round((new Date() - lastApplied) / 86400000)
    : 'n/a';

  return [
    ['JOB SCORE — TIMING', ''],
    ['Metric', 'Value'],
    ['Average days from scored to applied', avg_(latencies)],
    ['Most recent Applied Date', lastApplied ? formatDate_(lastApplied) : 'n/a'],
    ['Most recent Scored At', lastScored ? formatDate_(lastScored) : 'n/a'],
    ['Days since last application', daysSinceLastApplication],
  ];
}

// --- Job Tracker sections --------------------------------------------

function buildTrackerFunnelSection_(data) {
  const total = data.trackerRows.length;

  const statusCounts = {};
  JOB_TRACKER_STATUS_OPTIONS.forEach(function (s) { statusCounts[s] = 0; });
  data.trackerRows.forEach(function (t) {
    if (statusCounts[t.status] === undefined) statusCounts[t.status] = 0;
    statusCounts[t.status]++;
  });

  const rows = [
    ['JOB TRACKER — POST-APPLICATION FUNNEL', '', ''],
    ['Stage', 'Count', '% of tracked'],
    ['Tracked applications (Job Tracker rows)', total, ''],
    ['', '', ''],
  ];

  JOB_TRACKER_STATUS_OPTIONS.forEach(function (s) {
    rows.push(['Status: ' + s, statusCounts[s], pct_(statusCounts[s], total)]);
  });

  return rows;
}

function buildTrackerConversionSection_(data) {
  const total = data.trackerRows.length;

  const count = function (statuses) {
    return data.trackerRows.filter(function (t) {
      return statuses.indexOf(t.status) !== -1;
    }).length;
  };

  const recruiterPlus = count(['Recruiter', 'Interviewing', 'Offer']);
  const interviewPlus = count(['Interviewing', 'Offer']);
  const offers = count(['Offer']);
  const rejected = count(['Rejected']);
  const ghosted = count(['Ghosted']);

  return [
    ['JOB TRACKER — CONVERSION RATES', ''],
    ['Metric', 'Value'],
    ['Recruiter contact rate (Recruiter/Interviewing/Offer / tracked)', pct_(recruiterPlus, total)],
    ['Interview rate (Interviewing/Offer / tracked)', pct_(interviewPlus, total)],
    ['Offer rate (Offer / tracked)', pct_(offers, total)],
    ['Offer rate among interviewed (Offer / Interviewing+Offer)', pct_(offers, interviewPlus)],
    ['Rejection rate (Rejected / tracked)', pct_(rejected, total)],
    ['Ghost rate (Ghosted / tracked)', pct_(ghosted, total)],
    ['', ''],
    ['Note: rates use each row\'s current Status, a snapshot rather than ' +
     'full history — a row now at Offer also passed through Recruiter and ' +
     'Interviewing but no longer counts in those rows above.', ''],
  ];
}

function buildTrackerScoreSection_(data) {
  const withScore = data.trackerRows.filter(function (t) { return t.scoreNum !== null; });

  const rows = [
    ['JOB TRACKER — SCORE BY OUTCOME', ''],
    ['Status', 'Average Job Score'],
  ];
  JOB_TRACKER_STATUS_OPTIONS.forEach(function (s) {
    const scores = withScore
      .filter(function (t) { return t.status === s; })
      .map(function (t) { return t.scoreNum; });
    rows.push([s, avg_(scores)]);
  });
  return rows;
}

function buildTrackerTopCompaniesSection_(data) {
  const companyCounts = {};
  const companyRecruiterPlus = {};
  const companyInterviewPlus = {};

  data.trackerRows.forEach(function (t) {
    if (!t.company) return;
    bump_(companyCounts, t.company);
    if (['Recruiter', 'Interviewing', 'Offer'].indexOf(t.status) !== -1) {
      bump_(companyRecruiterPlus, t.company);
    }
    if (['Interviewing', 'Offer'].indexOf(t.status) !== -1) bump_(companyInterviewPlus, t.company);
  });

  const rows = [
    ['JOB TRACKER — TOP COMPANIES', '', '', ''],
    ['Company', 'Tracked', 'Recruiter+', 'Interviewing+'],
  ];
  topN_(companyCounts, 10).forEach(function (c) {
    rows.push([c.key, c.count, companyRecruiterPlus[c.key] || 0, companyInterviewPlus[c.key] || 0]);
  });
  return rows;
}

function buildTrackerTimingSection_(data) {
  const today = new Date();
  const activeDaysSinceApplied = [];
  let overdueFollowUps = 0;
  let mostRecentActivity = null;

  data.trackerRows.forEach(function (t) {
    const isActive = ANALYTICS_TRACKER_ACTIVE_STATUSES.indexOf(t.status) !== -1;
    if (isActive && t.appliedDate) {
      activeDaysSinceApplied.push(Math.round((today - t.appliedDate) / 86400000));
    }
    if (isActive && t.nextFollowUp && t.nextFollowUp < today) overdueFollowUps++;
    if (t.lastActivity && (!mostRecentActivity || t.lastActivity > mostRecentActivity)) {
      mostRecentActivity = t.lastActivity;
    }
  });

  return [
    ['JOB TRACKER — TIMING', ''],
    ['Metric', 'Value'],
    ['Average days since applied (active applications)', avg_(activeDaysSinceApplied)],
    ['Overdue follow-ups (Next Follow-up has passed, still active)', overdueFollowUps],
    ['Most recent activity logged', mostRecentActivity ? formatDate_(mostRecentActivity) : 'n/a'],
  ];
}

// --- Sheet writing -------------------------------------------------------

function getOrCreateAnalyticsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANALYTICS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ANALYTICS_SHEET_NAME);
  return sheet;
}

/** Recompute every section and rewrite the Analytics sheet from scratch. */
function refreshAnalytics() {
  const data = gatherAnalyticsData_();
  const sheet = getOrCreateAnalyticsSheet_();
  sheet.clear();

  const rows = [];
  rows.push(['Job Search Analytics', '', '', '']);
  rows.push(['Last refreshed: ' + Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), '', '', '']);
  rows.push(['', '', '', '']);

  const sections = [
    buildPipelineFunnelSection_(data),
    buildConversionSection_(data),
    buildScoreSection_(data),
    buildTopListsSection_(data),
    buildActivitySection_(data),
    buildTimingSection_(data),
    buildTrackerFunnelSection_(data),
    buildTrackerConversionSection_(data),
    buildTrackerScoreSection_(data),
    buildTrackerTopCompaniesSection_(data),
    buildTrackerTimingSection_(data),
  ];

  const sectionStartRows = [];
  const headerRowIndexes = [];

  sections.forEach(function (section, si) {
    if (si > 0) rows.push(['', '', '', '']);
    sectionStartRows.push(rows.length + 1);   // 1-based sheet row of the title
    section.forEach(function (row, ri) {
      if (ri === 1) headerRowIndexes.push(rows.length + 1);  // the column-header row
      // Pad every row to 4 columns so setValues() gets a rectangular array.
      const padded = row.slice();
      while (padded.length < 4) padded.push('');
      rows.push(padded);
    });
  });

  sheet.getRange(1, 1, rows.length, 4).setValues(rows);

  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(2, 1).setFontStyle('italic');
  sectionStartRows.forEach(function (r) {
    sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12);
  });
  headerRowIndexes.forEach(function (r) {
    sheet.getRange(r, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
  });

  sheet.setColumnWidth(1, 340);
  sheet.setColumnWidths(2, 3, 140);
  sheet.setFrozenRows(0);

  Logger.log('Analytics refreshed: %s scored job(s), %s tracked job(s).',
    data.jobs.length, data.trackerRows.length);
  return { jobs: data.jobs.length, tracked: data.trackerRows.length };
}

// --- Menu -----------------------------------------------------------------

/** Called from onOpen() in job-triage.gs. */
function jobAnalyticsMenu_(ui) {
  ui.createMenu('Analytics')
    .addItem('Refresh analytics', 'menuRefreshAnalytics')
    .addToUi();
}

function menuRefreshAnalytics() {
  const result = refreshAnalytics();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Refreshed — ' + result.jobs + ' scored job(s), ' + result.tracked +
      ' tracked job(s).', 'Analytics', 6);
  } catch (e) {
    Logger.log('UI Not Available in this Context');
  }
}
