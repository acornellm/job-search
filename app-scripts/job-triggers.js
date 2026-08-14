/**
 * Job Triggers — installs and reconciles the time-based triggers for the
 * pipeline. Triggers live in project state, not in file text, so this file
 * is the source of truth for what *should* be installed: edit the schedule
 * below, then run installTriggers() to make actual triggers match it.
 *
 * Scoring and Prep (job-scoring.gs, job-tailoring.gs) hit the Claude API and
 * are deliberately left unscheduled — run them from the sheet's custom menu
 * (jobScoreMenu_ / jobTailoringMenu_), never on a trigger.
 */

// --- Schedule ---------------------------------------------------------

// Job Alerts: Gmail scan only, no API cost, so it can run often.
const ALERTS_TRIGGER_HANDLER = 'exportToSheet';
const ALERTS_TRIGGER_HOURS = 2;   // every 2h -> 12x/day

// Job Triage: calls the Claude API per link, so it stays to one run a day.
const TRIAGE_TRIGGER_HANDLER = 'triageNextBatch';
const TRIAGE_TRIGGER_HOUR = 7;    // local time, per appsscript.json timeZone

// Handler functions this file is allowed to install/remove triggers for.
// Keeping this list explicit means installTriggers() never touches a
// trigger some other script or manual setup created.
const MANAGED_TRIGGER_HANDLERS = [ALERTS_TRIGGER_HANDLER, TRIAGE_TRIGGER_HANDLER];

// --- Install / remove -------------------------------------------------

/**
 * Reconcile installed triggers with the schedule above: delete any managed
 * trigger, then recreate both from scratch. Safe to re-run any time the
 * schedule changes — it never leaves duplicates.
 */
function installTriggers() {
  removeTriggers();

  ScriptApp.newTrigger(ALERTS_TRIGGER_HANDLER)
    .timeBased()
    .everyHours(ALERTS_TRIGGER_HOURS)
    .create();

  ScriptApp.newTrigger(TRIAGE_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(TRIAGE_TRIGGER_HOUR)
    .create();

  Logger.log('Installed triggers: %s every %sh, %s daily at %s:00.',
    ALERTS_TRIGGER_HANDLER, ALERTS_TRIGGER_HOURS, TRIAGE_TRIGGER_HANDLER, TRIAGE_TRIGGER_HOUR);
}

/** Delete every managed trigger without reinstalling. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (MANAGED_TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** Log what's actually installed — handy for verifying after a change. */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.length) {
    Logger.log('No triggers installed.');
    return;
  }
  triggers.forEach(function (t) {
    Logger.log('%s — %s trigger', t.getHandlerFunction(), t.getEventType());
  });
}

// --- Menu ---------------------------------------------------------------

/**
 * Job Triggers menu. Called from onOpen() in job-triage.gs — Apps Script
 * permits only one onOpen per project, so this is a builder rather than its
 * own hook.
 */
function jobTriggersMenu_(ui) {
  ui.createMenu('Job Triggers')
    .addItem('Install triggers', 'installTriggers')
    .addItem('Remove triggers', 'removeTriggers')
    .addItem('List triggers (see logs)', 'listTriggers')
    .addToUi();
}
