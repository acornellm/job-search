# Job Search Pipeline — Google Apps Script

A pipeline that turns Gmail job alerts into a scored, prioritized application
tracker, then follows each application through to an outcome. Runs entirely
in Google Apps Script, bound to one Google Sheet.

## Pipeline

```
Gmail label ──> Job Links ──> Job Triage ──> Job Score ──> prep Docs in Drive
 job-alerts       job-triage    job-scoring    job-tailoring
                                      │
                                      ├─ Status = Applied ──> Job Tracker
                                      │                         job-tracker
                                      v
                                  Analytics (job-analytics)
```

1. **job-alerts.gs** — scans a Gmail label, extracts job-posting URLs from
   message bodies, writes them to the **Job Links** sheet.
2. **job-triage.gs** — fetches each posting, calls the Claude API to extract
   structured fields, writes to the **Job Triage** sheet.
3. **job-scoring.gs** — scores each triaged job against the resume + rubric +
   calibration examples, writes rows to the **Job Score** sheet.
4. **job-tailoring.gs** — for rows the user marks `Applying`, generates a
   tailoring brief, cover letter, and outreach plan into a Google Doc.
5. **job-tracker.gs** — for rows the user marks `Applied` on Job Score,
   copies them into the **Job Tracker** sheet, which follows the application
   through its own post-submission pipeline (recruiter screen, interviews,
   offer/rejection).
6. **job-triggers.gs** — installs the time-based triggers that run stages 1
   and 2 automatically. Run `installTriggers()` once from the editor (or the
   sheet's Job Triggers menu) to create them; re-run it after changing the
   schedule constants at the top of the file. Stages 3, 4, and 5 cost Claude
   API calls or are meant to run on demand, so they're never scheduled —
   only run from the custom menu.
7. **job-analytics.gs** — read-only. Summarizes Job Links, Job Triage, and
   Job Score into an **Analytics** sheet: pipeline funnel, application/
   interview/offer rates, score distributions, top companies/sources, weekly
   activity, and timing. No Claude API calls. Run `refreshAnalytics()` from
   the editor or the sheet's Analytics menu whenever you want it current —
   it isn't kept live by formulas or a trigger.

## Hard constraints — read before editing

- **One global scope.** Apps Script shares a single global namespace across all
  files in a project. A duplicate `const` or `function` name across files is a
  load-time error or a silent override. Before adding a top-level name, grep
  all files. Existing helpers are deliberately prefixed to avoid this
  (`decodeEntities_` in job-alerts.gs vs `decodeEntitiesFull_` in job-triage.gs).
- **Exactly one `onOpen`.** It lives in job-triage.gs and calls
  `jobScoreMenu_(ui)` and `jobTailoringMenu_(ui)`, each behind a
  `typeof x === 'function'` guard. Never add a second `onOpen`.
- **`SpreadsheetApp.getUi()` throws** when there's no attached UI (triggers,
  some editor contexts). Always go through `getUiOrNull_()`, which logs
  "UI Not Available in this Context" and returns null.
- **6-minute execution ceiling** on consumer accounts. Every batch loop checks
  elapsed time and stops early with a partial write rather than losing the run.
- **No npm, no bundler, no local runtime.** Plain ES5-ish JS with `const`/`let`.
  Arrow functions and template literals work in V8 but the codebase uses
  `function` expressions throughout — match that style.

## Verification

There is no local test runner. The available checks:

```bash
node --check <file>.js        # syntax only; catches most edits
clasp push                    # deploy, then run from the Apps Script editor
```

Claude cannot execute Apps Script. After any change, say what the user needs to
run manually and which log lines confirm success.

## Secrets

Stored in Script Properties, never in files or the sheet:

| Property | Purpose |
|---|---|
| `ANTHROPIC_API_KEY` | Claude API auth |
| `RESUME_DOC_ID` | Google Doc holding the master resume |
| `TAILORING_FOLDER_ID` | Drive folder for prep Docs (auto-created if absent) |

Never write a key into a file, a comment, or an example. Never log a key value.

## Sheets

Sheet and column names are the integration contract between files. Code reads
columns **by header name** via `headerMap_()`, never by fixed index — preserve
that when adding columns.

- **Job Links** — `Date, Company/Host, URL, Subject, From, Email Link, Manual URL, Skip`.
  `Manual URL` is user-entry only — paste a resolved posting URL there (any
  host) and `readJobLinks_()` in job-triage.gs uses it instead of the URL
  scraped from the email, bypassing `TRIAGE_EXCLUDED_HOSTS` for that row.
  `Skip` is user-entry only — type `SKIP` to drop a row from triage entirely
  (no fetch, no API call, no Job Triage row); checked before `Manual URL`, so
  `SKIP` wins if both are set. `ensureManualUrlColumn_()` and
  `ensureSkipColumn_()` in job-alerts.gs add these columns if missing, so
  older sheets upgrade automatically.
- **Job Triage** — `Processed At, Role Title, Company, Posting Date, Locations,
  Salary Range, Top Keywords, Technical Skills, URL, Email Date, Source,
  Email Link, Status, Notes`. Status is `OK` / `SKIPPED` / `ERROR`.
  `TRIAGE_URL_COL` and `TRIAGE_STATUS_COL` are positional and must be updated
  together with any reorder.
- **Job Score** — the working tracker. `Status, Score, Verdict, Role Title,
  Company, Locations, Salary Range, Posting Date, Strengths, Gaps, Score Notes,
  My Notes, Applied Date, Source, URL, Email Link, Scored At`, plus
  `Tailoring Doc, Prepped At` added by job-tailoring.gs. Re-sorted highest
  score first after every write — `scoreTriagedJobs()` and `scoreOneUrl()`
  both call `sortScoreSheet()` once they're done writing.
- **Scoring Rubric** — `Type, Criterion, Weight, Notes`. Type is Must-have /
  Nice-to-have / Dealbreaker / **Equivalence**.
- **Calibration** — `Role Title, Company, Key Details, My Score, Reasoning`.
- **Job Tracker** — `Title, Company, Status, Applied Date, Days Since Applied,
  Location, Salary, Score, Referral Name(s), Recruiter/Contact, Contact
  Email, Interview Stage, Next Follow-up, Last Activity, Source, Job Link,
  Tailoring Link, Notes`. Status is its own post-submission pipeline —
  `Applied / Recruiter / Interviewing / Offer / Rejected / Ghosted / Closed` —
  deliberately separate from Job Score's `STATUS_OPTIONS`, which stops at
  `Applied`. `syncAppliedJobs()` in job-tracker.gs copies rows from Job Score
  where `Status = "Applied"` into Job Tracker, deduped by `Job Link` (the
  posting URL, matched against Job Score's `URL`); it only ever appends, so
  every column here is safe to hand-edit once a row exists. `Days Since
  Applied` is a live `TODAY() - Applied Date` formula, not a stored value.
- **Analytics** — generated by `refreshAnalytics()` in job-analytics.gs.
  Fully rewritten on each run (`sheet.clear()` then rewritten section by
  section), so nothing here is user-editable and no other file reads from it.

### User-owned columns

`Status`, `My Notes`, and `Applied Date` on Job Score belong to the user.
`updateScoreRow_()` skips them via `USER_OWNED_COLUMNS`. Re-scoring must never
overwrite them. This is the most important invariant in the project — the
same rule is why `syncAppliedJobs()` on Job Tracker only appends and never
updates an existing row.

## Claude API usage

- Endpoint `https://api.anthropic.com/v1/messages`, header `anthropic-version:
  2023-06-01`, key via `x-api-key`.
- Model IDs are pinned snapshots. Current: `claude-sonnet-5` for both
  extraction and scoring. Check
  https://docs.claude.com/en/docs/about-claude/models/overview before changing.
- **Prompt caching drives the scoring cost model.** The system prompt is two
  blocks: instructions, then resume + rubric + calibration with
  `cache_control: {type: 'ephemeral'}`. The breakpoint must sit on the last
  block that is byte-identical across requests. Anything varying per job goes
  in the user message. A timestamp or row count leaking into the cached prefix
  silently kills every cache hit — there is no error, just full-price input.
  Minimum cacheable prefix is 1,024 tokens on Sonnet 5; `warnIfUncacheable_()`
  flags when the context falls short.
- Server tools in use: `web_fetch_20250910` (beta header
  `web-fetch-2025-09-10`) as the fallback when a page won't fetch locally, and
  `web_search_20250305` for optional company research in job-tailoring.gs.
- Responses interleave `text`, `server_tool_use`, and tool-result blocks.
  Always select blocks **by type**, never by index — `extractText_()` does this.

## Cost and rate discipline

- `TRIAGE_EXCLUDED_HOSTS` (currently `linkedin.com`) short-circuits before any
  fetch or API call — LinkedIn sits behind auth and would burn tokens to fail.
  Excluded URLs are still written as `SKIPPED` so they aren't retried forever.
  To remediate one, paste the resolved posting URL (e.g. the company's own
  careers page) into that row's `Manual URL` cell on Job Links — it becomes
  the row's identity for triage and scoring, and shows up as a new Job Triage
  row noted `manual URL`. The original scraped row is left alone as-is.
- To keep an unwanted posting from being triaged at all — a bad host,
  a role you already know isn't a fit — type `SKIP` in that row's `Skip`
  cell on Job Links. `readJobLinks_()` drops the row before it's ever
  considered pending, so it costs nothing on every future run, same as
  excluded hosts but for any reason you choose.
- Every stage dedupes by URL against its output sheet. Nothing is ever sent to
  the API twice without an explicit `rescore` / `redo` flag.
- Batch limits are deliberately small: 5 for triage and scoring, 3 for
  tailoring. Raise them only with the runtime ceiling in mind.

## Style

- `function` expressions, not arrows. Private helpers end with `_`.
- Comments explain *why*, especially where a line exists to work around an
  Apps Script quirk or an API behavior. Don't strip those.
- Errors are caught per item and recorded in a Status/Notes cell so one bad
  posting never kills a batch.
- `muteHttpExceptions: true` plus explicit code checks on every `UrlFetchApp`
  call. Retry 429 and 5xx with exponential backoff; never retry 4xx.

## Domain context

The user is targeting Director-level Data Engineering roles, coming from a
Director of Software Engineering role at Oracle leading 65+ engineers. His
stack is Oracle-centric (Hadoop, Spark, Cassandra, OCI) against job
descriptions asking for the modern data stack (Snowflake, dbt, Airbyte).

**This gap is the whole reason the rubric exists.** Naive keyword matching
scores him ~62/100 and tells him nothing. The `Equivalence` rows in the rubric
encode what should count as transferable — e.g. Spark/Hadoop platform
ownership earning partial credit toward Snowflake requirements. When editing
scoring prompts, protect the instruction to judge capability transfer rather
than keyword overlap.