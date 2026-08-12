/**
 * Job Search Automation System
 * Google Apps Script for automating job search, email processing, and application tracking
 * * @version 2.9
 * @description Final, complete, and stable version with "best guess" parsing.
 */


// =============================================
// CONFIGURATION SETTINGS
// =============================================
const SETTINGS = {
  SHEETS: {
    JOB_EMAILS: "Job Emails", JOB_ALERTS: "Job Alerts", APPLICATIONS: "Applications",
    SETTINGS: "Settings", DASHBOARD: "Dashboard"
  },
  LABELS: {
    JOB_ALERTS: "Jobs/JobAlerts", APPLICATION: "Jobs/JobSearch/Application", INTERVIEW: "Jobs/JobSearch/Interview",
    OFFER: "Jobs/JobSearch/Offer", REJECTION: "Jobs/JobSearch/Rejection"
  },
  AI: {
    MODEL: "claude-sonnet-5",
    SYSTEM: "You are an expert executive tech leadership recruiter who screens 300 resumes daily hiring for my next level role.  You are an expert resume writer and editor and an ATS specialist",
    MAX_TOKENS: 1024,
    TEMPERATURE: 0.5,
    REQUEST_DELAY: 3000, // milliseconds between API calls
    MAX_JOBS_TO_SCORE_PER_RUN: 5 // NEW LINE: Limits jobs scored per run to prevent timeouts
  },
  MIN_DATE: new Date("2026-06-01T00:00:00Z"),
  MAX_EMAILS_PER_RUN: 50
};


// =============================================
// MAIN WORKFLOW & TRIGGER FUNCTIONS
// =============================================
function runCompleteWorkflow() {
  console.log("🚀 Starting complete workflow...");
  try {
    importTaggedEmails();
    importJobAlertEmails();
    populateAndUpdateApplications();
    scoreJobDescriptions();
    updateDashboardMetrics();
    sortAllSheets();
    console.log("✅ Workflow completed successfully!");
  } catch (error) {
    console.error("❌ Workflow error:", error);
  }
}


function createAutomatedTriggers() {
  deleteTriggers();
  console.log("⚙️ Creating automation triggers...");
  ScriptApp.newTrigger('runCompleteWorkflow').timeBased().everyDays(1).atHour(9).create();
  console.log("✓ Triggers created successfully!");
}


function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}


// =============================================
// EMAIL & APPLICATION PROCESSING
// =============================================
function importTaggedEmails() {
  console.log("Step 1: Importing tagged emails...");
  const sheet = getOrCreateSheet(SETTINGS.SHEETS.JOB_EMAILS);
  sheet.clear();
  initializeEmailSheet(sheet);


  Object.values(SETTINGS.LABELS).filter(l => l !== SETTINGS.LABELS.JOB_ALERTS).forEach(labelName => {
    processEmailLabel(labelName, sheet);
  });
}


function processEmailLabel(labelName, sheet) {
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    console.log(`Label not found: ${labelName}`);
    return;
  };
 
  const threads = label.getThreads(0, SETTINGS.MAX_EMAILS_PER_RUN);
  console.log(`Found ${threads.length} threads for label: ${labelName}`);


  threads.forEach(thread => {
    const msg = thread.getMessages()[0];
    if (msg.getDate() < SETTINGS.MIN_DATE) return;
    sheet.appendRow([
      msg.getSubject(), msg.getFrom(), msg.getDate(), labelName.split('/').pop(),
      guessEmailSource(msg.getFrom()), `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`,
      msg.getPlainBody().slice(0, 2000), thread.getId()
    ]);
  });
}


function populateAndUpdateApplications() {
    console.log("Step 3: Auto-populating and updating applications...");
    const jobEmailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS.SHEETS.JOB_EMAILS);
    const applicationsSheet = getOrCreateSheet(SETTINGS.SHEETS.APPLICATIONS);
    if (!jobEmailsSheet || jobEmailsSheet.getLastRow() < 2) {
        console.log("Job Emails sheet is empty. Nothing to process.");
        return;
    }


    const applications = getApplicationsMap(applicationsSheet);
    const emailData = jobEmailsSheet.getDataRange().getValues();
    let updatedCount = 0, newCount = 0;


    for (let i = 1; i < emailData.length; i++) {
        const [subject, sender, dateReceived, label] = emailData[i];
        const snippet = emailData[i][6];
       
        const appDetails = extractApplicationFromEmail(subject, sender, snippet);
       
        console.log(`Processing: "${subject}"`);
        console.log(`  > Parsed Title: ${appDetails.jobTitle}`);
        console.log(`  > Parsed Company: ${appDetails.company}`);


        const key = createApplicationKey(appDetails.jobTitle, appDetails.company);
        const existingApp = applications.get(key);
        const newStatus = getStatusFromLabel(label);
        if (!newStatus) continue;


        if (existingApp) {
            if (isNewStatusAdvanced(newStatus, existingApp.status)) {
                applicationsSheet.getRange(existingApp.row, 3).setValue(newStatus);
                applications.set(key, { ...existingApp, status: newStatus });
                updatedCount++;
                console.log(`  > Action: Updated status to "${newStatus}"`);
            }
        } else {
            addApplicationEntry(applicationsSheet, {
                jobTitle: appDetails.jobTitle,
                company: appDetails.company,
                status: newStatus,
                dateApplied: dateReceived,
                source: guessEmailSource(sender),
            });
            applications.set(key, { row: applicationsSheet.getLastRow(), status: newStatus });
            newCount++;
            console.log(`  > Action: Added new application.`);
        }
    }
    if (newCount > 0 || updatedCount > 0) sortApplicationsSheet();
    console.log(`Application processing complete. New: ${newCount}, Updated: ${updatedCount}.`);
}


function getApplicationsMap(sheet) {
    const applications = new Map();
    if (sheet.getLastRow() < 2) return applications;
    const data = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
    data.forEach((row, index) => {
        if (row[0] && row[1]) {
            applications.set(createApplicationKey(row[0], row[1]), { row: index + 2, status: row[2] });
        }
    });
    return applications;
}


function getStatusFromLabel(label) {
    const statusMap = { "Interview": "Interview Scheduled", "Rejection": "Rejected", "Offer": "Offer Received", "Application": "Applied" };
    return statusMap[label] || null;
}


function isNewStatusAdvanced(newStatus, currentStatus) {
    const statusOrder = ["Applied", "Interview Scheduled", "Offer Received"];
    if (newStatus === "Rejected") return true;
    return statusOrder.indexOf(newStatus) > statusOrder.indexOf(currentStatus);
}


function extractApplicationFromEmail(subject, sender, snippet) {
  let jobTitle = extractJobTitle(subject, snippet);
  let company = extractCompany(subject, sender, snippet);
 
  if (!jobTitle) {
      jobTitle = subject.includes('|') ? subject.split('|')[1].trim() : subject;
  }
  if (!company) {
      company = sender.split('<')[0].trim();
  }


  return { jobTitle, company };
}


function extractJobTitle(subject, snippet) {
    const patterns = [
        /application for (?:our|the)\s+(.+?)\s+(?:position|role)/i,
        /your application for\s+(.+?)(?:\s+at|$)/i,
        /interest in (?:the)?\s+(.+?)\s+position/i,
        /position of\s+(.+?)(?:\s+at|\.|\,)/i,
        /interview with .* for the (.+?) role/i,
        /Invitation:.*?:\s*(.+?)\s*@/i,
    ];
    for (const p of patterns) {
        let match = snippet.match(p) || subject.match(p);
        if (match && match[1]) return match[1].replace(/REQ\d+/, "").replace(/\(.*\)/, "").trim();
    }
    return null;
}


function extractCompany(subject, sender, snippet) {
    const patterns = [
        /thank you for applying to\s+(.+)/i,
        /^(.+?)\s*[|–-]\s*Thank you/i,
        /position at\s+(.+?)\./i,
    ];
     for (const p of patterns) {
        let match = subject.match(p) || snippet.match(p);
        if (match && match[1]) return match[1].trim();
    }


    let nameMatch = sender.match(/^([^<@]+)/);
    if (nameMatch) {
        let companyName = nameMatch[1].replace(/The Recruiting team at|Recruiting|Careers|Talent|HR|Recruitment/gi, "").trim();
        if (companyName.length > 2) return companyName;
    }


    let domainMatch = sender.match(/@([\w-]+)\./);
    if (domainMatch) {
        const genericDomains = ['gmail', 'yahoo', 'outlook', 'lever', 'greenhouse', 'google', 'workday', 'icims', 'smartrecruiters'];
        if (!genericDomains.includes(domainMatch[1].toLowerCase())) {
            return domainMatch[1].charAt(0).toUpperCase() + domainMatch[1].slice(1);
        }
    }
    return null;
}


function importJobAlertEmails() {
  console.log("Step 2: Importing job alert emails...");
  const sheet = getOrCreateSheet(SETTINGS.SHEETS.JOB_ALERTS);
  const label = GmailApp.getUserLabelByName(SETTINGS.LABELS.JOB_ALERTS);
  if (!label) {
    console.log("Job Alerts label not found.");
    return;
  }
 
  const existingLinks = getExistingJobLinks(sheet);
  const threads = label.getThreads(0, SETTINGS.MAX_EMAILS_PER_RUN);
  console.log(`Found ${threads.length} threads for Job Alerts.`);
  let newJobsCount = 0;


  threads.forEach(thread => {
    const message = thread.getMessages()[0];
    const jobsArray = parseGenericJobAlert(message); // Using the new generic parser
   
    jobsArray.forEach(jobData => {
      if (jobData.link && !existingLinks.has(jobData.link)) {
        sheet.appendRow([
          jobData.subject,
          jobData.snippet,
          jobData.link,
          jobData.date,
          "", // Fit Score
          "", // Resume Suggestions
          "", // JD Keywords
          ""  // Cover Letter Blurb
        ]);
        existingLinks.add(jobData.link);
        newJobsCount++;
      }
    });
  });


  if (newJobsCount > 0) {
    sortJobAlertsSheet();
    console.log(`Job alert import complete. Added ${newJobsCount} new jobs.`);
  } else {
    console.log("Job alert import complete. No new jobs found.");
  }
}


/**
 * NEW: A more generic parser for any email containing job links.
 */
function parseGenericJobAlert(message) {
    const body = message.getBody(); // Use HTML body to find links
    const plainBody = message.getPlainBody();
    const date = message.getDate();
    const jobs = [];


    // General regex to find any HTTP link within an <a> tag
    const linkRegex = /<a\s+(?:[^>]*?\s+)?href="([^"]*)"/gi;
    let match;


    while ((match = linkRegex.exec(body)) !== null) {
        let link = match[1];


        // Decode Google redirect links if present
        if (link.includes("google.com/url")) {
            const urlMatch = link.match(/url=([^&]+)/);
            if (urlMatch) link = decodeURIComponent(urlMatch[1]);
        }
       
        // Filter for links that look like job postings
        const isJobLink = /\/jobs?\/|\/careers\/|linkedin\.com\/jobs\/view/i.test(link);


        if (isJobLink) {
            // Find the link's position in the plain text to get context
            const plainLink = link.split('?')[0]; // Use a cleaner version for searching plain text
            const linkIndex = plainBody.indexOf(plainLink);
           
            let jobTitle = `Job Alert: ${message.getSubject()}`;
            let snippet = `Link: ${link}`;


            if (linkIndex !== -1) {
                // Try to get text immediately before the link as the title
                const precedingText = plainBody.substring(0, linkIndex).trim();
                const lastLineBreak = precedingText.lastIndexOf('\n');
                const potentialTitle = precedingText.substring(lastLineBreak + 1).trim();


                if (potentialTitle && potentialTitle.length > 5) {
                    jobTitle = potentialTitle;
                    snippet = `${potentialTitle}\n${link}`;
                }
            }


            jobs.push({
                subject: jobTitle,
                snippet: snippet,
                link: cleanJobLink(link),
                date: date
            });
        }
    }
    return jobs;
}


function parseJobsFromGoogleAlert(message) {
    const body = message.getBody();
    const jobs = [];
    const linkRegex = /<a href="(https:\/\/[^"]*google.com\/url[^"]*)"[^>]*>([\s\S]*?)<\/a>/g;
    let match;
    while ((match = linkRegex.exec(body)) !== null && jobs.length < 15) {
        const urlMatch = match[1].match(/url=([^&]+)/);
        if (!urlMatch) continue;
        const jobLink = decodeURIComponent(urlMatch[1]);
        const jobTitle = XmlService.parse(`<p>${match[2]}</p>`).getRootElement().getValue().replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
        if (isValidJobListing(jobTitle)) {
            jobs.push({ subject: jobTitle, snippet: jobTitle, link: cleanJobLink(jobLink), date: message.getDate() });
        }
    }
    return jobs;
}


// =============================================
// AI SCORING & DASHBOARD
// =============================================
function scoreJobDescriptions() {
  console.log("Step 4: Scoring job descriptions with AI...");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS.SHEETS.JOB_ALERTS);
  if (!sheet || sheet.getLastRow() < 2) {
    console.log("No jobs to score.");
    return;
  }


  const data = sheet.getDataRange().getValues();
  const resumeText = getUserResumeText();
  if (!resumeText || !resumeText.includes("specializes in")) {
    console.log("Resume text not found in Settings. Skipping scoring.");
    return;
  }
 
  let scoredInThisRun = 0; // Counter for this specific run


  for (let i = 1; i < data.length; i++) {
    // Exit if we've hit the batch limit for this run
    if (scoredInThisRun >= SETTINGS.AI.MAX_JOBS_TO_SCORE_PER_RUN) {
      console.log(`Scored batch of ${SETTINGS.AI.MAX_JOBS_TO_SCORE_PER_RUN}. More will be scored on the next run.`);
      break;
    }


    const [jobTitle, snippet, , , score] = data[i];
    if (score || !snippet) continue; // Skip if already scored or no snippet


    try {
      console.log(`Scoring: "${jobTitle}"`);
      const aiResult = scoreJobWithAI(jobTitle, snippet, resumeText);
      if (aiResult) {
          updateJobScore(sheet, i + 1, aiResult);
          scoredInThisRun++;
      }
      Utilities.sleep(SETTINGS.AI.REQUEST_DELAY); // Pause between API calls
    } catch (error) {
      console.error(`Error scoring job "${jobTitle}":`, error);
    }
  }


  if (scoredInThisRun > 0) {
    sortJobAlertsByScore();
  }
  console.log(`AI scoring complete for this run. Scored: ${scoredInThisRun} new jobs.`);
}


function scoreJobWithAI(jobTitle, snippet, resumeText) {
  const prompt = `Analyze this job:\nJOB: ${jobTitle}\nDESCRIPTION: ${snippet}\nCANDIDATE: ${resumeText}\nReturn a valid JSON with keys: "Role Title", "Company", "Salary and Compensation", "Brief role summary", "Fit Score" (1-100), "Priority Level" (High/Medium/Low)`;
  const response = callClaudeAIChat(prompt);
  if (!response) return null;
  console.log(response.length);
  console.log(response);
  try { return JSON.parse(response.replace(/```json\n|```/g, "").trim()); }
  catch (e) { return null; }
}


function updateJobScore(sheet, row, result) {
  sheet.getRange(row, 5).setValue(result["Fit Score"] || "");
  sheet.getRange(row, 6).setValue(result["Resume Suggestions"] || "");
  sheet.getRange(row, 7).setValue(result["JD Keywords"] || "");
  sheet.getRange(row, 8).setValue(result["Cover Letter Blurb"] || "");
}


function updateDashboardMetrics() {
    console.log("Step 5: Updating dashboard metrics...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const applicationsSheet = ss.getSheetByName(SETTINGS.SHEETS.APPLICATIONS);
    let dashboardSheet = getOrCreateSheet(SETTINGS.SHEETS.DASHBOARD);
    dashboardSheet.clear();
    initializeDashboardSheet(dashboardSheet);
    const numJobAlerts = (ss.getSheetByName(SETTINGS.SHEETS.JOB_ALERTS) || {getLastRow: () => 1}).getLastRow() -1;
    const numJobEmails = (ss.getSheetByName(SETTINGS.SHEETS.JOB_EMAILS) || {getLastRow: () => 1}).getLastRow() - 1;
    const numApplications = (applicationsSheet || {getLastRow: () => 1}).getLastRow() - 1;
    const statusBreakdown = getApplicationStatusBreakdown(applicationsSheet);
    const conversionRates = calculateConversionRates(statusBreakdown);
    let currentRow = 2;
    addSectionHeader(dashboardSheet, currentRow++, "OVERVIEW");
    currentRow = addMetricsSection(dashboardSheet, currentRow, [
        ["Total Job Alerts", numJobAlerts], ["Total Job Emails", numJobEmails],
        ["Total Applications", numApplications], ["Last Updated", new Date().toLocaleString()]
    ]);
    currentRow++;
    addSectionHeader(dashboardSheet, currentRow++, "APPLICATION STATUS BREAKDOWN");
    const statusMetrics = Object.entries(statusBreakdown).map(([status, count]) => [status, count]);
    currentRow = addMetricsSection(dashboardSheet, currentRow, statusMetrics);
    currentRow++;
    addSectionHeader(dashboardSheet, currentRow++, "CONVERSION RATES");
    addMetricsSection(dashboardSheet, currentRow, [
        ["Application → Response Rate", conversionRates.applicationToResponse],
        ["Application → Interview Rate", conversionRates.interviewConversion],
        ["Interview → Offer Rate", conversionRates.interviewToOffer],
        ["Overall Success Rate", conversionRates.overallSuccess]
    ]);
    dashboardSheet.autoResizeColumns(1, 2);
}


function addSectionHeader(sheet, row, headerText) {
  sheet.getRange(row, 1, 1, 2).merge().setValue(headerText).setFontWeight("bold").setBackground("#e8f4fd");
}


function addMetricsSection(sheet, startRow, metrics) {
    if (metrics.length > 0) {
        sheet.getRange(startRow, 1, metrics.length, 2).setValues(metrics);
        return startRow + metrics.length;
    }
    return startRow;
}


function getApplicationStatusBreakdown(applicationsSheet) {
  const breakdown = {};
  if (!applicationsSheet || applicationsSheet.getLastRow() <= 1) return breakdown;
  const statuses = applicationsSheet.getRange(2, 3, applicationsSheet.getLastRow() - 1, 1).getValues().flat();
  statuses.forEach(status => {
    breakdown[status] = (breakdown[status] || 0) + 1;
  });
  return breakdown;
}


function calculateConversionRates(statusBreakdown) {
  const safePercent = (num, den) => den > 0 ? `${((num / den) * 100).toFixed(1)}%` : "N/A";
  const totalApps = Object.values(statusBreakdown).reduce((a, b) => a + b, 0);
  const totalResponses = totalApps - (statusBreakdown["Applied"] || 0) - (statusBreakdown["No Response"] || 0) - (statusBreakdown["Withdrawn"] || 0);
  const totalInterviews = (statusBreakdown["Interview Scheduled"] || 0);
  const totalOffers = (statusBreakdown["Offer Received"] || 0);
  return {
    applicationToResponse: safePercent(totalResponses, totalApps),
    interviewConversion: safePercent(totalInterviews, totalApps),
    interviewToOffer: safePercent(totalOffers, totalInterviews),
    overallSuccess: safePercent(totalOffers, totalApps)
  };
}


// =============================================
// SHEET INITIALIZATION & SETUP
// =============================================
function initialSetup() {
  console.log("Starting initial setup...");
  Object.values(SETTINGS.SHEETS).forEach(sheetName => {
    const sheet = getOrCreateSheet(sheetName);
    if (sheetName === SETTINGS.SHEETS.JOB_EMAILS) initializeEmailSheet(sheet);
    else if (sheetName === SETTINGS.SHEETS.APPLICATIONS) initializeApplicationsSheet(sheet);
    else if (sheetName === SETTINGS.SHEETS.JOB_ALERTS) initializeJobAlertsSheet(sheet);
    else if (sheetName === SETTINGS.SHEETS.SETTINGS) initializeSettingsSheet(sheet);
    else if (sheetName === SETTINGS.SHEETS.DASHBOARD) initializeDashboardSheet(sheet);
  });
   console.log("✅ Initial setup complete! Headers are fixed.");
}


function initializeDashboardSheet(sheet) {
  if (sheet.getLastRow() < 1) {
    sheet.clear();
    sheet.appendRow(["Metric", "Value"]).getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff");
    sheet.setColumnWidth(1, 300).setColumnWidth(2, 120);
  }
}


function initializeEmailSheet(sheet) {
  sheet.clear();
  const headers = ["Subject", "Sender", "Date", "Label", "Source", "Link", "Snippet", "Thread ID"];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f0f0f0");
}


function initializeJobAlertsSheet(sheet) {
  sheet.clear();
  const headers = ["Subject", "Snippet", "Link", "Date", "Fit Score", "Resume Suggestions", "JD Keywords", "Cover Letter Blurb"];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f0f0f0");
}


function initializeApplicationsSheet(sheet) {
    sheet.clear();
    const headers = ["Job Title", "Company", "Status", "Location", "Date Applied", "Source", "Next Step Date", "Link", "Salary", "Notes"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f0f0f0");
    const statusOptions = ["Applied", "Interview Scheduled", "Offer Received", "Rejected"];
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build();
    sheet.getRange("C2:C").setDataValidation(rule);
}


function initializeSettingsSheet(sheet) {
  if (sheet.getLastRow() < 1) {
    sheet.clear();
    const defaultSettings = [
      ["Setting", "Value"], ["Resume Text", "I am a professional specializing in..."], ["Target Job Titles", "Software Engineer, Product Manager"],
      ["Target Locations", "San Francisco, Remote"], ["Exclusion Keywords", "intern, contract"]
    ];
    sheet.getRange(1, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }
}


// =============================================
// API & SECURITY FUNCTIONS
// =============================================
function callOpenAIChat(prompt) {
  const apiKey = getApiKey();
  if (!apiKey) { console.error("OpenAI API key not set."); return null; }
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = { model: SETTINGS.AI.MODEL, messages: [{ role: "user", content: prompt }], temperature: SETTINGS.AI.TEMPERATURE };
  const options = { method: "post", contentType: "application/json", headers: { Authorization: `Bearer ${apiKey}` }, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  if (result.error) { console.error("OpenAI API Error:", result.error.message); return null; }
  return result.choices[0].message.content;
}

function callClaudeAIChat(prompt) {
  var result;
  const apiKey = getApiKey();
  if (!apiKey) { console.error("Claude API key not set."); return null; }
  const url = "https://api.anthropic.com/v1/messages";
  const payload = { model: SETTINGS.AI.MODEL, system: "You are an expert executive tech leadership recruiter.", max_tokens: SETTINGS.AI.MAX_TOKENS, messages: [{ role: "user", content: prompt }] };
  console.log(JSON.stringify(payload));
  const options = { method: "post", contentType: "application/json", headers: { "x-api-key": `${apiKey}`, "anthropic-version": "2023-06-01"}, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  console.log(data);
  if (data.error) { console.error("Claude API Error:", data.error.message); return null; }

  for (const block of data.content) {
    if (block.type === 'text') {
      result = block.text;
    }
  }
  // Extract and log Claude's text response
  console.log("Claude's Response:", result);
  return result;
}


function setApiKey() {
  const apiKey = Browser.inputBox("API Key Setup", "Enter your OpenAI API key:", Browser.Buttons.OK_CANCEL);
  if (apiKey && apiKey !== "cancel") {
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey.trim());
  }
}


function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}


// =============================================
// UTILITY & HELPER FUNCTIONS
// =============================================
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}


function getProcessedThreads(sheet) {
  if (sheet.getLastRow() < 2) return new Set();
  const ids = sheet.getRange("H2:H" + sheet.getLastRow()).getValues().flat().filter(id => id);
  return new Set(ids);
}


function getExistingJobLinks(sheet) {
  if (sheet.getLastRow() < 2) return new Set();
  const links = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat().filter(link => link);
  return new Set(links);
}


function getUserResumeText() {
  console.log(getSettingValue("Resume Text", "N/A"));
    return getSettingValue("Resume Text", "N/A");
}


function getSettingValue(settingName, defaultValue = "") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS.SHEETS.SETTINGS);
    if (!sheet) return defaultValue;
    const data = sheet.getDataRange().getValues();
    const row = data.find(r => r[0] === settingName);
    return row && row[1] ? row[1] : defaultValue;
}


function addApplicationEntry(sheet, entry) {
  const newRow = [
    entry.jobTitle, entry.company, entry.status, "", entry.dateApplied,
    entry.source, "", "", "", "Auto-populated"
  ];
  sheet.appendRow(newRow);
  sheet.getRange(sheet.getLastRow(), 5).setNumberFormat("MM/dd/yyyy");
}


function createApplicationKey(jobTitle, company) {
  return `${(jobTitle || "").toLowerCase().trim()}_${(company || "").toLowerCase().trim()}`;
}


function guessEmailSource(fromEmailString) {
  const from = fromEmailString.toLowerCase();
  if (from.includes("greenhouse")) return "Greenhouse";
  if (from.includes("lever")) return "Lever";
  if (from.includes("workday")) return "Workday";
  if (from.includes("smartrecruiters")) return "SmartRecruiters";
  if (from.includes("linkedin")) return "LinkedIn";
  return "Direct Application";
}


function isValidJobListing(jobTitle) { return jobTitle && jobTitle.length > 5; }
function cleanJobLink(link) { return link ? link.split('&utm_')[0] : ""; }


// =============================================
// SORTING & MAINTENANCE
// =============================================
function sortAllSheets() {
  console.log("Step 6: Sorting all sheets...");
  sortSheetByColumn(SETTINGS.SHEETS.JOB_ALERTS, 4, false);
  sortSheetByColumn(SETTINGS.SHEETS.JOB_EMAILS, 3, false);
  sortSheetByColumn(SETTINGS.SHEETS.APPLICATIONS, 5, false);
}


function sortSheetByColumn(sheetName, column, ascending) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column, ascending });
}


function sortJobAlertsSheet() { sortSheetByColumn(SETTINGS.SHEETS.JOB_ALERTS, 4, false); }
function sortJobEmailsSheet() { sortSheetByColumn(SETTINGS.SHEETS.JOB_EMAILS, 3, false); }
function sortApplicationsSheet() { sortSheetByColumn(SETTINGS.SHEETS.APPLICATIONS, 5, false); }
function sortJobAlertsByScore() { sortSheetByColumn(SETTINGS.SHEETS.JOB_ALERTS, 5, false); }


function weeklyMaintenance() {
  console.log("Starting weekly maintenance...");
  cleanupOldData(90);
  updateDashboardMetrics();
  sortAllSheets();
}


function cleanupOldData(daysToKeep) {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
  [SETTINGS.SHEETS.JOB_EMAILS, SETTINGS.SHEETS.JOB_ALERTS].forEach((sheetName, index) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const dateColumn = index === 0 ? 3 : 4;
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = data.reduce((acc, row, i) => {
      if (i > 0 && row[dateColumn - 1] instanceof Date && row[dateColumn - 1] < cutoffDate) {
        acc.push(i + 1);
      }
      return acc;
    }, []);
    rowsToDelete.reverse().forEach(rowNum => sheet.deleteRow(rowNum));
  });
}


// SCRIPT END



