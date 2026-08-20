// ---------- CONFIG ----------
const CONFIG_SHEET_NAME = "Message_Content"; // Type | Subject | Body  (Type = Weekly or ThankYou)

// The tab where your MAIN sign-up form's responses land. Needed explicitly
// now that there's a second tab (Unsubscribe_Responses) in this spreadsheet --
// getActiveSheet() is no longer reliable once more than one tab exists.
const MAIN_SHEET_NAME = "Form_Responses";

const UNSUB_HEADER = "Unsubscribed";
const UNSUB_REASON_HEADER = "Unsubscribe Reason";
const SENT_WEEKLY_HEADER = "Weekly Sent"; // stores the date (yyyy-MM-dd) last sent

// ---- Weekly send batching (to stay well under Gmail/Workspace sending limits) ----
const WEEKLY_BATCH_SIZE = 150;   // max emails sent per run
const WEEKLY_START_HOUR = 6;     // don't send before this hour (24h, sheet's timezone)

// ---- Unsubscribe via a Google Form (no Web App / deployment needed) ----
// Tab where THAT form's responses land (link it to this same spreadsheet,
// then rename the new tab to exactly this):
const UNSUB_SHEET_NAME = "Unsubscribe_Responses";

// Just the plain Unsubscribe form URL -- no pre-fill needed (Google Forms
// won't let you pre-fill its built-in "collect email" field anyway). The
// person clicks this link, then types their own email in on the form.
const UNSUB_FORM_URL = "https://forms.gle/peUBk6czn6YL9N5o9";
// ----------------------------

/**
 * ===========================================================
 * SETUP STEPS
 * ===========================================================
 * 1. Fill in MAIN_SHEET_NAME above with your actual response tab name.
 *
 * 2. Add a tab named exactly: Message_Content
 *    Columns: Type | Subject | Body
 *    Two rows: one with Type = Weekly, one with Type = ThankYou
 *    Use {{name}} anywhere you want the recipient's first name inserted.
 *
 * 3. Create (or reuse) an Unsubscribe Google Form:
 *    - Has an Email question and a "why are you unsubscribing" question.
 *    - Link its responses to THIS spreadsheet (Responses tab -> green
 *      Sheets icon -> Select existing spreadsheet).
 *    - Rename the new response tab to exactly: Unsubscribe_Responses
 *    - Copy the form's plain URL (Send button -> link icon) into
 *      UNSUB_FORM_URL above.
 *
 * 4. Run setup() once (function dropdown -> setup -> Run). Installs both
 *    triggers together. Safe to re-run anytime.
 * ===========================================================
 */

// ====================== MESSAGE CONFIG ======================

// Reads Subject/Body for "Weekly" or "ThankYou" from the Message_Content sheet.
function getMessageConfig(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error("Could not find a tab named \"" + CONFIG_SHEET_NAME + "\". Create it with columns Type | Subject | Body.");
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf("Type");
  const subjectCol = headers.indexOf("Subject");
  const bodyCol = headers.indexOf("Body");

  for (let r = 1; r < data.length; r++) {
    if (data[r][typeCol] && data[r][typeCol].toString().trim().toLowerCase() === type.toLowerCase()) {
      return { subject: data[r][subjectCol], body: data[r][bodyCol] };
    }
  }
  throw new Error("No row found in " + CONFIG_SHEET_NAME + " for Type = \"" + type + "\".");
}

function applyPlaceholders(text, values) {
  let result = text;
  Object.keys(values).forEach(function(key) {
    result = result.split("{{" + key + "}}").join(values[key]);
  });
  return result;
}

// ====================== UNSUBSCRIBE ======================

function buildUnsubscribeFooterPlain() {
  return "\n\n---\nDon't want these emails? Unsubscribe here: " + UNSUB_FORM_URL;
}

function buildUnsubscribeFooterHtml() {
  return '<a href="' + UNSUB_FORM_URL + '" style="color:#888888;">Unsubscribe here</a>';
}

// Handles a submission on the Unsubscribe form -- flags that email as
// unsubscribed in the main sheet, and copies over their stated reason.
function handleUnsubscribeSubmission(e) {
  if (!e.namedValues) return;
  processUnsubscribeRow(e.range.getSheet(), e.range.getRow());
}

// Core logic, separated out so it can be tested directly on any row
// without needing a real form-submit event.
function processUnsubscribeRow(unsubSheet, row) {
  const headers = getHeadersRow(unsubSheet);

  const emailColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase().includes("email"));
  if (emailColIdx === -1) {
    Logger.log("Could not find an 'Email' column on the Unsubscribe form's response sheet. Headers found: " + headers.join(" | "));
    return;
  }

  const email = unsubSheet.getRange(row, emailColIdx + 1).getValue().toString().trim();
  if (!email) {
    Logger.log("Row " + row + " has no email value in column " + (emailColIdx + 1) + ".");
    return;
  }

  // The reason is whichever column isn't Timestamp or Email-like
  let reason = "";
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c] ? headers[c].toString().trim().toLowerCase() : "";
    if (h !== "timestamp" && !h.includes("email")) {
      reason = unsubSheet.getRange(row, c + 1).getValue().toString().trim();
      break;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) {
    const names = ss.getSheets().map(s => s.getName()).join(", ");
    throw new Error(
      `MAIN_SHEET_NAME is set to "${MAIN_SHEET_NAME}" but no tab with that exact name exists. ` +
      `Tabs that actually exist in this spreadsheet: ${names}. ` +
      `Update MAIN_SHEET_NAME to match exactly (case-sensitive, including spaces/underscores).`
    );
  }
  ensureTrackingColumns(mainSheet);
  const mainHeaders = getHeadersRow(mainSheet);
  const mainEmailCol = mainHeaders.indexOf("Email Address") + 1;
  const unsubCol = mainHeaders.indexOf(UNSUB_HEADER) + 1;
  const reasonCol = mainHeaders.indexOf(UNSUB_REASON_HEADER) + 1;

  Logger.log("Looking for email '" + email + "' in main sheet column " + mainEmailCol + " (Unsubscribed col=" + unsubCol + ", Reason col=" + reasonCol + ")");

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Main sheet has no data rows.");
    return;
  }

  const emails = mainSheet.getRange(2, mainEmailCol, lastRow - 1, 1).getValues();
  let matchCount = 0;
  for (let i = 0; i < emails.length; i++) {
    const candidate = emails[i][0] ? emails[i][0].toString().trim().toLowerCase() : "";
    if (candidate === email.toLowerCase()) {
      const mainRow = i + 2;
      mainSheet.getRange(mainRow, unsubCol).setValue(true);
      mainSheet.getRange(mainRow, reasonCol).setValue(reason);
      matchCount++;
    }
  }

  if (matchCount > 0) {
    Logger.log("SUCCESS: Marked " + matchCount + " row(s) for " + email + " as unsubscribed. Reason: " + reason);
  } else {
    Logger.log("NO MATCH: '" + email + "' was not found among these emails in " + MAIN_SHEET_NAME + ": " + emails.map(r => r[0]).join(", "));
  }
}

/**
 * TEST FUNCTION — manually processes one specific row from the
 * Unsubscribe_Responses sheet, so you can see exactly what happens
 * (or what error occurs) immediately, without waiting on triggers.
 * Edit rowToTest below, select "testUnsubscribeRow" from the function
 * dropdown, and click Run. Check the log (View > Logs, or Ctrl+Enter).
 */
function testUnsubscribeRow() {
  const rowToTest = 2; // change to the row number in Unsubscribe_Responses you want to test

  const unsubSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNSUB_SHEET_NAME);
  if (!unsubSheet) {
    Logger.log('Could not find a tab named "' + UNSUB_SHEET_NAME + '".');
    return;
  }

  processUnsubscribeRow(unsubSheet, rowToTest);
}

// ====================== TRACKING HELPERS ======================

function ensureTrackingColumns(sheet) {
  const headers = getHeadersRow(sheet);
  const toAdd = [UNSUB_HEADER, UNSUB_REASON_HEADER, SENT_WEEKLY_HEADER].filter(function(h) { return headers.indexOf(h) === -1; });
  toAdd.forEach(function(h, i) {
    sheet.getRange(1, headers.length + 1 + i).setValue(h);
  });
}

function getHeadersRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ====================== HTML RENDERING ======================

function linkify(text) {
  const urlPattern = /(https?:\/\/[^\s<]+)/g;
  return text.replace(urlPattern, function(url) {
    return '<a href="' + url + '" style="color:#2A9D8F;">' + url + '</a>';
  });
}

function textToHtml(text) {
  const normalized = text.replace(/\n\s*\n+/g, "\n\n"); // collapse multiple blank lines into one
  const paragraphs = normalized.split(/\n\n/);

  return paragraphs
    .map(function(p) {
      const trimmed = p.trim();
      if (!trimmed) return "";
      const withBreaks = trimmed.split("\n").join("<br>");
      return '<p style="margin:0 0 16px 0;color:#333333;font-size:15px;line-height:1.6;">' +
        linkify(withBreaks) + "</p>";
    })
    .filter(function(p) { return p !== ""; })
    .join("");
}

function buildEmailHtml(bodyHtml, unsubscribeLinkHtml) {
  return '\
<div style="background-color:#f4f4f4;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">\
  <table role="presentation" width="100%" style="max-width:600px;margin:0 auto;background-color:#ffffff;\
    border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08);">\
    <tr>\
      <td align="center" style="background-color:#1D3557;padding:26px 24px 22px 24px;">\
        <div style="color:#ffffff;font-size:20px;font-weight:bold;">\
          Finance 1-0-1\
        </div>\
      </td>\
    </tr>\
    <tr><td style="height:4px;background-color:#2A9D8F;"></td></tr>\
    <tr>\
      <td style="padding:26px 26px 10px 26px;">\
        ' + bodyHtml + '\
      </td>\
    </tr>\
    <tr>\
      <td style="padding:14px 26px 24px 26px;border-top:1px solid #eeeeee;">\
        <div style="color:#999999;font-size:12px;line-height:1.6;">\
          You are receiving this as part of Finance 1-0-1.<br>\
          Don\'t want these emails? ' + unsubscribeLinkHtml + '.\
        </div>\
      </td>\
    </tr>\
  </table>\
</div>';
}

// ====================== WEEKLY EMAIL ======================

// Weekly email function — MUST be a different name from sendThankYouEmail
// Runs every 15 minutes (see setup()/setupWeeklyTrigger()). Only actually
// sends on Sundays, starting at WEEKLY_START_HOUR. Sends up to
// WEEKLY_BATCH_SIZE emails per run, to people who haven't gotten today's
// weekly email yet -- tracked via the sheet itself, no separate state needed.
function sendWeeklyReminder() {
  const now = new Date();
  if (now.getDay() !== 0) return;                 // Only Sundays
  if (now.getHours() < WEEKLY_START_HOUR) return;  // Don't start before this hour

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  ensureTrackingColumns(sheet);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const emailColIndex = headers.indexOf("Email Address");
  const nameColIndex = headers.indexOf("Full Name");
  const unsubColIndex = headers.indexOf(UNSUB_HEADER);
  const weeklyColIndex = headers.indexOf(SENT_WEEKLY_HEADER);

  if (emailColIndex === -1) {
    Logger.log("Error: 'Email Address' column not found.");
    return;
  }

  const config = getMessageConfig("Weekly");
  const todayStr = formatDate(now); // yyyy-MM-dd, so each Sunday's run is distinct
  let sentThisRun = 0;

  for (let i = 1; i < data.length && sentThisRun < WEEKLY_BATCH_SIZE; i++) {
    const recipientEmail = data[i][emailColIndex]?.toString().trim();
    const name = (nameColIndex !== -1 && data[i][nameColIndex])
      ? data[i][nameColIndex].toString().trim()
      : "there";
    const isUnsubscribed = data[i][unsubColIndex] === true || data[i][unsubColIndex] === "TRUE";
    const alreadySentToday = data[i][weeklyColIndex] === todayStr;

    if (!recipientEmail || isUnsubscribed || alreadySentToday) continue;

    const subject = applyPlaceholders(config.subject, { name: name });
    const rawBody = applyPlaceholders(config.body, { name: name });
    const plainMessage = rawBody + buildUnsubscribeFooterPlain();
    const htmlMessage = buildEmailHtml(textToHtml(rawBody), buildUnsubscribeFooterHtml());

    GmailApp.sendEmail(recipientEmail, subject, plainMessage, { htmlBody: htmlMessage });
    sheet.getRange(i + 1, weeklyColIndex + 1).setValue(todayStr);
    sentThisRun++;
    Logger.log(`Weekly email sent to ${recipientEmail}`);
  }

  Logger.log("Batch complete: sent " + sentThisRun + " email(s) this run (" + todayStr + ").");
}

function formatDate(date) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, "yyyy-MM-dd");
}

// Run this ONCE to create the weekly trigger
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendWeeklyReminder") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendWeeklyReminder")
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("Weekly batch trigger set up (runs every 15 min, only sends Sundays from " + WEEKLY_START_HOUR + "am).");
}

// ====================== THANK YOU EMAIL ======================

// ✅ Sends email immediately when someone submits the form.
// Also routes to the Unsubscribe handler if the submission landed on that
// tab instead -- since one "from spreadsheet" trigger covers every form
// feeding into this spreadsheet, not just the main sign-up form.
function sendThankYouEmail(e) {
  if (!e || !e.namedValues || !e.range) return;

  const sheetName = e.range.getSheet().getName();
  if (sheetName === UNSUB_SHEET_NAME) {
    handleUnsubscribeSubmission(e);
    return;
  }

  const email = (e.namedValues["Email Address"] || [""])[0].trim();
  const name = (e.namedValues["Full Name"] || ["there"])[0].trim() || "there";

  if (!email) return;

  const config = getMessageConfig("ThankYou");
  const subject = applyPlaceholders(config.subject, { name: name });
  const rawBody = applyPlaceholders(config.body, { name: name });
  const plainMessage = rawBody + buildUnsubscribeFooterPlain();
  const htmlMessage = buildEmailHtml(textToHtml(rawBody), buildUnsubscribeFooterHtml());

  GmailApp.sendEmail(email, subject, plainMessage, { htmlBody: htmlMessage });
  Logger.log(`Thank you email sent to ${email}`);
}

// ✅ Run ONCE to set up form submit trigger
function setupFormTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendThankYouEmail") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendThankYouEmail")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log("Form submit trigger set up.");
}

// ====================== COMBINED SETUP ======================

// Run this ONCE to install BOTH triggers together. Safe to re-run anytime --
// it only clears its own previous sendThankYouEmail/sendWeeklyReminder
// triggers first (never touches other automations or other users' triggers,
// which this account can't see or affect anyway).
function setup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    throw new Error("Could not find a tab named \"" + MAIN_SHEET_NAME + "\". Update MAIN_SHEET_NAME to match your actual response tab.");
  }
  ensureTrackingColumns(sheet);

  const managedHandlers = ["sendThankYouEmail", "sendWeeklyReminder"];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managedHandlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendThankYouEmail")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("sendWeeklyReminder")
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("Setup complete. Form submit trigger + 15-minute weekly-batch trigger installed.");
}