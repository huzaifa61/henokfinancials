// ---------- CONFIG ----------
const EMAIL_FIELD = "Email Address";
const NAME_FIELD = "Full Name";
const SHEET_NAME = "Form Responses";
const SENT_WELCOME_HEADER = "Welcome Sent";
const SENT_WEEKLY_HEADER = "Weekly Sent";

const CONFIG_SHEET_NAME = "Message_Config"; // Type | Subject | Body  (Type = Welcome or Weekly)

const SPREADSHEET_ID = "1yX0JFvVQZWLfowoiwzQ6JOgeoshO6LBkgFaSV3zVtAc";

const UNSUB_HEADER = "Unsubscribed";
const UNSUB_REASON_HEADER = "Unsubscribe Reason";

// ---- Unsubscribe via a Google Form (no Web App / deployment needed) ----
// Tab where that form's responses land (see setup steps):
const UNSUB_SHEET_NAME = "Unsubscribe_Responses";

// Just the plain Unsubscribe form URL -- no pre-fill needed. The person
// clicks this link, then types their own email in on the form.
// (Get this from the form's Send button -> the link icon -> copy link.)
const UNSUB_FORM_URL = "https://forms.gle/DUjQnvGLuFAo468Z6";
// ----------------------------

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Fires when EITHER the main form or the Unsubscribe form is submitted --
// routes to the right handler based on which sheet the new row landed in.
function onFormSubmit(e) {
  if (!e || !e.range) return;

  const sheetName = e.range.getSheet().getName();

  if (sheetName === SHEET_NAME) {
    handleWelcomeSubmission(e);
  } else if (sheetName === UNSUB_SHEET_NAME) {
    handleUnsubscribeSubmission(e);
  }
}

// Handles a real signup on the main form -- sends the welcome email.
function handleWelcomeSubmission(e) {
  if (!e.namedValues) return;

  const email = (e.namedValues[EMAIL_FIELD] || [""])[0].trim();
  const name = (e.namedValues[NAME_FIELD] || ["there"])[0].trim() || "there";

  if (!email) return;

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);

  const config = getMessageConfig("Welcome");
  const subject = applyPlaceholders(config.subject, { name: name });
  const rawBody = applyPlaceholders(config.body, { name: name });
  const plainMessage = rawBody + buildUnsubscribeFooterPlain(email);
  const htmlMessage = buildEmailHtml(subject, textToHtml(rawBody), buildUnsubscribeFooterHtml(email));

  GmailApp.sendEmail(email, subject, plainMessage, { htmlBody: htmlMessage });

  // Mark welcome email as sent in the sheet
  markSent(email, SENT_WELCOME_HEADER);

  Logger.log("Welcome email sent to " + email);
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

  // Flexible match -- catches "Email", "Email Address", etc.
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

  const mainSheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!mainSheet) {
    throw new Error('Could not find a tab named "' + SHEET_NAME + '".');
  }
  ensureTrackingColumns(mainSheet);
  const mainHeaders = getHeadersRow(mainSheet);
  const mainEmailCol = mainHeaders.indexOf(EMAIL_FIELD) + 1;
  const unsubCol = mainHeaders.indexOf(UNSUB_HEADER) + 1;
  const reasonCol = mainHeaders.indexOf(UNSUB_REASON_HEADER) + 1;

  Logger.log("Looking for email '" + email + "' in main sheet column " + mainEmailCol);

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
    Logger.log("NO MATCH: '" + email + "' was not found among these emails in " + SHEET_NAME + ": " + emails.map(r => r[0]).join(", "));
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

  const unsubSheet = getSpreadsheet().getSheetByName(UNSUB_SHEET_NAME);
  if (!unsubSheet) {
    Logger.log('Could not find a tab named "' + UNSUB_SHEET_NAME + '".');
    return;
  }

  processUnsubscribeRow(unsubSheet, rowToTest);
}

// Fires every Sunday at 9 AM -- sends weekly reminder to everyone (skips unsubscribed)
function sendWeeklyEmails() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureTrackingColumns(sheet);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const emailIndex = headers.indexOf(EMAIL_FIELD);
  const nameIndex = headers.indexOf(NAME_FIELD);
  const weeklyIndex = headers.indexOf(SENT_WEEKLY_HEADER);
  const unsubIndex = headers.indexOf(UNSUB_HEADER);

  if (emailIndex === -1) {
    Logger.log("Email column not found.");
    return;
  }

  const config = getMessageConfig("Weekly");

  for (let i = 1; i < data.length; i++) {
    const email = data[i][emailIndex] ? data[i][emailIndex].toString().trim() : "";
    const name = data[i][nameIndex] ? data[i][nameIndex].toString().trim() : "there";
    const isUnsubscribed = data[i][unsubIndex] === true || data[i][unsubIndex] === "TRUE";

    if (!email || isUnsubscribed) continue;

    const subject = applyPlaceholders(config.subject, { name: name });
    const rawBody = applyPlaceholders(config.body, { name: name });
    const plainMessage = rawBody + buildUnsubscribeFooterPlain(email);
    const htmlMessage = buildEmailHtml(subject, textToHtml(rawBody), buildUnsubscribeFooterHtml(email));

    try {
      GmailApp.sendEmail(email, subject, plainMessage, { htmlBody: htmlMessage });
      sheet.getRange(i + 1, weeklyIndex + 1).setValue("YES");
      Logger.log("Weekly email sent to " + email);
    } catch (err) {
      Logger.log("Failed to send to " + email + ": " + err);
    }
  }
}

// ====================== MESSAGE CONFIG ======================

function getMessageConfig(type) {
  const sheet = getSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
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

function buildUnsubscribeFooterPlain(email) {
  return "\n\n---\nDon't want these emails? Unsubscribe here: " + UNSUB_FORM_URL;
}

function buildUnsubscribeFooterHtml(email) {
  return '<a href="' + UNSUB_FORM_URL + '" style="color:#888888;">Unsubscribe here</a>';
}

// ====================== HTML RENDERING ======================

function linkify(text) {
  const urlPattern = /(https?:\/\/[^\s<]+)/g;
  return text.replace(urlPattern, function(url) {
    return '<a href="' + url + '" style="color:#2A9D8F;">' + url + '</a>';
  });
}

function textToHtml(text) {
  const normalized = text.replace(/\n\s*\n+/g, "\n\n");
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

function buildEmailHtml(subject, bodyHtml, unsubscribeLinkHtml) {
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

// ====================== TRACKING HELPERS ======================

function ensureTrackingColumns(sheet) {
  const headers = getHeadersRow(sheet);
  const toAdd = [UNSUB_HEADER, UNSUB_REASON_HEADER].filter(function(h) { return headers.indexOf(h) === -1; });
  toAdd.forEach(function(h, i) {
    sheet.getRange(1, headers.length + 1 + i).setValue(h);
  });
}

function getHeadersRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// Run this ONCE -- sets up BOTH the form submit trigger and the weekly Sunday trigger.
// Only touches triggers for onFormSubmit / sendWeeklyEmails that YOU own --
// it can't see or affect triggers owned by other users on this project,
// and it won't touch any other automation you might add here later.
function setup() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error("Could not find a tab named \"" + SHEET_NAME + "\".");
  }
  ensureTrackingColumns(sheet);

  const managedHandlers = ["onFormSubmit", "sendWeeklyEmails"];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managedHandlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(getSpreadsheet())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("sendWeeklyEmails")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();

  Logger.log("Setup complete. Form submit trigger (covers both forms) + weekly Sunday trigger installed.");
}

// Helper -- marks a row as sent by email address
function markSent(email, columnHeader) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const emailIndex = headers.indexOf(EMAIL_FIELD);
  let colIndex = headers.indexOf(columnHeader);

  if (colIndex === -1) {
    colIndex = headers.length;
    sheet.getRange(1, colIndex + 1).setValue(columnHeader);
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIndex] && data[i][emailIndex].toString().trim() === email) {
      sheet.getRange(i + 1, colIndex + 1).setValue("YES");
      break;
    }
  }
}