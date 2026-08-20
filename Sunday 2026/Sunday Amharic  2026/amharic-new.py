// ---------- CONFIG ----------
const EMAIL_FIELD = "Email Address";
const NAME_FIELD = "Full Name";
const SHEET_NAME = "Form Responses";

const SENT_WELCOME_HEADER = "Welcome Sent";
const SENT_WEEKLY_HEADER = "Weekly Sent";

const CONFIG_SHEET_NAME = "Message_Config";
// Columns:
// Type | Subject | Body
//
// Type values:
// Welcome
// Weekly

const SPREADSHEET_ID =
  "1yX0JFvVQZWLfowoiwzQ6JOgeoshO6LBkgFaSV3zVtAc";

const UNSUB_HEADER = "Unsubscribed";
const UNSUB_REASON_HEADER = "Unsubscribe Reason";

// ---------- UNSUBSCRIBE ----------
const UNSUB_SHEET_NAME =
  "Unsubscribe_Responses";

const UNSUB_FORM_URL =
  "https://forms.gle/DUjQnvGLuFAo468Z6";

// ---------- AMAZON SES ----------
/*
 * Google Apps Script
 *        ↓
 * API Gateway
 *        ↓
 * Lambda
 *        ↓
 * Amazon SES
 */

const SES_API_URL =
  "https://vs5njf4n2f.execute-api.us-east-2.amazonaws.com/send-email";

// Current sandbox rate.
// After Production Access, adjust/remove this based on
// your actual SES MaxSendRate.
const SES_DELAY_MS = 1100;


// ============================================================
// SPREADSHEET
// ============================================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(
    SPREADSHEET_ID
  );
}


// ============================================================
// FORM SUBMIT ROUTER
// ============================================================

// Fires when either the main form or unsubscribe form is submitted.
function onFormSubmit(e) {

  if (!e || !e.range) {
    return;
  }

  const sheetName =
    e.range.getSheet().getName();

  if (sheetName === SHEET_NAME) {

    handleWelcomeSubmission(e);

  } else if (
    sheetName === UNSUB_SHEET_NAME
  ) {

    handleUnsubscribeSubmission(e);
  }
}


// ============================================================
// AMAZON SES
// ============================================================

/**
 * Sends an email through:
 *
 * Apps Script
 *   ↓
 * API Gateway
 *   ↓
 * Lambda
 *   ↓
 * SES
 *
 * The subject/body are supplied by the caller.
 */
function sendEmailViaSES(
  to,
  subject,
  textBody,
  htmlBody
) {

  if (!to) {
    throw new Error(
      "SES: recipient email is required."
    );
  }

  if (!subject) {
    throw new Error(
      "SES: subject is required."
    );
  }

  if (!textBody && !htmlBody) {
    throw new Error(
      "SES: textBody or htmlBody is required."
    );
  }

  const payload = {
    to: to,
    subject: subject,
    textBody: textBody || "",
    htmlBody: htmlBody || ""
  };

  const response =
    UrlFetchApp.fetch(
      SES_API_URL,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

  const statusCode =
    response.getResponseCode();

  const responseText =
    response.getContentText();

  Logger.log(
    "SES API response: HTTP " +
    statusCode +
    " - " +
    responseText
  );

  if (
    statusCode < 200 ||
    statusCode >= 300
  ) {

    throw new Error(
      "SES API request failed. HTTP " +
      statusCode +
      ": " +
      responseText
    );
  }

  let result;

  try {

    result =
      JSON.parse(responseText);

  } catch (err) {

    throw new Error(
      "SES API returned invalid JSON: " +
      responseText
    );
  }

  if (!result.success) {

    throw new Error(
      "SES rejected email: " +
      (result.message || responseText)
    );
  }

  return result;
}


// ============================================================
// WELCOME EMAIL
// ============================================================

// Handles a real signup on the main form.
function handleWelcomeSubmission(e) {

  if (!e.namedValues) {
    return;
  }

  const email =
    (
      e.namedValues[EMAIL_FIELD] ||
      [""]
    )[0]
      .trim();

  const name =
    (
      e.namedValues[NAME_FIELD] ||
      ["there"]
    )[0]
      .trim() ||
    "there";

  if (!email) {
    return;
  }

  const sheet =
    getSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Could not find a tab named "' +
      SHEET_NAME +
      '".'
    );
  }


  // Get Welcome email content from Message_Config.
  const config =
    getMessageConfig("Welcome");


  // Apply {{name}} placeholder.
  const subject =
    applyPlaceholders(
      config.subject,
      {
        name: name
      }
    );


  const rawBody =
    applyPlaceholders(
      config.body,
      {
        name: name
      }
    );


  const plainMessage =
    rawBody +
    buildUnsubscribeFooterPlain(
      email
    );


  const htmlMessage =
    buildEmailHtml(
      subject,
      textToHtml(rawBody),
      buildUnsubscribeFooterHtml(email)
    );


  // -------------------------------
  // AMAZON SES
  // -------------------------------
  const result =
    sendEmailViaSES(
      email,
      subject,
      plainMessage,
      htmlMessage
    );


  // Mark only after SES accepts it.
  markSent(
    email,
    SENT_WELCOME_HEADER
  );


  Logger.log(
    "Welcome email sent to " +
    email +
    ". SES MessageId: " +
    result.messageId
  );
}


// ============================================================
// UNSUBSCRIBE
// ============================================================

function handleUnsubscribeSubmission(e) {

  if (!e.namedValues) {
    return;
  }

  processUnsubscribeRow(
    e.range.getSheet(),
    e.range.getRow()
  );
}


// Core unsubscribe processing.
function processUnsubscribeRow(
  unsubSheet,
  row
) {

  const headers =
    getHeadersRow(unsubSheet);


  // Flexible email column detection.
  const emailColIdx =
    headers.findIndex(
      function(h) {

        return (
          h &&
          h
            .toString()
            .trim()
            .toLowerCase()
            .includes("email")
        );

      }
    );


  if (emailColIdx === -1) {

    Logger.log(
      "Could not find an 'Email' column on the " +
      "Unsubscribe form's response sheet. " +
      "Headers found: " +
      headers.join(" | ")
    );

    return;
  }


  const email =
    unsubSheet
      .getRange(
        row,
        emailColIdx + 1
      )
      .getValue()
      .toString()
      .trim();


  if (!email) {

    Logger.log(
      "Row " +
      row +
      " has no email value in column " +
      (emailColIdx + 1) +
      "."
    );

    return;
  }


  // Find unsubscribe reason.
  let reason = "";


  for (
    let c = 0;
    c < headers.length;
    c++
  ) {

    const h =
      headers[c]
        ? headers[c]
            .toString()
            .trim()
            .toLowerCase()
        : "";


    if (
      h !== "timestamp" &&
      !h.includes("email")
    ) {

      reason =
        unsubSheet
          .getRange(
            row,
            c + 1
          )
          .getValue()
          .toString()
          .trim();

      break;
    }
  }


  const mainSheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_NAME
      );


  if (!mainSheet) {

    throw new Error(
      'Could not find a tab named "' +
      SHEET_NAME +
      '".'
    );
  }


  ensureTrackingColumns(
    mainSheet
  );


  const mainHeaders =
    getHeadersRow(
      mainSheet
    );


  const mainEmailCol =
    mainHeaders.indexOf(
      EMAIL_FIELD
    ) + 1;


  const unsubCol =
    mainHeaders.indexOf(
      UNSUB_HEADER
    ) + 1;


  const reasonCol =
    mainHeaders.indexOf(
      UNSUB_REASON_HEADER
    ) + 1;


  Logger.log(
    "Looking for email '" +
    email +
    "' in main sheet column " +
    mainEmailCol
  );


  const lastRow =
    mainSheet.getLastRow();


  if (lastRow < 2) {

    Logger.log(
      "Main sheet has no data rows."
    );

    return;
  }


  const emails =
    mainSheet.getRange(
      2,
      mainEmailCol,
      lastRow - 1,
      1
    ).getValues();


  let matchCount = 0;


  for (
    let i = 0;
    i < emails.length;
    i++
  ) {

    const candidate =
      emails[i][0]
        ? emails[i][0]
            .toString()
            .trim()
            .toLowerCase()
        : "";


    if (
      candidate ===
      email.toLowerCase()
    ) {

      const mainRow =
        i + 2;


      mainSheet
        .getRange(
          mainRow,
          unsubCol
        )
        .setValue(true);


      mainSheet
        .getRange(
          mainRow,
          reasonCol
        )
        .setValue(reason);


      matchCount++;
    }
  }


  if (matchCount > 0) {

    Logger.log(
      "SUCCESS: Marked " +
      matchCount +
      " row(s) for " +
      email +
      " as unsubscribed. " +
      "Reason: " +
      reason
    );

  } else {

    Logger.log(
      "NO MATCH: '" +
      email +
      "' was not found among these emails in " +
      SHEET_NAME +
      ": " +
      emails
        .map(function(r) {
          return r[0];
        })
        .join(", ")
    );
  }
}


// ============================================================
// TEST UNSUBSCRIBE
// ============================================================

function testUnsubscribeRow() {

  const rowToTest = 2;


  const unsubSheet =
    getSpreadsheet()
      .getSheetByName(
        UNSUB_SHEET_NAME
      );


  if (!unsubSheet) {

    Logger.log(
      'Could not find a tab named "' +
      UNSUB_SHEET_NAME +
      '".'
    );

    return;
  }


  processUnsubscribeRow(
    unsubSheet,
    rowToTest
  );
}


// ============================================================
// WEEKLY EMAIL
// ============================================================

/**
 * Fires every Sunday at 9 AM.
 *
 * Sends weekly email to all eligible
 * recipients who are not unsubscribed.
 */
function sendWeeklyEmails() {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      'Could not find a tab named "' +
      SHEET_NAME +
      '".'
    );
  }


  ensureTrackingColumns(
    sheet
  );


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (data.length <= 1) {

    Logger.log(
      "No recipient rows found."
    );

    return;
  }


  const headers =
    data[0];


  const emailIndex =
    headers.indexOf(
      EMAIL_FIELD
    );


  const nameIndex =
    headers.indexOf(
      NAME_FIELD
    );


  const weeklyIndex =
    headers.indexOf(
      SENT_WEEKLY_HEADER
    );


  const unsubIndex =
    headers.indexOf(
      UNSUB_HEADER
    );


  if (emailIndex === -1) {

    Logger.log(
      "Email column not found."
    );

    return;
  }


  const config =
    getMessageConfig(
      "Weekly"
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const email =
      data[i][emailIndex]
        ? data[i][emailIndex]
            .toString()
            .trim()
        : "";


    const name =
      (
        nameIndex !== -1 &&
        data[i][nameIndex]
      )
        ? data[i][nameIndex]
            .toString()
            .trim()
        : "there";


    const isUnsubscribed =
      (
        unsubIndex !== -1 &&
        (
          data[i][unsubIndex] === true ||
          data[i][unsubIndex] === "TRUE"
        )
      );


    if (
      !email ||
      isUnsubscribed
    ) {
      continue;
    }


    const subject =
      applyPlaceholders(
        config.subject,
        {
          name: name
        }
      );


    const rawBody =
      applyPlaceholders(
        config.body,
        {
          name: name
        }
      );


    const plainMessage =
      rawBody +
      buildUnsubscribeFooterPlain(
        email
      );


    const htmlMessage =
      buildEmailHtml(
        subject,
        textToHtml(rawBody),
        buildUnsubscribeFooterHtml(
          email
        )
      );


    try {

      // -------------------------------
      // AMAZON SES
      // -------------------------------
      const result =
        sendEmailViaSES(
          email,
          subject,
          plainMessage,
          htmlMessage
        );


      // Mark only after SES succeeds.
      sheet
        .getRange(
          i + 1,
          weeklyIndex + 1
        )
        .setValue("YES");


      Logger.log(
        "Weekly email sent to " +
        email +
        ". SES MessageId: " +
        result.messageId
      );


      // Current SES sandbox rate:
      // approximately 1 email/sec.
      //
      // Remove/reduce this after Production Access
      // according to your actual MaxSendRate.
      Utilities.sleep(
        SES_DELAY_MS
      );


    } catch (err) {

      Logger.log(
        "Failed to send to " +
        email +
        ": " +
        err.message
      );
    }
  }
}


// ============================================================
// MESSAGE CONFIG
// ============================================================

function getMessageConfig(type) {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        CONFIG_SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      'Could not find a tab named "' +
      CONFIG_SHEET_NAME +
      '". ' +
      'Create it with columns Type | Subject | Body.'
    );
  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (data.length < 2) {

    throw new Error(
      CONFIG_SHEET_NAME +
      " does not contain any message rows."
    );
  }


  const headers =
    data[0];


  const typeCol =
    headers.indexOf(
      "Type"
    );


  const subjectCol =
    headers.indexOf(
      "Subject"
    );


  const bodyCol =
    headers.indexOf(
      "Body"
    );


  if (
    typeCol === -1 ||
    subjectCol === -1 ||
    bodyCol === -1
  ) {

    throw new Error(
      CONFIG_SHEET_NAME +
      " must contain columns: " +
      "Type | Subject | Body."
    );
  }


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    if (
      data[r][typeCol] &&
      data[r][typeCol]
        .toString()
        .trim()
        .toLowerCase() ===
        type.toLowerCase()
    ) {

      return {
        subject:
          data[r][subjectCol],

        body:
          data[r][bodyCol]
      };
    }
  }


  throw new Error(
    "No row found in " +
    CONFIG_SHEET_NAME +
    ' for Type = "' +
    type +
    '".'
  );
}


// ============================================================
// PLACEHOLDERS
// ============================================================

function applyPlaceholders(
  text,
  values
) {

  let result =
    text == null
      ? ""
      : text.toString();


  Object.keys(values)
    .forEach(function(key) {

      result =
        result
          .split(
            "{{" + key + "}}"
          )
          .join(
            values[key] == null
              ? ""
              : values[key]
          );

    });


  return result;
}


// ============================================================
// UNSUBSCRIBE FOOTER
// ============================================================

function buildUnsubscribeFooterPlain(
  email
) {

  return (
    "\n\n---\n" +
    "Don't want these emails? " +
    "Unsubscribe here: " +
    UNSUB_FORM_URL
  );
}


function buildUnsubscribeFooterHtml(
  email
) {

  return (
    '<a href="' +
    UNSUB_FORM_URL +
    '" style="color:#888888;">' +
    "Unsubscribe here" +
    "</a>"
  );
}


// ============================================================
// HTML RENDERING
// ============================================================

function linkify(text) {

  const urlPattern =
    /(https?:\/\/[^\s<]+)/g;


  return text.replace(
    urlPattern,
    function(url) {

      return (
        '<a href="' +
        url +
        '" style="color:#2A9D8F;">' +
        url +
        "</a>"
      );
    }
  );
}


function textToHtml(text) {

  const normalized =
    text.replace(
      /\n\s*\n+/g,
      "\n\n"
    );


  const paragraphs =
    normalized.split(
      /\n\n/
    );


  return paragraphs
    .map(function(p) {

      const trimmed =
        p.trim();


      if (!trimmed) {
        return "";
      }


      const withBreaks =
        trimmed
          .split("\n")
          .join("<br>");


      return (
        '<p style="' +
        'margin:0 0 16px 0;' +
        'color:#333333;' +
        'font-size:15px;' +
        'line-height:1.6;">' +
        linkify(withBreaks) +
        "</p>"
      );

    })
    .filter(function(p) {

      return p !== "";

    })
    .join("");
}


function buildEmailHtml(
  subject,
  bodyHtml,
  unsubscribeLinkHtml
) {

  return (
    '<div style="' +
    'background-color:#f4f4f4;' +
    'padding:32px 16px;' +
    'font-family:Arial,Helvetica,sans-serif;">' +

    '<table role="presentation" width="100%" ' +
    'style="' +
    'max-width:600px;' +
    'margin:0 auto;' +
    'background-color:#ffffff;' +
    'border-radius:10px;' +
    'overflow:hidden;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.08);">' +

    '<tr>' +

    '<td align="center" ' +
    'style="' +
    'background-color:#1D3557;' +
    'padding:26px 24px 22px 24px;">' +

    '<div style="' +
    'color:#ffffff;' +
    'font-size:20px;' +
    'font-weight:bold;">' +

    'Finance 1-0-1' +

    '</div>' +

    '</td>' +

    '</tr>' +

    '<tr>' +

    '<td style="' +
    'height:4px;' +
    'background-color:#2A9D8F;">' +

    '</td>' +

    '</tr>' +

    '<tr>' +

    '<td style="' +
    'padding:26px 26px 10px 26px;">' +

    bodyHtml +

    '</td>' +

    '</tr>' +

    '<tr>' +

    '<td style="' +
    'padding:14px 26px 24px 26px;' +
    'border-top:1px solid #eeeeee;">' +

    '<div style="' +
    'color:#999999;' +
    'font-size:12px;' +
    'line-height:1.6;">' +

    'You are receiving this as part of Finance 1-0-1.' +

    '<br>' +

    "Don't want these emails? " +

    unsubscribeLinkHtml +

    '.' +

    '</div>' +

    '</td>' +

    '</tr>' +

    '</table>' +

    '</div>'
  );
}


// ============================================================
// TRACKING HELPERS
// ============================================================

function ensureTrackingColumns(
  sheet
) {

  const headers =
    getHeadersRow(sheet);


  const toAdd = [
    UNSUB_HEADER,
    UNSUB_REASON_HEADER
  ].filter(function(h) {

    return (
      headers.indexOf(h) === -1
    );
  });


  toAdd.forEach(
    function(h, i) {

      sheet
        .getRange(
          1,
          headers.length + 1 + i
        )
        .setValue(h);

    }
  );
}


function getHeadersRow(sheet) {

  return sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];
}


// ============================================================
// SETUP
// ============================================================

/**
 * Run this ONCE.
 *
 * Creates:
 *
 * 1. Form submission trigger
 * 2. Weekly Sunday 9 AM trigger
 */
function setup() {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      'Could not find a tab named "' +
      SHEET_NAME +
      '".'
    );
  }


  ensureTrackingColumns(
    sheet
  );


  const managedHandlers = [
    "onFormSubmit",
    "sendWeeklyEmails"
  ];


  ScriptApp
    .getProjectTriggers()
    .forEach(
      function(trigger) {

        if (
          managedHandlers.indexOf(
            trigger.getHandlerFunction()
          ) !== -1
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );
        }
      }
    );


  // Main + unsubscribe form trigger.
  ScriptApp
    .newTrigger(
      "onFormSubmit"
    )
    .forSpreadsheet(
      getSpreadsheet()
    )
    .onFormSubmit()
    .create();


  // Weekly Sunday at 9 AM.
  ScriptApp
    .newTrigger(
      "sendWeeklyEmails"
    )
    .timeBased()
    .onWeekDay(
      ScriptApp.WeekDay.SUNDAY
    )
    .atHour(9)
    .create();


  Logger.log(
    "Setup complete. " +
    "Form submit trigger + weekly Sunday trigger installed."
  );
}


// ============================================================
// MARK SENT
// ============================================================

function markSent(
  email,
  columnHeader
) {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_NAME
      );


  const data =
    sheet
      .getDataRange()
      .getValues();


  const headers =
    data[0];


  const emailIndex =
    headers.indexOf(
      EMAIL_FIELD
    );


  let colIndex =
    headers.indexOf(
      columnHeader
    );


  if (colIndex === -1) {

    colIndex =
      headers.length;

    sheet
      .getRange(
        1,
        colIndex + 1
      )
      .setValue(
        columnHeader
      );
  }


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      data[i][emailIndex] &&
      data[i][emailIndex]
        .toString()
        .trim()
        .toLowerCase() ===
      email
        .trim()
        .toLowerCase()
    ) {

      sheet
        .getRange(
          i + 1,
          colIndex + 1
        )
        .setValue("YES");

      break;
    }
  }
}