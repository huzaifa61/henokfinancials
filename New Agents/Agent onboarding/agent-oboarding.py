/**
 * TEAM EXCELLENCE — New Member Drip Campaign (v3)
 * ---------------------------------------------------------
 * WHAT'S NEW IN V3
 * • No Web App deployment needed — unsubscribe now works through a
 *   second Google Form. This avoids org/Workspace policies that block
 *   "Anyone access" Web App deployments.
 * • Messages/days are configured in a SHEET TAB (Message_Config).
 * • Emails render as styled HTML (branded header, bullets, CTA button).
 *
 * ===========================================================
 * SETUP — DO THESE STEPS IN ORDER
 * ===========================================================
 *
 * STEP 1 — Message_Config tab
 *   Create a sheet tab named exactly: Message_Config
 *   Headers in row 1: Offset Days | Subject | Body | Link Label | Link URL
 *   (Or just import Message_Config.csv via File > Import > Replace current sheet)
 *
 * STEP 2 — Create the Unsubscribe form
 *   a. Create a new Google Form named "Unsubscribe".
 *   b. Add ONE short-answer question titled exactly: Token
 *   c. Go to Responses tab -> click the green Sheets icon -> "Select
 *      existing spreadsheet" -> choose THIS spreadsheet. This adds a new
 *      response tab — rename it to Unsubscribe_Responses if it isn't already.
 *   d. On the form, click the 3-dot menu (top right) -> "Get pre-filled link".
 *      Type any placeholder text as the answer, click "Get link", copy it.
 *      It looks like:
 *      https://docs.google.com/forms/d/e/AbCdEf.../viewform?usp=pp_url&entry.123456789=placeholder
 *   e. In the CONFIG section below, paste everything BEFORE "&entry." into
 *      UNSUB_FORM_BASE_URL, and the "entry.123456789" part into
 *      UNSUB_FORM_TOKEN_ENTRY.
 *
 * STEP 3 — Paste this script
 *   Extensions > Apps Script (opened from within your Sheet, not
 *   script.google.com directly) > paste this file in.
 *   Fill in SPREADSHEET_ID (from your Sheet's URL), UNSUB_FORM_BASE_URL,
 *   and UNSUB_FORM_TOKEN_ENTRY.
 *
 * STEP 4 — Install triggers
 *   Function dropdown -> select "setup" -> Run. Approve permissions.
 *   This installs BOTH triggers: on-form-submit (fires for either form,
 *   since both write into this spreadsheet) and a daily check.
 *
 * STEP 5 — Test
 *   Submit the main onboarding form. You should get a styled HTML email
 *   with a working Unsubscribe link within a minute.
 * ---------------------------------------------------------
 */

// ====================== CONFIG ======================

const SHEET_NAME = 'Form_Responses';
const CONFIG_SHEET_NAME = 'Message_Config';

const COL_TIMESTAMP = 'Timestamp';
const COL_NAME = 'Full name';
const COL_EMAIL = 'Email';

const TRACKING_COL = 'Last Step Sent';
const UNSUB_COL = 'Unsubscribed';
const TOKEN_COL = 'Unsub Token';

// (Web app deployment is no longer needed — unsubscribe now works via a
// second Google Form, which sidesteps org policies that block "Anyone"
// access Web App deployments.)

// PASTE YOUR SPREADSHEET ID HERE.
// Find it in your Sheet's URL:
// https://docs.google.com/spreadsheets/d/THIS_LONG_ID_HERE/edit
const SPREADSHEET_ID = '1MCG4X48Blay3V5MjryXqn2Xnt4RZTEkBKfmRc2aWgco';

const UNSUB_SHEET_NAME = 'Unsubscribe_Responses';
const UNSUB_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSeE6rejSnVys96fVbKu6hpB7N_f8HlUaGdPsFUOJUwZjSK3eA/viewform?usp=pp_url'; // e.g. https://docs.google.com/forms/d/e/xxxx/viewform?usp=pp_url
const UNSUB_FORM_TOKEN_ENTRY = 'entry.1193594610';      // e.g. entry.123456789

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function buildUnsubUrl(token) {
  return `${UNSUB_FORM_BASE_URL}&${UNSUB_FORM_TOKEN_ENTRY}=${encodeURIComponent(token)}`;
}

// ====================== CORE LOGIC ======================

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_NAME) {
    ensureColumns(sheet);
    processRow(sheet, e.range.getRow());
  } else if (sheetName === UNSUB_SHEET_NAME) {
    processUnsubscribeSubmission(sheet, e.range.getRow());
  }
}

/**
 * Called when someone submits the Unsubscribe form.
 * Looks up their token in Form_Responses and flags them.
 */
function processUnsubscribeSubmission(unsubSheet, row) {
  const headers = getHeaders(unsubSheet);
  const tokenCol = findColumn(headers, 'Token');
  const token = unsubSheet.getRange(row, tokenCol).getValue();
  if (!token) return;

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const mainHeaders = getHeaders(sheet);
  const tokenColMain = findColumn(mainHeaders, TOKEN_COL);
  const unsubCol = findColumn(mainHeaders, UNSUB_COL);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tokens = sheet.getRange(2, tokenColMain, lastRow - 1, 1).getValues();
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i][0] === token) {
      sheet.getRange(i + 2, unsubCol).setValue(true);
      break;
    }
  }
}

function checkAndSendDripEmails() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureColumns(sheet);
  const lastRow = sheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    processRow(sheet, row);
  }
}

function processRow(sheet, row) {
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);

  const timestamp = sheet.getRange(row, col(COL_TIMESTAMP)).getValue();
  const email = sheet.getRange(row, col(COL_EMAIL)).getValue();
  const name = (sheet.getRange(row, col(COL_NAME)).getValue() || 'there').toString().split(' ')[0];

  if (!timestamp || !email) return;

  // Assign a permanent unsubscribe token if this row doesn't have one yet
  let token = sheet.getRange(row, col(TOKEN_COL)).getValue();
  if (!token) {
    token = Utilities.getUuid();
    sheet.getRange(row, col(TOKEN_COL)).setValue(token);
  }

  // Skip unsubscribed people entirely
  const unsubscribed = sheet.getRange(row, col(UNSUB_COL)).getValue();
  if (unsubscribed === true || unsubscribed === 'TRUE') return;

  let lastStepSent = sheet.getRange(row, col(TRACKING_COL)).getValue();
  lastStepSent = (lastStepSent === '' || lastStepSent === null) ? -1 : Number(lastStepSent);

  const steps = getStepsFromConfig();
  const daysSince = Math.floor((new Date() - new Date(timestamp)) / (1000 * 60 * 60 * 24));

  for (let i = lastStepSent + 1; i < steps.length; i++) {
    if (steps[i].offsetDays <= daysSince) {
      sendStepEmail(email, name, token, steps[i]);
      sheet.getRange(row, col(TRACKING_COL)).setValue(i);
    } else {
      break; // steps sorted ascending by offsetDays
    }
  }
}

/**
 * Reads the Message_Config sheet and returns an array of steps,
 * sorted by Offset Days ascending.
 */
function getStepsFromConfig() {
  const sheet = getSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = (name) => headers.indexOf(name);

  const steps = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[idx('Subject')] === '' && row[idx('Body')] === '') continue; // skip blank rows
    steps.push({
      offsetDays: Number(row[idx('Offset Days')]),
      subject: row[idx('Subject')],
      body: row[idx('Body')],
      linkLabel: row[idx('Link Label')],
      linkUrl: row[idx('Link URL')]
    });
  }
  steps.sort((a, b) => a.offsetDays - b.offsetDays);
  return steps;
}

function sendStepEmail(email, firstName, token, step) {
  const rawBody = step.body.replace(/{{name}}/g, firstName);
  const unsubUrl = buildUnsubUrl(token);

  const htmlBody = buildEmailHtml(step.subject, textToHtml(rawBody), step.linkLabel, step.linkUrl, unsubUrl);

  let plainBody = rawBody;
  if (step.linkUrl) plainBody += `\n\n${step.linkLabel || 'Link'}: ${step.linkUrl}`;
  plainBody += `\n\n---\nDon't want these emails? Unsubscribe here: ${unsubUrl}`;

  MailApp.sendEmail({
    to: email,
    subject: step.subject,
    body: plainBody,     // fallback for clients that can't render HTML
    htmlBody: htmlBody
  });
}

/**
 * Converts simple plain text into styled HTML.
 * - Blank lines separate paragraphs.
 * - Lines starting with "- " become a bulleted list.
 * - Everything else becomes a paragraph with <br> for line breaks.
 */
function textToHtml(text) {
  const lines = text.split('\n');
  let html = '';
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === '') { i++; continue; }

    if (line.trim().startsWith('- ')) {
      const items = [];
      while (i < lines.length && lines[i].trim().startsWith('- ')) {
        items.push(lines[i].trim().substring(2));
        i++;
      }
      html += '<ul style="margin:0 0 18px 0;padding-left:22px;">' +
        items.map(it => `<li style="margin-bottom:9px;color:#3a3a3a;font-size:15px;line-height:1.55;">${it}</li>`).join('') +
        '</ul>';
    } else {
      const para = [];
      while (i < lines.length && lines[i].trim() !== '' && !lines[i].trim().startsWith('- ')) {
        para.push(lines[i]);
        i++;
      }
      html += `<p style="margin:0 0 18px 0;color:#3a3a3a;font-size:15px;line-height:1.6;">${para.join('<br>')}</p>`;
    }
  }
  return html;
}

/**
 * Wraps content in a branded HTML email template
 * (gold/navy theme matching the Team Excellence playbook).
 */
function buildEmailHtml(subject, bodyHtml, linkLabel, linkUrl, unsubUrl) {
  const goldButton = linkUrl ? `
    <tr>
      <td align="center" style="padding-top:4px;padding-bottom:24px;">
        <a href="${linkUrl}" style="background-color:#B8952B;color:#ffffff;text-decoration:none;
          font-size:15px;font-weight:bold;padding:12px 28px;border-radius:6px;display:inline-block;">
          ${linkLabel || 'View More'} →
        </a>
      </td>
    </tr>` : '';

  return `
<div style="background-color:#f4f4f4;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" style="max-width:600px;margin:0 auto;background-color:#ffffff;
    border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08);">

    <!-- Header -->
    <tr>
      <td align="center" style="background-color:#1F3B57;padding:28px 24px 22px 24px;">
        <div style="color:#D9B84A;font-size:14px;font-weight:bold;letter-spacing:1px;text-transform:uppercase;">
          ⭐ Team Excellence ⭐
        </div>
        <div style="color:#ffffff;font-size:22px;font-weight:bold;margin-top:10px;">
          ${subject}
        </div>
      </td>
    </tr>

    <!-- Gold divider -->
    <tr><td style="height:4px;background-color:#B8952B;"></td></tr>

    <!-- Body -->
    <tr>
      <td style="padding:28px 28px 8px 28px;">
        ${bodyHtml}
      </td>
    </tr>

    <table role="presentation" width="100%"><tbody>${goldButton}</tbody></table>

    <!-- Footer -->
    <tr>
      <td style="padding:18px 28px 26px 28px;border-top:1px solid #eeeeee;">
        <div style="color:#999999;font-size:12px;line-height:1.6;">
          You're receiving this as part of your Team Excellence onboarding.<br>
          Don't want these emails? <a href="${unsubUrl}" style="color:#999999;">Unsubscribe here</a>.
        </div>
      </td>
    </tr>
  </table>
</div>`;
}

// (doGet removed — unsubscribe is now handled via the Unsubscribe Google
// Form + processUnsubscribeSubmission(), not a Web App endpoint.)

// ====================== SETUP HELPERS ======================

function ensureColumns(sheet) {
  const headers = getHeaders(sheet);
  const normalized = headers.map(h => h ? h.toString().trim().toLowerCase() : '');
  const toAdd = [TRACKING_COL, UNSUB_COL, TOKEN_COL].filter(
    h => normalized.indexOf(h.trim().toLowerCase()) === -1
  );
  toAdd.forEach((h, i) => {
    sheet.getRange(1, headers.length + 1 + i).setValue(h);
  });
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Finds a column by header name, ignoring leading/trailing whitespace and
 * letter case (Google Forms sometimes writes headers with stray spaces).
 * Returns the 1-based column number, or 0 if truly not found.
 */
function findColumn(headers, name) {
  const target = name.toString().trim().toLowerCase();
  const idx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === target);
  if (idx === -1) {
    throw new Error(`Column "${name}" not found. Headers present: ${headers.join(' | ')}`);
  }
  return idx + 1;
}

function setup() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const names = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(
      `Could not find a tab named "${SHEET_NAME}". ` +
      `Tabs that actually exist in this spreadsheet: ${names}. ` +
      `Update the SHEET_NAME constant to match exactly (case-sensitive).`
    );
  }
  ensureColumns(sheet);

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('checkAndSendDripEmails')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('Setup complete. Triggers installed.');
}

// ====================== MANUAL TEST HELPER ======================
 
/**
 * TEST FUNCTION — sends one specific step to one specific email,
 * without touching Form_Responses or the Last Step Sent tracking.
 * Edit the values below, then select "testSendStep" from the function
 * dropdown at the top of the editor and click Run.
 */
function testSendStep() {
  const testEmail = 'work.shaikhhuzaifa@gmail.com';
  const testFirstName = 'Huzaifa';
  const offsetDaysToTest = 3; // 0 = Welcome, 3 = Core Values, 6 = Standards, 9 = What It Takes to Win
 
  const steps = getStepsFromConfig();
  const step = steps.find(s => s.offsetDays === offsetDaysToTest);
 
  if (!step) {
    Logger.log(`No step found with Offset Days = ${offsetDaysToTest}. Check your Message_Config sheet.`);
    return;
  }
 
  const testToken = 'TEST-' + Utilities.getUuid(); // dummy token, just for the unsubscribe link preview
  sendStepEmail(testEmail, testFirstName, testToken, step);
 
  Logger.log(`Sent "${step.subject}" to ${testEmail}`);
}

/**
 * TEST FUNCTION — verifies the unsubscribe gate actually blocks sending.
 * Requires that testEmail already has a row in Form_Responses (submit the
 * form once with that email first if it doesn't).
 *
 * What it does:
 *   Phase 1: Sets Unsubscribed = TRUE, forces a step to be "due", runs
 *            processRow, and checks that NO email was sent.
 *   Phase 2: Sets Unsubscribed = FALSE, same due step, runs processRow
 *            again, and checks that an email WAS sent this time.
 *   Then restores that row's original Unsubscribed / Last Step Sent values.
 *
 * NOTE: Phase 2 genuinely sends one real email — check that inbox to
 * confirm only ONE email arrives (from Phase 2, not Phase 1).
 */
function testUnsubscribeGate() {
  const testEmail = 'work.shaikhhuzaifa@gmail.com'; // must already exist as a row in Form_Responses
 
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
 
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in Form_Responses yet.');
    return;
  }
 
  const emails = sheet.getRange(2, col(COL_EMAIL), lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < emails.length; i++) {
    if (emails[i][0].toString().trim().toLowerCase() === testEmail.trim().toLowerCase()) {
      targetRow = i + 2;
      break;
    }
  }
 
  if (targetRow === -1) {
    Logger.log(`No row found for ${testEmail}. Submit the form once with that email first, or change testEmail to an address that already exists in Form_Responses.`);
    return;
  }
 
  // Backup original values so we can restore them afterward
  const originalUnsub = sheet.getRange(targetRow, col(UNSUB_COL)).getValue();
  const originalLastStep = sheet.getRange(targetRow, col(TRACKING_COL)).getValue();
 
  // ---- PHASE 1: Unsubscribed = TRUE, force a step to be "due" ----
  sheet.getRange(targetRow, col(UNSUB_COL)).setValue(true);
  sheet.getRange(targetRow, col(TRACKING_COL)).setValue(-1);
  processRow(sheet, targetRow);
 
  const afterPhase1 = sheet.getRange(targetRow, col(TRACKING_COL)).getValue();
  const phase1Passed = (afterPhase1 === -1 || afterPhase1 === '');
  Logger.log(phase1Passed
    ? '✅ PHASE 1 PASSED: Unsubscribed = TRUE correctly blocked the email (Last Step Sent stayed unchanged).'
    : '❌ PHASE 1 FAILED: An email may have been sent even though Unsubscribed = TRUE.');
 
  // ---- PHASE 2: Unsubscribed = FALSE, same due step ----
  sheet.getRange(targetRow, col(UNSUB_COL)).setValue(false);
  sheet.getRange(targetRow, col(TRACKING_COL)).setValue(-1);
  processRow(sheet, targetRow);
 
  const afterPhase2 = sheet.getRange(targetRow, col(TRACKING_COL)).getValue();
  const phase2Passed = (afterPhase2 !== -1 && afterPhase2 !== '');
  Logger.log(phase2Passed
    ? `✅ PHASE 2 PASSED: Unsubscribed = FALSE correctly allowed sending (Last Step Sent is now ${afterPhase2}).`
    : '❌ PHASE 2 FAILED: No email appears to have been sent even though Unsubscribed = FALSE.');
 
  // ---- Restore original values ----
  sheet.getRange(targetRow, col(UNSUB_COL)).setValue(originalUnsub);
  sheet.getRange(targetRow, col(TRACKING_COL)).setValue(originalLastStep);
 
  Logger.log(`Original values restored for ${testEmail}. Check that inbox — exactly ONE email should have arrived (from Phase 2).`);
}
 
 