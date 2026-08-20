/**
 * TEAM EXCELLENCE — Client Birthday Wishes
 * ---------------------------------------------------------
 * WHAT THIS DOES
 * Runs once a day and checks every row in your sheet. Based on today's
 * date, it sends one of three types of email to the client:
 *
 *   1. If it's the CLIENT's own birthday -> sends them a direct
 *      "Happy Birthday" wish.
 *   2. If it's their CHILD's birthday -> sends the client a reminder
 *      to wish their child a happy birthday.
 *   3. If it's their SPOUSE's birthday -> sends the client a reminder
 *      to wish their spouse a happy birthday.
 *      
 * A row can trigger more than one of these on the same day if, say,
 * the child's birthday happens to match — each is independent.
 *
 * ===========================================================
 * SETUP — DO THESE STEPS IN ORDER
 * ===========================================================
 *
 * STEP 1 — Add an Email column
 *   In your sheet, add a column titled exactly: Email
 *   Fill in each client's email address in that column.
 *
 * STEP 2 — Add the Birthday_Config tab
 *   Create a new tab named exactly: Birthday_Config
 *   Headers in row 1: Type | Subject | Body
 *   Add exactly 3 rows, one per Type: Own, Child, Spouse
 *   (Or import Birthday_Config.csv via File > Import > Replace current sheet
 *   with that tab active.)
 *
 *   Placeholders you can use inside Subject/Body:
 *     {{name}}       -> the client's first name (always available)
 *     {{childName}}  -> only fills in for the Child row
 *     {{spouseName}} -> only fills in for the Spouse row
 *   Lines starting with "- " become bullet points automatically.
 *
 * STEP 3 — Paste this script
 *   Open your Sheet -> Extensions > Apps Script -> paste this file in.
 *   Fill in SPREADSHEET_ID below (from your Sheet's URL).
 *
 * STEP 4 — Run setup()
 *   Function dropdown -> select "setup" -> Run. Approve permissions.
 *   This installs a daily trigger (runs once a day, default 8am) and
 *   adds 3 hidden tracking columns so nobody gets the same wish twice
 *   in one day even if the script runs more than once.
 *
 * STEP 5 — Test it
 *   Use testBirthdayRow() at the bottom of this file — edit the row
 *   number, then Run it. This forces that row's emails to send right
 *   now, regardless of today's actual date, so you can preview them
 *   without waiting for a real birthday.
 * ---------------------------------------------------------
 */

// ====================== CONFIG ======================

const SHEET_NAME = 'Birthdays';
const CONFIG_SHEET_NAME = 'Birthday_Config';

const COL_NAME = 'Name';
const COL_EMAIL = 'Email';
const COL_BIRTHDAY = 'Birthday';
const COL_CHILD_NAME = 'ChildName';
const COL_CHILD_BIRTHDAY = 'ChildBirthday';
const COL_SPOUSE_NAME = 'Spouse Name';
const COL_SPOUSE_BIRTHDAY = 'Spouse Birthday';

// Tracking columns (auto-added if missing) — store the date (yyyy-MM-dd)
// each wish type was last sent, so re-running the same day doesn't double-send.
const TRACK_OWN = 'Last Own Bday Sent';
const TRACK_CHILD = 'Last Child Bday Sent';
const TRACK_SPOUSE = 'Last Spouse Bday Sent';

// PASTE YOUR SPREADSHEET ID HERE (from your Sheet's URL, between /d/ and /edit)
const SPREADSHEET_ID = '1oC1QEAuhmtP073cp6vCXcPPFVp2X5lxts97hhDY_CFg';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ====================== CORE LOGIC ======================

function checkBirthdaysToday() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureColumns(sheet);

  const todayStr = formatDate(new Date());
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    processBirthdayRow(sheet, row, todayStr);
  }
}

function processBirthdayRow(sheet, row, todayStr) {
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);

  const name = sheet.getRange(row, col(COL_NAME)).getValue();
  const email = sheet.getRange(row, col(COL_EMAIL)).getValue();
  if (!name || !email) return;

  const firstName = name.toString().trim().split(' ')[0];

  // ---- Client's own birthday ----
  const birthday = sheet.getRange(row, col(COL_BIRTHDAY)).getValue();
  if (isTodayMonthDay(birthday) && sheet.getRange(row, col(TRACK_OWN)).getValue() !== todayStr) {
    sendConfiguredEmail('Own', email, { name: firstName });
    sheet.getRange(row, col(TRACK_OWN)).setValue(todayStr);
  }

  // ---- Child's birthday ----
  const childName = sheet.getRange(row, col(COL_CHILD_NAME)).getValue();
  const childBirthday = sheet.getRange(row, col(COL_CHILD_BIRTHDAY)).getValue();
  if (childName && isTodayMonthDay(childBirthday) && sheet.getRange(row, col(TRACK_CHILD)).getValue() !== todayStr) {
    sendConfiguredEmail('Child', email, { name: firstName, childName: childName.toString().trim() });
    sheet.getRange(row, col(TRACK_CHILD)).setValue(todayStr);
  }

  // ---- Spouse's birthday ----
  const spouseName = sheet.getRange(row, col(COL_SPOUSE_NAME)).getValue();
  const spouseBirthday = sheet.getRange(row, col(COL_SPOUSE_BIRTHDAY)).getValue();
  if (spouseName && isTodayMonthDay(spouseBirthday) && sheet.getRange(row, col(TRACK_SPOUSE)).getValue() !== todayStr) {
    sendConfiguredEmail('Spouse', email, { name: firstName, spouseName: spouseName.toString().trim() });
    sheet.getRange(row, col(TRACK_SPOUSE)).setValue(todayStr);
  }
}

// ====================== MESSAGE CONFIG (from Birthday_Config sheet) ======================

/**
 * Reads the Birthday_Config sheet and returns {subject, body} for the
 * given type ("Own", "Child", or "Spouse"), matched case-insensitively.
 */
function getMessageConfig(type) {
  const sheet = getSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Could not find a tab named "${CONFIG_SHEET_NAME}". Create it with columns Type | Subject | Body.`);
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = findColumn(headers, 'Type') - 1;
  const subjectCol = findColumn(headers, 'Subject') - 1;
  const bodyCol = findColumn(headers, 'Body') - 1;

  for (let r = 1; r < data.length; r++) {
    if (data[r][typeCol] && data[r][typeCol].toString().trim().toLowerCase() === type.toLowerCase()) {
      return { subject: data[r][subjectCol], body: data[r][bodyCol] };
    }
  }
  throw new Error(`No row found in ${CONFIG_SHEET_NAME} for Type = "${type}". Add a row with Type set to exactly "${type}".`);
}

/**
 * Replaces {{placeholder}} tokens in a string with values from a map.
 * Unmatched placeholders are left as-is (harmless if a type doesn't use them).
 */
function applyPlaceholders(text, values) {
  let result = text;
  Object.keys(values).forEach(key => {
    result = result.split(`{{${key}}}`).join(values[key]);
  });
  return result;
}

function sendConfiguredEmail(type, email, values) {
  const config = getMessageConfig(type);
  const subject = applyPlaceholders(config.subject, values);
  const rawBody = applyPlaceholders(config.body, values);
  const html = buildEmailHtml(subject, textToHtml(rawBody));

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: rawBody,
    htmlBody: html
  });
}

/**
 * Converts simple plain text into styled HTML.
 * - Blank lines separate paragraphs.
 * - Lines starting with "- " become a bulleted list.
 * - Everything else becomes a paragraph with <br> for line breaks.
 */
function textToHtml(text) {
  const lines = text.toString().split('\n');
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
      html += '<ul style="margin:0 0 16px 0;padding-left:20px;">' +
        items.map(it => `<li style="margin-bottom:8px;color:#3a3a3a;font-size:15px;line-height:1.5;">${it}</li>`).join('') +
        '</ul>';
    } else {
      const para = [];
      while (i < lines.length && lines[i].trim() !== '' && !lines[i].trim().startsWith('- ')) {
        para.push(lines[i]);
        i++;
      }
      html += `<p style="margin:0 0 16px 0;color:#3a3a3a;font-size:15px;line-height:1.6;">${para.join('<br>')}</p>`;
    }
  }
  return html;
}

/**
 * Wraps content in a branded HTML email template
 * (gold/navy theme matching the Team Excellence style).
 */
function buildEmailHtml(subject, bodyHtml) {
  return `
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light only">
<meta name="supported-color-schemes" content="light only">
<style>
  :root { color-scheme: light only; supported-color-schemes: light only; }
  body { margin:0; padding:0; -webkit-text-size-adjust:100%; }
</style>
</head>
<body style="margin:0;padding:0;background-color:#f4f4f4;">
<div style="background-color:#f4f4f4;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" style="max-width:560px;margin:0 auto;background-color:#3D1A0B;
    border-radius:24px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.18);">
 
    <!-- Top decoration -->
    <tr>
      <td align="center" bgcolor="#3D1A0B" style="background-color:#3D1A0B;padding:32px 24px 0 24px;font-size:30px;line-height:1;">
        🎈🎉🎁🎊🎈
      </td>
    </tr>
 
    <!-- Big Happy Birthday banner -->
    <tr>
      <td align="center" bgcolor="#3D1A0B" style="background-color:#3D1A0B;padding:12px 24px 4px 24px;">
        <div style="font-family:'Brush Script MT','Segoe Script',cursive;color:#FF7A3D !important;font-size:48px;
          line-height:1.1;font-style:italic;">
          Happy
        </div>
        <div style="color:#FCE8C9 !important;font-size:42px;font-weight:900;letter-spacing:2px;line-height:1.1;
          margin-top:-6px;font-family:Arial,Helvetica,sans-serif;">
          BIRTHDAY
        </div>
      </td>
    </tr>
 
    <!-- Specific subject line -->
    <tr>
      <td align="center" bgcolor="#3D1A0B" style="background-color:#3D1A0B;padding:10px 24px 28px 24px;">
        <div style="color:#FFD98E !important;font-size:17px;font-weight:bold;">
          ${subject}
        </div>
      </td>
    </tr>
 
    <!-- Message card -- warm colorful gradient instead of plain cream -->
    <tr>
      <td bgcolor="#3D1A0B" style="background-color:#3D1A0B;padding:0 24px 32px 24px;">
        <table role="presentation" width="100%" style="border-radius:16px;overflow:hidden;">
          <tr>
            <td bgcolor="#FFE3B3" style="background-color:#FFE3B3;padding:26px 24px;
              background:linear-gradient(135deg,#FFE3B3 0%,#FFD1A9 45%,#FFC2C2 100%);">
              ${bodyHtml}
            </td>
          </tr>
        </table>
      </td>
    </tr>
 
    <!-- Bottom decoration -->
    <tr>
      <td align="center" bgcolor="#3D1A0B" style="background-color:#3D1A0B;padding:0 24px 32px 24px;font-size:28px;line-height:1;">
        🎁🎈🎊🎉🎁
      </td>
    </tr>
  </table>
</div>
</body>
</html>`;
}
// ====================== DATE HELPERS ======================

function isTodayMonthDay(value) {
  if (!value) return false;
  const d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return false;

  const today = new Date();
  return d.getMonth() === today.getMonth() && d.getDate() === today.getDate();
}

function formatDate(date) {
  const tz = getSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

// ====================== SETUP HELPERS ======================

function ensureColumns(sheet) {
  const headers = getHeaders(sheet);
  const normalized = headers.map(h => h ? h.toString().trim().toLowerCase() : '');
  const required = [COL_EMAIL, TRACK_OWN, TRACK_CHILD, TRACK_SPOUSE];
  const toAdd = required.filter(h => normalized.indexOf(h.trim().toLowerCase()) === -1);
  toAdd.forEach((h, i) => {
    sheet.getRange(1, headers.length + 1 + i).setValue(h);
  });
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

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
      `Could not find a tab named "${SHEET_NAME}". Tabs found: ${names}. ` +
      `Update SHEET_NAME to match exactly.`
    );
  }
  ensureColumns(sheet);

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkBirthdaysToday')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Setup complete. Daily birthday check installed (runs ~8am).');
}

// ====================== MANUAL TEST HELPER ======================

/**
 * TEST FUNCTION — forces a specific row's birthday emails to send RIGHT
 * NOW, regardless of what today's actual date is. Useful for previewing
 * content without waiting for a real birthday to roll around.
 *
 * Edit rowToTest below, select "testBirthdayRow" from the function
 * dropdown, and click Run.
 */
function testBirthdayRow() {
  const rowToTest = 2; // change to the row number you want to test (row 2 = first data row)

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);

  const name = sheet.getRange(rowToTest, col(COL_NAME)).getValue();
  const email = sheet.getRange(rowToTest, col(COL_EMAIL)).getValue();
  const childName = sheet.getRange(rowToTest, col(COL_CHILD_NAME)).getValue();
  const spouseName = sheet.getRange(rowToTest, col(COL_SPOUSE_NAME)).getValue();

  if (!email) {
    Logger.log(`Row ${rowToTest} has no Email value. Add one first.`);
    return;
  }

  const firstName = name.toString().trim().split(' ')[0];

  sendConfiguredEmail('Own', email, { name: firstName });
  Logger.log(`Sent OWN birthday email to ${email}`);

  if (childName) {
    sendConfiguredEmail('Child', email, { name: firstName, childName: childName.toString().trim() });
    Logger.log(`Sent CHILD birthday reminder to ${email}`);
  }

  if (spouseName) {
    sendConfiguredEmail('Spouse', email, { name: firstName, spouseName: spouseName.toString().trim() });
    Logger.log(`Sent SPOUSE birthday reminder to ${email}`);
  }
}