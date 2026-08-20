/**
 * CLIENT HOLIDAY & MONTHLY REMINDERS
 * ---------------------------------------------------------
 * Completely separate from the Team Excellence onboarding drip script --
 * different audience (a dedicated Clients list), different trigger model
 * (calendar dates, not "days since signup"), so it's kept isolated to
 * avoid any risk of interfering with the working onboarding system.
 *
 * WHAT THIS DOES
 * 1. HOLIDAY / SPECIFIC DATE REMINDERS
 *    You maintain a list of dates in the Holiday_Dates tab. Each row has
 *    its own "Days Before" setting -- 0 means send exactly ON that date,
 *    1 means send the day before, 2 means two days before, etc. Mix and
 *    match freely: e.g. a "Merry Christmas!" row with Days Before = 0
 *    (sent on Dec 25 itself), alongside a "Christmas is coming!" row
 *    with Days Before = 3 (an earlier heads-up). Dates can also be:
 *      - Recurring Annually = TRUE  -> fires every year on that month/day
 *        (e.g. Christmas, New Year's)
 *      - Recurring Annually = FALSE -> fires once, on that exact date
 *        (e.g. a one-off promotion deadline)
 *
 * 2. MONTHLY CLIENT REMINDER (June through March)
 *    Fires on the 1st of every month, but only actually sends during
 *    June-March -- April and May are automatically skipped.
 *
 * ===========================================================
 * SETUP STEPS
 * ===========================================================
 * 1. Create a new Google Sheet (or add tabs to an existing one) with:
 *
 *    Tab "Clients" -- columns: Name | Email
 *      One row per client. Add/remove rows anytime, no code changes needed.
 *
 *    Tab "Holiday_Dates" -- columns: Date | Days Before | Recurring Annually | Subject | Body
 *      One row per holiday/specific date. "Days Before" is a number
 *      (0 = send on the date itself, 1 = day before, etc). "Recurring
 *      Annually" is a checkbox (TRUE/FALSE). Use {{name}} in Subject/Body
 *      for the client's first name.
 *
 *    Tab "Reminders_Config" -- columns: Type | Subject | Body
 *      One row with Type = Monthly for the recurring June-March email.
 *      (Named distinctly from "Message_Config" so it never collides with
 *      your onboarding script's tab of the same common name, if you're
 *      using the same spreadsheet for both.)
 *
 * 2. Paste this script in (Extensions > Apps Script, opened from inside
 *    the Sheet). Fill in SPREADSHEET_ID below.
 *
 * 3. Run setup() once. Installs both triggers together.
 *
 * 4. Test with testHolidayReminder() and testMonthlyReminder() -- both
 *    bypass their normal timing gates so you can preview content anytime.
 * ---------------------------------------------------------
 */

// ====================== CONFIG ======================

const CLIENTS_SHEET_NAME = 'Clients';
const HOLIDAY_SHEET_NAME = 'Holiday_Dates';
const CONFIG_SHEET_NAME = 'Reminders_Config';

const COL_CLIENT_NAME = 'Name';
const COL_CLIENT_EMAIL = 'Email';

const COL_HOLIDAY_DATE = 'Date';
const COL_HOLIDAY_DAYS_BEFORE = 'Days Before'; // 0 = send ON the date itself; 1+ = send that many days earlier
const COL_HOLIDAY_RECURRING = 'Recurring Annually';
const COL_HOLIDAY_SUBJECT = 'Subject';
const COL_HOLIDAY_BODY = 'Body';
const COL_HOLIDAY_LAST_SENT = 'Last Sent Date'; // auto-added tracking column

// Monthly reminder skips these months (1=Jan ... 12=Dec). April & May skipped
// so it effectively runs June through March.
const MONTHLY_SKIP_MONTHS = [4, 5];

// PASTE YOUR SPREADSHEET ID HERE (from the Sheet's URL, between /d/ and /edit)
const SPREADSHEET_ID = '1pMLeXyfYR4wncq7iZWDeE8wwGQRzsxSkO8blSZN5E7k';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Looks up a tab by name and throws a clear, specific error (listing the
 * tabs that actually exist) if it's not found -- instead of letting the
 * caller crash later with a generic "Cannot read properties of null" error.
 */
function getSheetOrThrow(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const names = ss.getSheets().map((s) => s.getName()).join(', ');
    throw new Error(
      `Could not find a tab named "${sheetName}". Tabs that actually exist in this spreadsheet: ${names}. ` +
      `Either rename your tab to match exactly (case-sensitive), or check that SPREADSHEET_ID at the top ` +
      `of the script points to the right spreadsheet.`
    );
  }
  return sheet;
}

// ====================== HOLIDAY / SPECIFIC DATE REMINDERS ======================

// Runs daily. Checks every row in Holiday_Dates -- if TOMORROW matches
// that row's date (recurring annually or exact one-time), sends the
// reminder to everyone in Clients.
function checkHolidayReminders() {
  const sheet = getSheetOrThrow(HOLIDAY_SHEET_NAME);
  ensureHolidayTrackingColumn(sheet);

  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const todayStr = formatDate(new Date());

  for (let row = 2; row <= lastRow; row++) {
    const dateValue = sheet.getRange(row, col(COL_HOLIDAY_DATE)).getValue();
    const daysBeforeRaw = sheet.getRange(row, col(COL_HOLIDAY_DAYS_BEFORE)).getValue();
    const daysBefore = (daysBeforeRaw === '' || daysBeforeRaw === null) ? 0 : Number(daysBeforeRaw);
    const recurring = sheet.getRange(row, col(COL_HOLIDAY_RECURRING)).getValue() === true;
    const subject = sheet.getRange(row, col(COL_HOLIDAY_SUBJECT)).getValue();
    const body = sheet.getRange(row, col(COL_HOLIDAY_BODY)).getValue();
    const lastSent = sheet.getRange(row, col(COL_HOLIDAY_LAST_SENT)).getValue();

    if (!dateValue || !subject || !body) continue;
    if (lastSent === todayStr) continue; // already sent today, avoid duplicate if trigger reruns

    const targetDate = addDays(new Date(), daysBefore);
    const matches = recurring
      ? isSameMonthDay(dateValue, targetDate)
      : isSameCalendarDate(dateValue, targetDate);

    if (!matches) continue;

    sendToAllClients(subject, body);
    sheet.getRange(row, col(COL_HOLIDAY_LAST_SENT)).setValue(todayStr);
    Logger.log('Holiday reminder sent for row ' + row + ' ("' + subject + '"), ' + daysBefore + ' day(s) before.');
  }
}

/**
 * TEST FUNCTION — sends one specific Holiday_Dates row's reminder right
 * now, bypassing the "1 day before" timing check entirely. Edit
 * rowToTest below, select "testHolidayReminder" from the function
 * dropdown, and click Run.
 */
function testHolidayReminder() {
  const rowToTest = 2; // change to the row number in Holiday_Dates you want to test

  const sheet = getSheetOrThrow(HOLIDAY_SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);

  const subject = sheet.getRange(rowToTest, col(COL_HOLIDAY_SUBJECT)).getValue();
  const body = sheet.getRange(rowToTest, col(COL_HOLIDAY_BODY)).getValue();

  if (!subject || !body) {
    Logger.log('Row ' + rowToTest + ' is missing Subject or Body.');
    return;
  }

  sendToAllClients(subject, body);
  Logger.log('Test holiday reminder sent for row ' + rowToTest + ' ("' + subject + '").');
}

// ====================== MONTHLY CLIENT REMINDER ======================

// Runs on the 1st of every month (see setup()). Skips MONTHLY_SKIP_MONTHS
// (April & May by default), so it effectively sends June through March.
function sendMonthlyClientReminder() {
  const currentMonth = new Date().getMonth() + 1; // 1-12
  if (MONTHLY_SKIP_MONTHS.indexOf(currentMonth) !== -1) {
    Logger.log('Monthly reminder skipped -- month ' + currentMonth + ' is in the skip list.');
    return;
  }

  const config = getMessageConfig('Monthly');
  sendToAllClients(config.subject, config.body);
  Logger.log('Monthly client reminder sent (month ' + currentMonth + ').');
}

/**
 * TEST FUNCTION — sends the Monthly reminder right now, bypassing the
 * month-skip check entirely. Select "testMonthlyReminder" from the
 * function dropdown and click Run.
 */
function testMonthlyReminder() {
  const config = getMessageConfig('Monthly');
  sendToAllClients(config.subject, config.body);
  Logger.log('Test monthly reminder sent.');
}

// ====================== SENDING ======================

// Sends the given subject/body (with {{name}} personalization) to every
// row in the Clients tab.
function sendToAllClients(rawSubject, rawBody) {
  const sheet = getSheetOrThrow(CLIENTS_SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No clients found in ' + CLIENTS_SHEET_NAME + '.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  data.forEach((row) => {
    const name = row[col(COL_CLIENT_NAME) - 1];
    const email = row[col(COL_CLIENT_EMAIL) - 1];
    if (!email) return;

    const firstName = name ? name.toString().trim().split(' ')[0] : 'there';
    const subject = applyPlaceholders(rawSubject, { name: firstName });
    const rawBodyText = applyPlaceholders(rawBody, { name: firstName });
    const htmlBody = buildEmailHtml(subject, textToHtml(rawBodyText));

    try {
      MailApp.sendEmail({
        to: email.toString().trim(),
        subject: subject,
        body: rawBodyText,
        htmlBody: htmlBody
      });
    } catch (err) {
      Logger.log('Failed to send to ' + email + ': ' + err);
    }
  });
}

// ====================== MESSAGE CONFIG ======================

function getMessageConfig(type) {
  const sheet = getSheetOrThrow(CONFIG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeCol = headers.indexOf('Type');
  const subjectCol = headers.indexOf('Subject');
  const bodyCol = headers.indexOf('Body');

  for (let r = 1; r < data.length; r++) {
    if (data[r][typeCol] && data[r][typeCol].toString().trim().toLowerCase() === type.toLowerCase()) {
      return { subject: data[r][subjectCol], body: data[r][bodyCol] };
    }
  }
  throw new Error('No row found in ' + CONFIG_SHEET_NAME + ' for Type = "' + type + '".');
}

function applyPlaceholders(text, values) {
  let result = text.toString();
  Object.keys(values).forEach((key) => {
    result = result.split('{{' + key + '}}').join(values[key]);
  });
  return result;
}

// ====================== HTML RENDERING (same style as your other scripts) ======================

function linkify(text) {
  const urlPattern = /(https?:\/\/[^\s<]+)/g;
  return text.replace(urlPattern, (url) => `<a href="${url}" style="color:#2A9D8F;">${url}</a>`);
}

function textToHtml(text) {
  const normalized = text.replace(/\n\s*\n+/g, '\n\n');
  const paragraphs = normalized.split(/\n\n/);
  return paragraphs
    .map((p) => p.trim())
    .filter((p) => p !== '')
    .map((p) => {
      const lines = p.split('\n');
      if (lines.every((l) => l.trim().startsWith('- '))) {
        return '<ul style="margin:0 0 16px 0;padding-left:20px;">' +
          lines.map((l) => `<li style="margin-bottom:8px;color:#3a3a3a;font-size:15px;line-height:1.5;">${linkify(l.trim().substring(2))}</li>`).join('') +
          '</ul>';
      }
      return `<p style="margin:0 0 16px 0;color:#3a3a3a;font-size:15px;line-height:1.6;">${linkify(lines.join('<br>'))}</p>`;
    })
    .join('');
}

function buildEmailHtml(subject, bodyHtml) {
  return `
<div style="background-color:#f4f4f4;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" style="max-width:560px;margin:0 auto;background-color:#ffffff;
    border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
    <tr>
      <td align="center" style="background-color:#1F3B57;padding:26px 24px 20px 24px;">
        <div style="color:#D9B84A;font-size:13px;font-weight:bold;letter-spacing:1px;text-transform:uppercase;">
          ⭐ Team Excellence ⭐
        </div>
        <div style="color:#ffffff;font-size:20px;font-weight:bold;margin-top:8px;">
          ${subject}
        </div>
      </td>
    </tr>
    <tr><td style="height:4px;background-color:#B8952B;"></td></tr>
    <tr>
      <td style="padding:26px 26px 20px 26px;">
        ${bodyHtml}
      </td>
    </tr>
  </table>
</div>`;
}

// ====================== DATE HELPERS ======================

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function isSameMonthDay(dateValue, compareDate) {
  const d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return false;
  return d.getMonth() === compareDate.getMonth() && d.getDate() === compareDate.getDate();
}

function isSameCalendarDate(dateValue, compareDate) {
  const d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return false;
  return d.getFullYear() === compareDate.getFullYear() &&
    d.getMonth() === compareDate.getMonth() &&
    d.getDate() === compareDate.getDate();
}

function formatDate(date) {
  const tz = getSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

// ====================== SETUP HELPERS ======================

function ensureHolidayTrackingColumn(sheet) {
  const headers = getHeaders(sheet);
  if (headers.indexOf(COL_HOLIDAY_LAST_SENT) === -1) {
    sheet.getRange(1, headers.length + 1).setValue(COL_HOLIDAY_LAST_SENT);
  }
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function findColumn(headers, name) {
  const target = name.toString().trim().toLowerCase();
  const idx = headers.findIndex((h) => h && h.toString().trim().toLowerCase() === target);
  if (idx === -1) {
    throw new Error(`Column "${name}" not found. Headers present: ${headers.join(' | ')}`);
  }
  return idx + 1;
}

// Run this ONCE -- installs both triggers together. Safe to re-run anytime;
// only clears its own previous checkHolidayReminders/sendMonthlyClientReminder
// triggers first.
function setup() {
  const ss = getSpreadsheet();
  ['Clients', 'Holiday_Dates', 'Reminders_Config'].forEach((name) => {
    if (!ss.getSheetByName(name)) {
      throw new Error(`Could not find a tab named "${name}". Create it before running setup().`);
    }
  });

  const holidaySheet = ss.getSheetByName(HOLIDAY_SHEET_NAME);
  ensureHolidayTrackingColumn(holidaySheet);

  const managedHandlers = ['checkHolidayReminders', 'sendMonthlyClientReminder'];
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (managedHandlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkHolidayReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  ScriptApp.newTrigger('sendMonthlyClientReminder')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('Setup complete. Daily holiday-check trigger + monthly (1st of month) trigger installed.');
}