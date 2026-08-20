/**
 * CLIENT REFERRAL FORM — Digest + Checklist System
 * ---------------------------------------------------------
 * WORKFLOW (v2)
 * 1. Data comes in however you like -- form submissions, or bulk-pasted
 *    rows from Excel. No per-row instant email is sent anymore (the old
 *    onFormSubmit trigger is removed).
 * 2. Every 3 days, each Agent who has pending referrals gets ONE digest
 *    email listing all their clients:
 *      - Not Contacted clients first
 *      - Contacted (but budget form not yet done) clients below, marked
 *        with a checkmark
 *    The email includes a single "Update My Referrals" button.
 * 3. Clicking that button opens a real webpage (not inside the email --
 *    email clients strip interactive forms, so this has to be a separate
 *    page) showing every referral with two checkboxes: Contacted, and
 *    Budget Form Uploaded. The agent can tick any/all of them, anytime,
 *    and click Save -- all updates apply at once.
 * 4. Any row where BOTH boxes are checked is automatically colored green
 *    in the sheet, and excluded from future digest emails (though it still
 *    shows on the checklist page, in its own "Completed" section).
 *
 * All wording lives in the Message_Config sheet tab.
 *
 * ===========================================================
 * SETUP STEPS
 * ===========================================================
 * 1. In Message_Config, make sure there's a row with Type = Digest.
 *    Subject/Body support {{agentName}}. The Not Contacted / Contacted
 *    lists and the Update button are appended automatically after your
 *    Body text.
 *
 * 2. Fill in SPREADSHEET_ID and BUDGET_FORM_URL below if not already set.
 *
 * 3. Deploy the Web App (needed for the checklist page):
 *    Deploy > New deployment > Web app > Execute as: Me > Access: Anyone
 *    Copy the URL into WEBAPP_URL below, save.
 *    IMPORTANT: if you ever edit doGet() later, use Deploy > Manage
 *    deployments > Edit > New version, or the checklist page keeps
 *    running old code.
 *
 * 4. Run setup() once. Installs ONLY the 3-day digest trigger now --
 *    there is no more form-submit trigger in this version.
 * ---------------------------------------------------------
 */

// ====================== CONFIG ======================

const SHEET_NAME = "Form Responses 1";
const CONFIG_SHEET_NAME = "Message_Config";

const COL_CLIENT_NAME = "Client Name";
const COL_CLIENT_TEL = "Client Tel";
const COL_CLIENT_EMAIL = "Client Email";
const COL_AGENT_NAME = "Agent Name";
const COL_AGENT_EMAIL = "Agent Email";

// Tracking columns (auto-added if missing)
const COL_CONTACTED = "Contacted";                 // boolean checkbox
const COL_BUDGET_UPLOADED = "Budget Form Uploaded"; // boolean checkbox

// ---- Optional: copy the uploaded file link from a separate "Budget
// Upload" Google Form into the main sheet, purely as reference info.
// This does NOT touch the Budget Form Uploaded checkbox or any digest/
// checklist logic -- the agent's manual tick remains the only source of
// truth for that. This just adds a link for convenience.
const BUDGET_UPLOAD_SHEET_NAME = "BUDGET_SHEET_NAME"; // tab where that form's responses land
const COL_BUDGET_LINK = "Budget Form Link"; // new column added to the main sheet

const DIGEST_INTERVAL_DAYS = 3;

const SPREADSHEET_ID = "1sWUEwLUgufGBvqkSHhIKVTLAqESFrkedOLSU9Posg4c";
const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbz7Hw4-PLlYF9uOnR3rlu38V6UF6667h4a2-A6cZ6VvKptoaVf8a9y1WDsr_jAUoLsq/exec";
const BUDGET_FORM_URL = "https://forms.gle/t6vvUxK89fqPeGPu5";

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ====================== BUDGET UPLOAD LINK COPY (additive, optional) ======================

// Fires when the separate Budget Upload form is submitted. Ignores any
// other tab. Purely copies the uploaded file link into the main sheet by
// matching client email -- does not touch Contacted/Budget Form Uploaded
// checkboxes or any digest/checklist behavior.
function onFormSubmit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== BUDGET_UPLOAD_SHEET_NAME) return;

  const row = e.range.getRow();
  const headers = getHeaders(sheet);
  const emailColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase().includes("email"));
  const linkColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase().includes("upload"));

  if (emailColIdx === -1 || linkColIdx === -1) {
    Logger.log("Could not find Email/Upload columns on \"" + BUDGET_UPLOAD_SHEET_NAME + "\". Headers found: " + headers.join(" | "));
    return;
  }

  const clientEmail = sheet.getRange(row, emailColIdx + 1).getValue().toString().trim();
  const fileLink = sheet.getRange(row, linkColIdx + 1).getValue().toString().trim();
  if (!clientEmail || !fileLink) return;

  copyBudgetLinkToMainSheet(clientEmail, fileLink);
}

// Copies the uploaded file link into every main-sheet row matching this
// client email. Adds the "Budget Form Link" column if it doesn't exist yet.
function copyBudgetLinkToMainSheet(clientEmail, fileLink) {
  const mainSheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureBudgetLinkColumn(mainSheet);

  const headers = getHeaders(mainSheet);
  const emailCol = findColumn(headers, COL_CLIENT_EMAIL);
  const linkCol = findColumn(headers, COL_BUDGET_LINK);

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return;

  const emails = mainSheet.getRange(2, emailCol, lastRow - 1, 1).getValues();
  let matchCount = 0;

  for (let i = 0; i < emails.length; i++) {
    const candidate = emails[i][0] ? emails[i][0].toString().trim().toLowerCase() : "";
    if (candidate !== clientEmail.trim().toLowerCase()) continue;
    mainSheet.getRange(i + 2, linkCol).setValue(fileLink);
    matchCount++;
  }

  if (matchCount > 0) {
    Logger.log("Copied budget form link for " + clientEmail + " into " + matchCount + " row(s) in " + SHEET_NAME + ".");
  } else {
    Logger.log("Budget upload from " + clientEmail + " didn't match any row in " + SHEET_NAME + ".");
  }
}

function ensureBudgetLinkColumn(sheet) {
  const headers = getHeaders(sheet);
  if (headers.indexOf(COL_BUDGET_LINK) === -1) {
    sheet.getRange(1, headers.length + 1).setValue(COL_BUDGET_LINK);
  }
}

/**
 * TEST FUNCTION — simulates a Budget Upload form submission for one
 * specific client email and link, without needing a real submission.
 * Edit the two values below, select "testBudgetLinkCopy" from the
 * function dropdown, and click Run.
 */
function testBudgetLinkCopy() {
  const clientEmailToTest = "henme3@yahoo.com";
  const fileLinkToTest = "https://drive.google.com/some-test-link";
  copyBudgetLinkToMainSheet(clientEmailToTest, fileLinkToTest);
}

// ====================== DIGEST EMAILS ======================

// Runs every DIGEST_INTERVAL_DAYS (see setup()). Groups all pending
// referrals by agent and sends each agent ONE summary email.
function sendDigestEmails() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureColumns(sheet);

  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  // Group pending rows by agent email
  const byAgent = {}; // agentEmail -> { agentName, notContacted: [...], contacted: [...] }

  data.forEach((row, idx) => {
    const clientNameRaw = row[col(COL_CLIENT_NAME) - 1];
    const clientTel = row[col(COL_CLIENT_TEL) - 1];
    const clientEmail = row[col(COL_CLIENT_EMAIL) - 1];
    const agentName = row[col(COL_AGENT_NAME) - 1];
    const agentEmail = row[col(COL_AGENT_EMAIL) - 1];
    const contacted = row[col(COL_CONTACTED) - 1] === true;
    const budgetUploaded = row[col(COL_BUDGET_UPLOADED) - 1] === true;

    if (!agentEmail) return;
    if (!clientNameRaw && !clientEmail) return; // need at least SOME identifier
    if (contacted && budgetUploaded) return; // fully done -- skip from digest

    const displayName = clientNameRaw ? clientNameRaw.toString().trim() : clientEmail.toString().trim();

    const key = agentEmail.toString().trim().toLowerCase();
    if (!byAgent[key]) {
      byAgent[key] = { agentEmail: agentEmail.toString().trim(), agentName: agentName, clients: [] };
    }
    byAgent[key].clients.push({
      name: displayName,
      tel: clientTel,
      email: clientEmail,
      contacted: contacted,
      budgetUploaded: budgetUploaded
    });
  });

  const config = getMessageConfig("Digest");

  Object.keys(byAgent).forEach((key) => {
    const agent = byAgent[key];
    const checklistUrl = WEBAPP_URL + "?agent=" + encodeURIComponent(buildAgentToken(agent.agentEmail));

    const subject = applyPlaceholders(config.subject, { agentName: agent.agentName });
    const introBody = applyPlaceholders(config.body, { agentName: agent.agentName });

    const htmlBody = buildDigestEmailHtml(introBody, agent.clients, checklistUrl);
    const plainBody = buildDigestPlainText(introBody, agent.clients, checklistUrl);

    MailApp.sendEmail({
      to: agent.agentEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });

    const notContactedCount = agent.clients.filter(c => !c.contacted).length;
    Logger.log("Digest sent to " + agent.agentEmail + " (" + notContactedCount + " not contacted, " + (agent.clients.length - notContactedCount) + " contacted)");
  });
}

/**
 * TEST FUNCTION — sends the digest email to ONE specific agent right now,
 * bypassing the 3-day schedule. Edit agentEmailToTest below, select
 * "testDigestForAgent" from the function dropdown, and click Run.
 */
function testDigestForAgent() {
  const agentEmailToTest = "munazzashaikh2000@gmail.com";

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  ensureColumns(sheet);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  let agentName = "";
  const clients = [];

  data.forEach((row) => {
    const rowAgentEmail = (row[col(COL_AGENT_EMAIL) - 1] || "").toString().trim().toLowerCase();
    if (rowAgentEmail !== agentEmailToTest.trim().toLowerCase()) return;

    const clientNameRaw = row[col(COL_CLIENT_NAME) - 1];
    const clientTel = row[col(COL_CLIENT_TEL) - 1];
    const clientEmail = row[col(COL_CLIENT_EMAIL) - 1];
    const isContacted = row[col(COL_CONTACTED) - 1] === true;
    const isBudgetDone = row[col(COL_BUDGET_UPLOADED) - 1] === true;
    agentName = row[col(COL_AGENT_NAME) - 1];

    if (!clientNameRaw && !clientEmail) return;
    if (isContacted && isBudgetDone) return; // fully done -- skip

    const displayName = clientNameRaw ? clientNameRaw.toString().trim() : clientEmail.toString().trim();
    clients.push({ name: displayName, tel: clientTel, email: clientEmail, contacted: isContacted, budgetUploaded: isBudgetDone });
  });

  if (!agentName && clients.length === 0) {
    Logger.log("No rows found for agent email: " + agentEmailToTest);
    return;
  }

  const config = getMessageConfig("Digest");
  const checklistUrl = WEBAPP_URL + "?agent=" + encodeURIComponent(buildAgentToken(agentEmailToTest));
  const subject = applyPlaceholders(config.subject, { agentName: agentName });
  const introBody = applyPlaceholders(config.body, { agentName: agentName });
  const htmlBody = buildDigestEmailHtml(introBody, clients, checklistUrl);
  const plainBody = buildDigestPlainText(introBody, clients, checklistUrl);

  MailApp.sendEmail({ to: agentEmailToTest, subject: subject, body: plainBody, htmlBody: htmlBody });
  Logger.log("Test digest sent to " + agentEmailToTest + " with " + clients.length + " client(s). Checklist URL: " + checklistUrl);
}

// ====================== CHECKLIST PAGE (Web App) ======================

// Handles VIEWING the checklist page only. Saving now happens via
// google.script.run (see submitChecklistUpdates below) instead of a
// second HTTP request -- this sidesteps all the GET/POST URL-length,
// encoding, and deployment-routing issues entirely, since
// google.script.run talks to the server through Apps Script's own
// internal bridge, not a public URL.
function doGet(e) {
  const token = e.parameter.agent;
  if (!token) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Missing agent link. Please use the link from your email.</p>');
  }

  let agentEmail;
  try {
    agentEmail = parseAgentToken(token);
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Invalid or corrupted link.</p>');
  }

  return renderChecklistPage(agentEmail);
}

/**
 * Called directly from the checklist page's JavaScript via google.script.run.
 * @param {string} token - the agent's token (same one embedded in the page)
 * @param {Array<{row:number, contacted:boolean, budgetUploaded:boolean}>} updates
 */
function submitChecklistUpdates(token, updates) {
  const agentEmail = parseAgentToken(token);
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);

  updates.forEach((u) => {
    const rowNum = u.row;
    // Defense-in-depth: confirm this row actually belongs to this agent
    // before applying any update, in case row numbers were tampered with.
    const rowAgentEmail = sheet.getRange(rowNum, col(COL_AGENT_EMAIL)).getValue().toString().trim().toLowerCase();
    if (rowAgentEmail !== agentEmail.trim().toLowerCase()) return;

    const contacted = !!u.contacted;
    const budgetUploaded = !!u.budgetUploaded;

    sheet.getRange(rowNum, col(COL_CONTACTED)).setValue(contacted);
    sheet.getRange(rowNum, col(COL_BUDGET_UPLOADED)).setValue(budgetUploaded);

    const rowRange = sheet.getRange(rowNum, 1, 1, headers.length);
    if (contacted && budgetUploaded) {
      rowRange.setBackground("#d9ead3"); // light green
    } else {
      rowRange.setBackground(null);
    }
  });

  Logger.log("Saved checklist updates for " + agentEmail + ": " + updates.length + " row(s)");
  return { success: true, count: updates.length };
}

function renderChecklistPage(agentEmail) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = getHeaders(sheet);
  const col = (name) => findColumn(headers, name);
  const lastRow = sheet.getLastRow();

  const notContacted = [];
  const contactedPending = [];
  const completed = [];
  const allRowNumbers = [];

  for (let r = 2; r <= lastRow; r++) {
    const rowAgentEmail = sheet.getRange(r, col(COL_AGENT_EMAIL)).getValue().toString().trim().toLowerCase();
    if (rowAgentEmail !== agentEmail.trim().toLowerCase()) continue;

    const clientName = sheet.getRange(r, col(COL_CLIENT_NAME)).getValue();
    if (!clientName) continue;

    const clientTel = sheet.getRange(r, col(COL_CLIENT_TEL)).getValue();
    const contacted = sheet.getRange(r, col(COL_CONTACTED)).getValue() === true;
    const budgetUploaded = sheet.getRange(r, col(COL_BUDGET_UPLOADED)).getValue() === true;

    allRowNumbers.push(r);
    const entry = { row: r, name: clientName, tel: clientTel, contacted: contacted, budgetUploaded: budgetUploaded };

    if (contacted && budgetUploaded) completed.push(entry);
    else if (contacted) contactedPending.push(entry);
    else notContacted.push(entry);
  }

  const rowCheckbox = (entry) => `
    <tr>
      <td style="padding:10px 12px;border-bottom:1px solid #eee;">${entry.name}${entry.tel ? ' <span style="color:#999;font-size:12px;">(' + entry.tel + ')</span>' : ''}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #eee;text-align:center;">
        <input type="checkbox" id="c_${entry.row}" data-row="${entry.row}" class="row-checkbox" ${entry.contacted ? "checked" : ""} style="width:18px;height:18px;">
      </td>
      <td style="padding:10px 12px;border-bottom:1px solid #eee;text-align:center;">
        <input type="checkbox" id="b_${entry.row}" data-row="${entry.row}" class="row-checkbox" ${entry.budgetUploaded ? "checked" : ""} style="width:18px;height:18px;">
      </td>
    </tr>`;

  const sectionRows = (title, entries, color) => {
    if (entries.length === 0) return "";
    return `
      <tr><td colspan="3" style="padding:16px 12px 6px 12px;font-weight:bold;color:${color};">${title}</td></tr>
      ${entries.map(rowCheckbox).join("")}
    `;
  };

  const tableRows =
    sectionRows("Not Contacted", notContacted, "#B8952B") +
    sectionRows("Contacted — Budget Form Pending", contactedPending, "#1F3B57") +
    sectionRows("Completed", completed, "#2E7D32");

  const html = `
<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:24px auto;padding:0 16px;">
  <div id="statusBanner"></div>
  <h2 style="color:#1F3B57;">Your Client Referrals</h2>
  <p style="color:#555;font-size:14px;">Tick the boxes below as you go, then click Save. You can come back and update this anytime.</p>
  <p style="color:#555;font-size:13px;">Budget / Fundamentals form to send clients: <a href="${BUDGET_FORM_URL}" target="_blank">${BUDGET_FORM_URL}</a></p>

  <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">
    <tr style="background:#1F3B57;color:#fff;">
      <td style="padding:10px 12px;font-weight:bold;">Client</td>
      <td style="padding:10px 12px;font-weight:bold;text-align:center;">Contacted</td>
      <td style="padding:10px 12px;font-weight:bold;text-align:center;">Budget Form Uploaded</td>
    </tr>
    ${tableRows || '<tr><td colspan="3" style="padding:20px;text-align:center;color:#999;">No referrals found.</td></tr>'}
  </table>

  ${allRowNumbers.length > 0 ? `
  <div style="text-align:center;margin-top:20px;">
    <button id="saveBtn" style="background:#B8952B;color:#fff;border:none;padding:12px 32px;
      border-radius:6px;font-size:15px;font-weight:bold;cursor:pointer;">
      Save Updates
    </button>
  </div>` : ""}
</div>

<script>
  var AGENT_TOKEN = "${buildAgentToken(agentEmail)}";
  var ALL_ROWS = [${allRowNumbers.join(",")}];

  var saveBtn = document.getElementById("saveBtn");
  if (saveBtn) {
    saveBtn.addEventListener("click", function() {
      var updates = ALL_ROWS.map(function(r) {
        return {
          row: r,
          contacted: document.getElementById("c_" + r).checked,
          budgetUploaded: document.getElementById("b_" + r).checked
        };
      });

      saveBtn.disabled = true;
      saveBtn.innerText = "Saving...";

      google.script.run
        .withSuccessHandler(function() {
          document.getElementById("statusBanner").innerHTML =
            '<div style="background:#d9ead3;color:#2E7D32;padding:12px 16px;border-radius:6px;margin-bottom:16px;font-weight:bold;">✅ Saved!</div>';
          saveBtn.disabled = false;
          saveBtn.innerText = "Save Updates";
          window.scrollTo(0, 0);
        })
        .withFailureHandler(function(err) {
          document.getElementById("statusBanner").innerHTML =
            '<div style="background:#fdecea;color:#c62828;padding:12px 16px;border-radius:6px;margin-bottom:16px;font-weight:bold;">Error saving: ' + err.message + '</div>';
          saveBtn.disabled = false;
          saveBtn.innerText = "Save Updates";
          window.scrollTo(0, 0);
        })
        .submitChecklistUpdates(AGENT_TOKEN, updates);
    });
  }
</script>`;

  return HtmlService.createHtmlOutput(html).setTitle("Your Client Referrals");
}

// ====================== AGENT TOKEN (stateless email encoding) ======================

function buildAgentToken(agentEmail) {
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(agentEmail.trim().toLowerCase()).getBytes());
}

function parseAgentToken(token) {
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
}

// ====================== MESSAGE CONFIG ======================

function getMessageConfig(type) {
  const sheet = getSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Could not find a tab named "' + CONFIG_SHEET_NAME + '". Create it with columns Type | Subject | Body.');
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
  throw new Error('No row found in ' + CONFIG_SHEET_NAME + ' for Type = "' + type + '".');
}

function applyPlaceholders(text, values) {
  let result = text;
  Object.keys(values).forEach(function(key) {
    result = result.split("{{" + key + "}}").join(values[key] != null ? values[key] : "");
  });
  return result;
}

// ====================== EMAIL RENDERING ======================

function textToHtml(text) {
  const normalized = text.replace(/\n\s*\n+/g, "\n\n");
  return normalized.split(/\n\n/)
    .map(p => p.trim())
    .filter(p => p !== "")
    .map(p => `<p style="margin:0 0 14px 0;color:#333333;font-size:15px;line-height:1.6;">${p.split("\n").join("<br>")}</p>`)
    .join("");
}

function buildDigestEmailHtml(introBody, clients, checklistUrl) {
  // Not contacted first, then contacted (budget-pending) -- both always visible
  const sorted = clients.slice().sort((a, b) => (a.contacted === b.contacted) ? 0 : (a.contacted ? 1 : -1));

  const statusBadge = (isDone, doneLabel, pendingLabel) => isDone
    ? `<span style="color:#2E7D32;font-weight:bold;">✅ ${doneLabel}</span>`
    : `<span style="color:#B8952B;font-weight:bold;">⏳ ${pendingLabel}</span>`;

  const row = (c, i) => `
    <tr style="background:${i % 2 === 0 ? '#ffffff' : '#f9f7f2'};">
      <td style="padding:10px 12px;border-bottom:1px solid #eee;font-size:13px;">
        <div style="font-weight:bold;color:#333;">${c.name}</div>
        ${c.tel ? `<div style="color:#888;font-size:12px;">📞 ${c.tel}</div>` : ""}
        ${c.email ? `<div style="color:#888;font-size:12px;">✉️ ${c.email}</div>` : ""}
      </td>
      <td style="padding:10px 12px;border-bottom:1px solid #eee;font-size:13px;text-align:center;">
        ${statusBadge(c.contacted, "Contacted", "Not Contacted")}
      </td>
      <td style="padding:10px 12px;border-bottom:1px solid #eee;font-size:13px;text-align:center;">
        ${c.contacted ? statusBadge(c.budgetUploaded, "Budget Uploaded", "Budget Pending") : '<span style="color:#bbb;">—</span>'}
      </td>
    </tr>`;

  const notContactedCount = clients.filter(c => !c.contacted).length;
  const contactedCount = clients.length - notContactedCount;

  const tableHtml = clients.length === 0 ? '<p style="color:#999;font-size:14px;">No pending referrals right now.</p>' : `
    <table role="presentation" width="100%" style="border-collapse:collapse;margin-top:4px;">
      <tr style="background:#1F3B57;color:#fff;">
        <td style="padding:10px 12px;font-weight:bold;font-size:13px;">Client</td>
        <td style="padding:10px 12px;font-weight:bold;font-size:13px;text-align:center;">Contacted</td>
        <td style="padding:10px 12px;font-weight:bold;font-size:13px;text-align:center;">Budget Form</td>
      </tr>
      ${sorted.map(row).join("")}
    </table>`;

  return `
<div style="background-color:#f4f4f4;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" style="max-width:600px;margin:0 auto;background-color:#ffffff;
    border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
    <tr>
      <td align="center" style="background-color:#1F3B57;padding:24px;">
        <div style="color:#ffffff;font-size:19px;font-weight:bold;">Your Referral Update</div>
      </td>
    </tr>
    <tr><td style="height:4px;background-color:#B8952B;"></td></tr>
    <tr>
      <td style="padding:24px 24px 8px 24px;">
        ${textToHtml(introBody)}
        <p style="color:#555;font-size:13px;margin:0 0 14px 0;">
          <strong>${notContactedCount}</strong> not contacted &nbsp;•&nbsp; <strong>${contactedCount}</strong> contacted
        </p>
        ${tableHtml}
        <p style="color:#555;font-size:13px;margin:16px 0 0 0;">
          Budget / Fundamentals form to send clients: <a href="${BUDGET_FORM_URL}" style="color:#2A6F97;">${BUDGET_FORM_URL}</a>
        </p>
        <div style="text-align:center;padding:22px 0 8px 0;">
          <a href="${checklistUrl}" style="background-color:#B8952B;color:#ffffff;text-decoration:none;
            font-size:15px;font-weight:bold;padding:12px 28px;border-radius:6px;display:inline-block;">
            Update My Referrals
          </a>
        </div>
      </td>
    </tr>
  </table>
</div>`;
}

function buildDigestPlainText(introBody, clients, checklistUrl) {
  const sorted = clients.slice().sort((a, b) => (a.contacted === b.contacted) ? 0 : (a.contacted ? 1 : -1));
  const notContactedCount = clients.filter(c => !c.contacted).length;
  const contactedCount = clients.length - notContactedCount;

  let text = introBody + "\n\n";
  text += notContactedCount + " not contacted, ";

  sorted.forEach((c) => {
    text += "- " + c.name;
    if (c.tel) text += " | " + c.tel;
    if (c.email) text += " | " + c.email;
    text += "\n";
    text += "  Contacted: " + (c.contacted ? "YES" : "NO");
    if (c.contacted) text += " | Budget Form: " + (c.budgetUploaded ? "UPLOADED" : "PENDING");
    text += "\n\n";
  });

  text += "Budget / Fundamentals form to send clients: " + BUDGET_FORM_URL + "\n\n";
  text += "Update your referrals here: " + checklistUrl;
  return text;
}

// ====================== TRACKING / HELPERS ======================

function ensureColumns(sheet) {
  const headers = getHeaders(sheet);
  const toAdd = [COL_CONTACTED, COL_BUDGET_UPLOADED].filter(h => headers.indexOf(h) === -1);
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
    throw new Error('Column "' + name + '" not found. Headers present: ' + headers.join(' | '));
  }
  return idx + 1;
}

// ====================== SETUP ======================

// Run this ONCE -- installs the 3-day digest trigger, plus a form-submit
// trigger that ONLY reacts to the separate Budget Upload form (for
// copying the file link into the main sheet -- see the section above).
// "Budget Form Uploaded" itself is still confirmed manually by the agent
// on the checklist page -- that remains the source of truth, unaffected
// by any of this.
function setup() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const names = ss.getSheets().map(s => s.getName()).join(", ");
    throw new Error('Could not find a tab named "' + SHEET_NAME + '". Tabs found: ' + names);
  }
  ensureColumns(sheet);

  const managedHandlers = ["sendDigestEmails", "onFormSubmit"];
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (managedHandlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendDigestEmails")
    .timeBased()
    .everyDays(DIGEST_INTERVAL_DAYS)
    .atHour(9)
    .create();

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("Setup complete. Digest trigger (every " + DIGEST_INTERVAL_DAYS + " days) + Budget Upload link-copy trigger installed.");
}