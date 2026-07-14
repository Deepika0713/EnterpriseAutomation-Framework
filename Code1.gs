/**
 * ============================================================
 *  Google Sheets Email Automation
 *  Version: 2.0.0  — Corrections applied
 * ============================================================
 */

const SHEET_RECIPIENTS  = 'Recipients';
const SHEET_TEMPLATE    = 'Template';
const SHEET_WRITEUP     = 'Write-up';
const STATUS_SENT       = 'Sent';
const COL_STATUS        = 3;
const COL_NAME          = 1;
const COL_EMAIL         = 2;


// ── UI Safety Helpers ─────────────────────────────────────────

function getUiSafe() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

function safeAlert(title, message) {
  const ui = getUiSafe();
  if (ui) {
    ui.alert(title, message, ui.ButtonSet.OK);
  } else {
    Logger.log('[ALERT] ' + title + ': ' + message);
  }
}

function safeConfirm(title, message) {
  const ui = getUiSafe();
  if (ui) {
    return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES;
  }
  Logger.log('[AUTO-CONFIRM] ' + title);
  return true;
}


// ── Menu ──────────────────────────────────────────────────────

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📧 Email Automation')
      .addItem('⚙️  Setup Sheets',           'setupSheets')
      .addItem('📤 Send Emails',             'sendEmails')
      .addItem('🔄 Clear Sent Status',       'clearSentStatus')
      .addSeparator()
      .addItem('🗑️  Empty Recipients',        'emptyRecipients')
      .addItem('🧹 Empty Template',          'emptyTemplate')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen: no UI context — ' + e.message);
  }
}


// ── 1. Setup Sheets ───────────────────────────────────────────
// FIX #3: Always writes headers regardless of existing data
// FIX #4: No sample recipients added
// FIX #5: Write-up sheet is created blank — never auto-populated
// FIX #6: Template always restores correct structure

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Recipients ────────────────────────────────────────────
  let recipientsSheet = ss.getSheetByName(SHEET_RECIPIENTS);
  if (!recipientsSheet) {
    recipientsSheet = ss.insertSheet(SHEET_RECIPIENTS);
    Logger.log('Created sheet: ' + SHEET_RECIPIENTS);
  }

  // ALWAYS write headers to row 1 — works even if sheet already has data
  recipientsSheet.getRange(1, 1, 1, 3)
    .setValues([['Name', 'Email', 'Status']])
    .setFontWeight('bold')
    .setBackground('#4A86E8')
    .setFontColor('#FFFFFF');

  recipientsSheet.autoResizeColumns(1, 3);
  Logger.log('Recipients headers set.');

  // ── Template ──────────────────────────────────────────────
  let templateSheet = ss.getSheetByName(SHEET_TEMPLATE);
  if (!templateSheet) {
    templateSheet = ss.insertSheet(SHEET_TEMPLATE);
    Logger.log('Created sheet: ' + SHEET_TEMPLATE);
  }

  // ALWAYS write headers to row 1
  templateSheet.getRange(1, 1, 1, 2)
    .setValues([['Field', 'Value']])
    .setFontWeight('bold')
    .setBackground('#4A86E8')
    .setFontColor('#FFFFFF');

  // ALWAYS ensure "Subject" label exists in row 2 col A
  // Only set default Value if col B is currently empty (preserve user edits)
  templateSheet.getRange(2, 1).setValue('Subject');
  if (!templateSheet.getRange(2, 2).getValue()) {
    templateSheet.getRange(2, 2).setValue('Welcome {{Name}}');
  }

  // ALWAYS ensure "Body" label exists in row 3 col A
  templateSheet.getRange(3, 1).setValue('Body');
  if (!templateSheet.getRange(3, 2).getValue()) {
    templateSheet.getRange(3, 2).setValue(
      'Hello {{Name}},\n\n' +
      'Welcome to our internship program.\n\n' +
      'Best regards,\nThe Internship Team'
    );
  }

  templateSheet.getRange(3, 2).setWrap(true);
  templateSheet.setColumnWidth(2, 400);
  templateSheet.autoResizeColumn(1);
  Logger.log('Template structure ensured.');

  // ── Write-up ──────────────────────────────────────────────
  // FIX #5: Only CREATE the sheet if missing — never write any content
  let writeupSheet = ss.getSheetByName(SHEET_WRITEUP);
  if (!writeupSheet) {
    writeupSheet = ss.insertSheet(SHEET_WRITEUP);
    Logger.log('Created blank Write-up sheet.');
  } else {
    Logger.log('Write-up sheet already exists — left untouched.');
  }

  ss.setActiveSheet(recipientsSheet);

  safeAlert(
    '✅ Setup Complete',
    'All 3 sheets are ready.\n\n' +
    '• Recipients — add your real recipients manually\n' +
    '• Template   — customize Subject & Body as needed\n' +
    '• Write-up   — About the automation tools and structure (200+ words required)'
  );
}


// ── 2. Send Emails ────────────────────────────────────────────

function sendEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const recipientsSheet = ss.getSheetByName(SHEET_RECIPIENTS);
  const templateSheet   = ss.getSheetByName(SHEET_TEMPLATE);

  if (!recipientsSheet || !templateSheet) {
    safeAlert('⚠️ Missing Sheets', 'Please run "Setup Sheets" first.');
    return;
  }

  // Read Subject and Body from Template sheet (never hardcoded)
  const templateData = templateSheet.getDataRange().getValues();
  let subjectTemplate = '';
  let bodyTemplate    = '';

  for (let i = 1; i < templateData.length; i++) {
    const field = String(templateData[i][0]).trim().toLowerCase();
    const value = String(templateData[i][1]).trim();
    if (field === 'subject') subjectTemplate = value;
    if (field === 'body')    bodyTemplate    = value;
  }

  if (!subjectTemplate || !bodyTemplate) {
    safeAlert(
      '⚠️ Incomplete Template',
      'The Template sheet must have both a "Subject" row and a "Body" row.'
    );
    return;
  }

  const lastRow = recipientsSheet.getLastRow();
  if (lastRow < 2) {
    safeAlert('ℹ️ No Recipients', 'The Recipients sheet has no data rows.');
    return;
  }

  const rows = recipientsSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  const pendingCount = rows.filter(r =>
    String(r[COL_STATUS - 1]).trim() !== STATUS_SENT &&
    String(r[COL_EMAIL  - 1]).trim() !== ''
  ).length;

  if (pendingCount === 0) {
    safeAlert('ℹ️ Nothing to Send', 'All rows are already marked "Sent" or have no email address.');
    return;
  }

  const confirmed = safeConfirm(
    '📤 Confirm Send',
    `You are about to send emails to ${pendingCount} recipient(s).\n\nContinue?`
  );
  if (!confirmed) return;

  let sentCount = 0, skipCount = 0, errorCount = 0;
  const errors = [];

  for (let i = 0; i < rows.length; i++) {
    const name   = String(rows[i][COL_NAME   - 1]).trim();
    const email  = String(rows[i][COL_EMAIL  - 1]).trim();
    const status = String(rows[i][COL_STATUS - 1]).trim();

    if (status === STATUS_SENT)   { skipCount++; continue; }
    if (!email) {
      errors.push(`Row ${i + 2}: Missing email for "${name || '(no name)'}" — skipped.`);
      errorCount++;
      continue;
    }

    const personalSubject = subjectTemplate.replace(/\{\{Name\}\}/g, name);
    const personalBody    = bodyTemplate.replace(/\{\{Name\}\}/g, name);

    try {
      GmailApp.sendEmail(email, personalSubject, personalBody);

      recipientsSheet.getRange(i + 2, COL_STATUS)
        .setValue(STATUS_SENT)
        .setBackground('#B7E1CD')
        .setFontWeight('bold');

      sentCount++;
      Logger.log('✅ Sent to ' + name + ' <' + email + '>');

    } catch (err) {
      Logger.log('❌ ERROR: ' + email + ' — ' + err.message);
      errors.push(`Row ${i + 2} (${name} – ${email}): ${err.message}`);
      errorCount++;
    }

    if (sentCount > 0 && sentCount % 20 === 0) Utilities.sleep(2000);
  }

  let summary =
    `✅ Sent:     ${sentCount}\n` +
    `⏭️  Skipped:  ${skipCount} (already sent)\n` +
    `❌ Errors:   ${errorCount}`;

  if (errors.length > 0) summary += '\n\nErrors:\n' + errors.join('\n');
  safeAlert('📊 Send Complete', summary);
}


// ── 3. Clear Sent Status ──────────────────────────────────────

function clearSentStatus() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPIENTS);

  if (!sheet) { safeAlert('⚠️ Sheet Not Found', 'Run "Setup Sheets" first.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { safeAlert('ℹ️ Nothing to Clear', 'No data rows found.'); return; }

  if (!safeConfirm('🔄 Confirm Clear', 'Reset the Status column for all recipients?')) return;

  sheet.getRange(2, COL_STATUS, lastRow - 1, 1)
    .clearContent()
    .setBackground(null)
    .setFontWeight('normal');

  safeAlert('✅ Done', 'All Status values have been cleared.');
}


// ── 4a. Empty Recipients Only ─────────────────────────────────

function emptyRecipients() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPIENTS);

  if (!sheet) {
    safeAlert('⚠️ Sheet Not Found', 'Run "Setup Sheets" first.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    safeAlert('ℹ️ Already Empty', 'No recipient data rows to remove.');
    return;
  }

  if (!safeConfirm(
    '🗑️ Empty Recipients',
    `This will delete all ${lastRow - 1} recipient data row(s).\n\nHeaders will be kept.\nTemplate will NOT be affected.\n\nContinue?`
  )) return;

  sheet.deleteRows(2, lastRow - 1);

  safeAlert(
    '✅ Done',
    'All recipient data has been removed.\nHeaders are intact.\nTemplate was not touched.'
  );
}


// ── 4b. Empty Template Only ───────────────────────────────────

function emptyTemplate() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TEMPLATE);

  if (!sheet) {
    safeAlert('⚠️ Sheet Not Found', 'Run "Setup Sheets" first.');
    return;
  }

  if (!safeConfirm(
    '🧹 Empty Template',
    'This will clear the Subject and Body values from the Template sheet.\n\nField labels and headers will be kept.\nRecipients will NOT be affected.\n\nContinue?'
  )) return;

  sheet.getRange(2, 2).clearContent(); // Clear Subject value
  sheet.getRange(3, 2).clearContent(); // Clear Body value

  safeAlert(
    '✅ Done',
    'Template Subject and Body have been cleared.\nField labels and headers are intact.\nRecipients were not touched.'
  );
}
