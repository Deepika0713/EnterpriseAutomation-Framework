/**
 * ════════════════════════════════════════════════════════════════
 *  DYNAMIC EMAIL TEMPLATE FRAMEWORK
 *  v2.0.0  —  Parse any template → Auto-generate columns → Send
 *  Continuation from v4.0.0 | Task 3 Upgrade
 * ════════════════════════════════════════════════════════════════
 */

'use strict';

// ══════════════════════════════════════════════════════════════
//  CONFIG
// ══════════════════════════════════════════════════════════════
const CONFIG = Object.freeze({
  SHEET_TEMPLATE  : 'Email_Template',
  SHEET_RECIPIENTS: 'Recipients',
  SHEET_WRITEUP   : 'Write-up',

  SUBJECT_ROW: 2,
  BODY_ROW   : 3,
  COL_FIELD  : 1,
  COL_VALUE  : 2,

  REC_NAME  : 'Name',
  REC_EMAIL : 'Email',
  REC_STATUS: 'Status',

  STATUS_SENT   : 'Sent',
  STATUS_FAILED : 'FAILED',
  STATUS_INVALID: 'INVALID EMAIL',

  EMAIL_REGEX  : /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/,
  SEND_DELAY_MS: 250,
});


// ══════════════════════════════════════════════════════════════
//  MENU
// ══════════════════════════════════════════════════════════════
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📧 Email Automation')
      .addItem('⚙️  Setup Sheets',                  'setupSheets')
      .addSeparator()
      .addItem('🔍 Parse Template → Build Columns', 'parseTemplateAndBuildColumns')
      .addSeparator()
      .addItem('✅ Validate Recipients',             'validateRecipients')
      .addItem('👁️  Preview Next Email',            'previewNextEmail')
      .addItem('📧 Send Emails',                    'sendEmails')
      .addItem('🔄 Clear Sent Status',              'clearSentStatus')
      .addSeparator()
      .addItem('🧹 Clear Template',                 'emptyTemplate')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen: ' + e.message);
  }
}


// ══════════════════════════════════════════════════════════════
//  UI HELPERS
// ══════════════════════════════════════════════════════════════
function _ui() {
  try { return SpreadsheetApp.getUi(); } catch (_) { return null; }
}

function safeAlert(title, msg) {
  const ui = _ui();
  if (ui) ui.alert(title, msg, ui.ButtonSet.OK);
  else    Logger.log(`[ALERT] ${title}: ${msg}`);
}

function safeConfirm(title, msg) {
  const ui = _ui();
  if (!ui) { Logger.log(`[AUTO-YES] ${title}`); return true; }
  return ui.alert(title, msg, ui.ButtonSet.YES_NO) === ui.Button.YES;
}


// ══════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ══════════════════════════════════════════════════════════════
function _getSheet(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error(`Sheet "${name}" not found. Run ⚙️ Setup Sheets first.`);
  return s;
}

function _getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function _styleHeader(range) {
  range.setFontWeight('bold').setBackground('#4A86E8').setFontColor('#FFFFFF');
}


// ══════════════════════════════════════════════════════════════
//  TEMPLATE READER
// ══════════════════════════════════════════════════════════════
function _getTemplates() {
  const sheet   = _getSheet(CONFIG.SHEET_TEMPLATE);
  const subject = sheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE).getValue().toString().trim();
  const body    = sheet.getRange(CONFIG.BODY_ROW,    CONFIG.COL_VALUE).getValue().toString().trim();
  if (!subject) throw new Error('Subject (B2) is empty.\nOpen Email_Template and type your subject in the yellow cell.');
  if (!body)    throw new Error('Body (B3) is empty.\nOpen Email_Template and type your email body in the yellow cell.');
  return { subject, body };
}


// ══════════════════════════════════════════════════════════════
//  PLACEHOLDER RENDERER
//  Works for any {{variable}} — zero hardcoding needed
// ══════════════════════════════════════════════════════════════
function _render(template, data) {
  return template.replace(/\{\{\s*(.+?)\s*\}\}/g, (match, rawKey) => {
    const key = rawKey.trim();
    const val = data[key];

    if (val === undefined) {
      Logger.log(`⚠️ {{${key}}} — no matching column.`);
      return match;
    }

    // ✅ FIX: Detect Date objects and format cleanly
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy');
    }

    if (typeof val === 'number') {
      return Number.isInteger(val) ? String(val) : val.toFixed(2);
    }

    return String(val);
  });
}


// ══════════════════════════════════════════════════════════════
//  1. SETUP SHEETS
// ══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Email_Template (visual draft form) ─────────────────────
  const tplSheet = _getOrCreate(ss, CONFIG.SHEET_TEMPLATE);
  tplSheet.setColumnWidth(CONFIG.COL_FIELD, 90);
  tplSheet.setColumnWidth(CONFIG.COL_VALUE, 520);

  // Row 1 — Title banner
  tplSheet.getRange(1, 1, 1, 2).merge()
    .setValue('📧  Email Draft  —  Type Subject in B2 and Body in B3, then click "🔍 Parse Template → Build Columns"')
    .setFontWeight('bold').setFontSize(11)
    .setBackground('#E8F0FE').setFontColor('#1A73E8')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  tplSheet.setRowHeight(1, 40);

  // Row 2 — Subject
  tplSheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_FIELD)
    .setValue('Subject').setFontWeight('bold')
    .setBackground('#F1F3F4').setFontColor('#444444')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  tplSheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE)
    .setBackground('#FFFDE7')
    .setBorder(true, true, true, true, null, null, '#FBBC04', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setNote(
      '✏️ Type your email subject here.\n\n' +
      'Use {{VariableName}} for personalization.\n\n' +
      'Examples:\n' +
      '  Welcome to {{City}}, {{Name}}!\n' +
      '  Your {{Course}} offer letter is ready\n' +
      '  Internship at {{Company}} — {{Name}}\n\n' +
      '💡 After typing, click:\n' +
      '"🔍 Parse Template → Build Columns"'
    );
  tplSheet.setRowHeight(CONFIG.SUBJECT_ROW, 40);

  // Row 3 — Body
  tplSheet.getRange(CONFIG.BODY_ROW, CONFIG.COL_FIELD)
    .setValue('Body').setFontWeight('bold')
    .setBackground('#F1F3F4').setFontColor('#444444')
    .setHorizontalAlignment('right').setVerticalAlignment('top');
  tplSheet.getRange(CONFIG.BODY_ROW, CONFIG.COL_VALUE)
    .setBackground('#FFFDE7').setWrap(true).setVerticalAlignment('top')
    .setBorder(true, true, true, true, null, null, '#FBBC04', SpreadsheetApp.BorderStyle.SOLID)
    .setNote(
      '✏️ Type your full email body here.\n\n' +
      'Use {{VariableName}} for ANY custom field.\n' +
      'No limit on number of variables!\n\n' +
      '── Example ─────────────────────────\n' +
      'Dear {{Name}},\n\n' +
      'Your {{Course}} batch starts on {{Start Date}}\n' +
      'at {{Venue}}, {{City}}.\n\n' +
      'Please bring your {{Document}} on day one.\n\n' +
      'Regards,\n{{Coordinator}}\n' +
      '────────────────────────────────────\n\n' +
      '💡 Every {{Variable}} here becomes a\n' +
      'column in the Recipients sheet automatically!'
    );
  tplSheet.setRowHeight(CONFIG.BODY_ROW, 260);

  // Row 4 — Tip footer
  tplSheet.getRange(4, 1, 1, 2).merge()
    .setValue('💡  Any {{Variable}} you type = one column auto-created in Recipients sheet  |  No code changes needed for any template')
    .setFontStyle('italic').setFontSize(10)
    .setFontColor('#888888').setBackground('#FAFAFA')
    .setHorizontalAlignment('center');
  tplSheet.setRowHeight(4, 28);

  // ── Recipients (placeholder until parsed) ─────────────────
  const recSheet = _getOrCreate(ss, CONFIG.SHEET_RECIPIENTS);
  if (recSheet.getLastRow() === 0) {
    recSheet.getRange('A1')
      .setValue('⚡ This sheet is auto-generated. Type your template in Email_Template → click "🔍 Parse Template → Build Columns" → columns appear here automatically.')
      .setFontStyle('italic').setFontColor('#888888').setWrap(true);
    recSheet.setColumnWidth(1, 640);
    recSheet.setRowHeight(1, 52);
  }

  // ── Write-up (blank — never touched by script) ─────────────
  _getOrCreate(ss, CONFIG.SHEET_WRITEUP);

  ss.setActiveSheet(tplSheet);

  safeAlert('✅ Setup Complete',
    '3 sheets ready:\n\n' +
    '• Email_Template  — type Subject in B2, Body in B3\n' +
    '  Use {{AnyVariable}} — no limit!\n\n' +
    '• Recipients      — auto-built after parsing\n\n' +
    '• Write-up        — your personal reflection\n\n' +
    'Next: type your template then click\n' +
    '"🔍 Parse Template → Build Columns"'
  );
}


// ══════════════════════════════════════════════════════════════
//  2. PARSE TEMPLATE → BUILD COLUMNS   ★ Core of Task 3 ★
//
//  Reads Subject + Body from Email_Template.
//  Scans for every {{variable}} using regex.
//  Rebuilds Recipients sheet with exactly those columns.
//  Fixed columns: Name (first) | Email (second) | Status (last).
//  Dynamic columns: everything else — in order of appearance.
//  Preserves existing recipient data where column names match.
// ══════════════════════════════════════════════════════════════
function parseTemplateAndBuildColumns() {
  try {
    const { subject, body } = _getTemplates();
    const fullText          = subject + ' ' + body;

    // ── Extract all {{variable}} names (preserve order, dedupe) ─
    const regex   = /\{\{\s*(.+?)\s*\}\}/g;
    const seen    = new Set();
    const allVars = [];
    let   m;
    while ((m = regex.exec(fullText)) !== null) {
      const key = m[1].trim();
      if (!seen.has(key)) { seen.add(key); allVars.push(key); }
    }

    if (allVars.length === 0) {
      safeAlert('⚠️ No Placeholders Found',
        'No {{variable}} placeholders found in your template.\n\n' +
        'Go to Email_Template and add placeholders like:\n' +
        '  {{Name}}, {{City}}, {{Course}}, {{Start Date}}\n\n' +
        'Then try again.'
      );
      return;
    }

    // ── Separate fixed vs dynamic columns ────────────────────
    // Name + Email → always fixed first two columns
    // Status       → always last (auto-managed)
    // Everything else → dynamic middle columns
    const fixedFirst  = [CONFIG.REC_NAME, CONFIG.REC_EMAIL];
    const fixedLast   = [CONFIG.REC_STATUS];
    const dynamicCols = allVars.filter(v =>
      !fixedFirst.includes(v) && !fixedLast.includes(v)
    );
    const newHeaders = [...fixedFirst, ...dynamicCols, ...fixedLast];

    // ── Snapshot existing data before clearing ────────────────
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const recSheet = _getOrCreate(ss, CONFIG.SHEET_RECIPIENTS);
    let existingHeaders = [];
    let existingData    = [];
    const lastRow = recSheet.getLastRow();

    if (lastRow >= 1 && recSheet.getLastColumn() >= 1) {
      existingHeaders = recSheet.getRange(1, 1, 1, recSheet.getLastColumn())
        .getValues()[0].map(h => h.toString().trim());
      if (lastRow >= 2) {
        existingData = recSheet.getRange(2, 1, lastRow - 1, recSheet.getLastColumn()).getValues();
      }
    }

    // ── Clear and rebuild sheet ───────────────────────────────
    recSheet.clearContents();
    recSheet.clearFormats();

    // Write headers with colour coding
    _styleHeader(recSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]));

    // Name + Email columns → blue (from _styleHeader, no change)
    // Dynamic columns      → amber (user must fill these)
    if (dynamicCols.length > 0) {
      recSheet.getRange(1, fixedFirst.length + 1, 1, dynamicCols.length)
        .setBackground('#F9A825').setFontColor('#1A1A1A').setFontWeight('bold');
    }
    // Status column → grey
    recSheet.getRange(1, newHeaders.length, 1, 1)
      .setBackground('#90A4AE').setFontColor('#FFFFFF').setFontWeight('bold');

    // Header cell notes
    recSheet.getRange(1, 1).setNote('Required.\nRecipient\'s full name.');
    recSheet.getRange(1, 2).setNote('Required.\nValid email address.');
    dynamicCols.forEach((col, i) => {
      recSheet.getRange(1, fixedFirst.length + 1 + i)
        .setNote(`Fill this value per recipient row.\nUsed as {{${col}}} in the email.`);
    });
    recSheet.getRange(1, newHeaders.length)
      .setNote('Auto-managed by script.\nDo NOT type in this column.');

    // ── Restore matching existing data ────────────────────────
    let restoredCount = 0;
    if (existingData.length > 0) {
      const restoredRows = existingData
        .filter(row => row.some(c => c !== '' && c !== null))
        .map(row =>
          newHeaders.map(h => {
            const oldIdx = existingHeaders.indexOf(h);
            return oldIdx !== -1 ? row[oldIdx] : '';
          })
        );
      if (restoredRows.length > 0) {
        recSheet.getRange(2, 1, restoredRows.length, newHeaders.length).setValues(restoredRows);
        restoredCount = restoredRows.length;
      }
    }

    recSheet.autoResizeColumns(1, newHeaders.length);
    ss.setActiveSheet(recSheet);

    // ── Summary ───────────────────────────────────────────────
    safeAlert('✅ Recipients Sheet Built',
      `Found ${allVars.length} placeholder(s) — columns generated:\n\n` +
      `  🔵 Name        (fixed)\n` +
      `  🔵 Email       (fixed)\n` +
      dynamicCols.map(c => `  🟡 ${c}   ← fill per recipient`).join('\n') +
      `\n  ⚫ Status      (auto)\n\n` +
      (restoredCount > 0
        ? `✅ ${restoredCount} existing row(s) restored.\n\n`
        : '') +
      'Fill in your recipient rows then click\n' +
      '"✅ Validate" or "📧 Send Emails".'
    );

  } catch (e) { _handleError('Parse Template → Build Columns', e); }
}


// ══════════════════════════════════════════════════════════════
//  RECIPIENT DATA READER  (shared by validate / preview / send)
// ══════════════════════════════════════════════════════════════
function _getRecipientData() {
  const sheet     = _getSheet(CONFIG.SHEET_RECIPIENTS);
  const allValues = sheet.getDataRange().getValues();

  if (allValues.length < 1 || !String(allValues[0][0]).trim()) {
    throw new Error(
      'Recipients sheet has no column structure yet.\n' +
      'Run "🔍 Parse Template → Build Columns" first.'
    );
  }

  const headers = allValues[0].map(h => h.toString().trim());

  // Auto-append Status if somehow missing
  let statusColIdx = headers.indexOf(CONFIG.REC_STATUS);
  if (statusColIdx === -1) {
    statusColIdx = headers.length;
    headers.push(CONFIG.REC_STATUS);
    sheet.getRange(1, statusColIdx + 1)
      .setValue(CONFIG.REC_STATUS).setFontWeight('bold');
    SpreadsheetApp.flush();
  }

  const reserved    = [CONFIG.REC_NAME, CONFIG.REC_EMAIL, CONFIG.REC_STATUS];
  const dynamicCols = headers.filter(h => h && !reserved.includes(h));

  if (allValues.length < 2) {
    throw new Error(
      'Recipients sheet has headers but no data rows.\n' +
      'Add recipient rows below the header row.'
    );
  }

  // Filter out blank rows
  const recipients = allValues.slice(1)
    .map((row, i) => {
      const rec = {};
      headers.forEach((h, ci) => { rec[h] = row[ci] ?? ''; });
      rec.__rowIndex = i + 2;
      return rec;
    })
    .filter(rec =>
      String(rec[CONFIG.REC_NAME]  || '').trim() ||
      String(rec[CONFIG.REC_EMAIL] || '').trim()
    );

  return { recipients, headers, dynamicCols, statusColIdx, sheet };
}


// ══════════════════════════════════════════════════════════════
//  3. VALIDATE RECIPIENTS
// ══════════════════════════════════════════════════════════════
function validateRecipients() {
  try {
    const { recipients, dynamicCols } = _getRecipientData();
    const issues = [];

    recipients.forEach(rec => {
      const row   = `Row ${rec.__rowIndex}`;
      const name  = String(rec[CONFIG.REC_NAME]  || '').trim() || '(no name)';
      const email = String(rec[CONFIG.REC_EMAIL] || '').trim();

      if (!String(rec[CONFIG.REC_NAME] || '').trim())
        issues.push(`${row}: Missing Name.`);
      if (!email) {
        issues.push(`${row} [${name}]: Missing Email.`);
      } else if (!CONFIG.EMAIL_REGEX.test(email)) {
        issues.push(`${row} [${name}]: Invalid email → "${email}".`);
      }
      dynamicCols.forEach(col => {
        if (String(rec[col] ?? '').trim() === '')
          issues.push(`${row} [${name}]: "${col}" is empty.`);
      });
    });

    if (issues.length === 0) {
      safeAlert('✅ All Good',
        `All ${recipients.length} recipient(s) passed!\n\n` +
        `Columns validated: ${dynamicCols.join(', ') || '(none)'}\n\n` +
        'Ready to send. Click 📧 Send Emails.'
      );
    } else {
      safeAlert('⚠️ Issues Found',
        `${issues.length} issue(s):\n\n` +
        issues.slice(0, 20).map(i => `• ${i}`).join('\n') +
        (issues.length > 20 ? `\n\n...and ${issues.length - 20} more.` : '')
      );
    }
  } catch (e) { _handleError('Validate Recipients', e); }
}


// ══════════════════════════════════════════════════════════════
//  4. PREVIEW NEXT EMAIL  (Sidebar)
// ══════════════════════════════════════════════════════════════
function previewNextEmail() {
  try {
    const { recipients, dynamicCols } = _getRecipientData();
    const { subject, body }           = _getTemplates();

    const candidate = recipients.find(r =>
      String(r[CONFIG.REC_STATUS] || '').trim() !== CONFIG.STATUS_SENT
    );
    if (!candidate) { safeAlert('📭 All Done', 'All recipients are already marked Sent.'); return; }

    const fieldRows = dynamicCols.map(col =>
      `<tr>
         <td class="lbl">${col}</td>
         <td class="val">${String(candidate[col] || '') !== '' ? candidate[col] : '<em style="color:#ccc">empty</em>'}</td>
       </tr>`
    ).join('');

    SpreadsheetApp.getUi().showSidebar(
      HtmlService
        .createHtmlOutput(_buildPreviewHtml(
          candidate,
          _render(subject, candidate),
          _render(body,    candidate),
          fieldRows
        ))
        .setTitle('👁️ Email Preview')
        .setWidth(480)
    );
  } catch (e) { _handleError('Preview Next Email', e); }
}


// ══════════════════════════════════════════════════════════════
//  5. SEND EMAILS
// ══════════════════════════════════════════════════════════════
function sendEmails() {
  try {
    const { recipients, statusColIdx, sheet } = _getRecipientData();
    const { subject: subjectTpl, body: bodyTpl } = _getTemplates();

    const pending = recipients.filter(r =>
      String(r[CONFIG.REC_STATUS] || '').trim() !== CONFIG.STATUS_SENT
    );

    if (pending.length === 0) {
      safeAlert('ℹ️ Nothing to Send', 'All recipients are already marked Sent.');
      return;
    }

    if (!safeConfirm('📧 Confirm Send',
      `Send personalized emails to ${pending.length} recipient(s)?\n\nAlready-sent rows are skipped automatically.`
    )) return;

    let sent = 0, failed = 0;
    const errors = [];

    pending.forEach((rec, idx) => {
      const email = String(rec[CONFIG.REC_EMAIL] || '').trim();
      const name  = String(rec[CONFIG.REC_NAME]  || '').trim() || `Row ${rec.__rowIndex}`;

      if (!email || !CONFIG.EMAIL_REGEX.test(email)) {
        _writeStatus(sheet, rec.__rowIndex, statusColIdx, CONFIG.STATUS_INVALID);
        errors.push(`• Row ${rec.__rowIndex} [${name}]: Invalid or missing email.`);
        failed++;
        return;
      }

      try {
        // Render plain text first (used for line-break conversion)
        const renderedSubject  = _render(subjectTpl, rec);
        const renderedBodyText = _render(bodyTpl, rec);

        // ✅ FIX: Convert \n newlines → <br> so HTML email shows line breaks
        const renderedBodyHtml = renderedBodyText
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/\n/g, '<br>');

        GmailApp.sendEmail(
          email,
          renderedSubject,
          renderedBodyText,          // plain-text fallback (keeps \n)
          {
            htmlBody :
              `<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.8;color:#222;">` +
              renderedBodyHtml +
              `</div>`,
          }
        );

        _writeStatus(sheet, rec.__rowIndex, statusColIdx, CONFIG.STATUS_SENT);
        sent++;

      } catch (err) {
        _writeStatus(sheet, rec.__rowIndex, statusColIdx, CONFIG.STATUS_FAILED);
        errors.push(`• Row ${rec.__rowIndex} [${name}]: ${err.message}`);
        failed++;
      }

      if (CONFIG.SEND_DELAY_MS > 0 && idx < pending.length - 1)
        Utilities.sleep(CONFIG.SEND_DELAY_MS);
    });

    SpreadsheetApp.flush();
    let summary = `✅ Sent:   ${sent}\n❌ Failed: ${failed}`;
    if (errors.length) summary += '\n\nFailures:\n' + errors.join('\n');
    safeAlert('📬 Emails Sent', summary);

  } catch (e) { _handleError('Send Emails', e); }
}


// ══════════════════════════════════════════════════════════════
//  6. CLEAR SENT STATUS
// ══════════════════════════════════════════════════════════════
function clearSentStatus() {
  try {
    const { recipients, statusColIdx, sheet } = _getRecipientData();
    if (recipients.length === 0) { safeAlert('ℹ️ Empty', 'No recipient data found.'); return; }
    if (!safeConfirm('🔄 Clear Sent Status', 'Reset Status for all recipients?\nAll other data kept.')) return;

    const firstRow = recipients[0].__rowIndex;
    const count    = recipients[recipients.length - 1].__rowIndex - firstRow + 1;
    sheet.getRange(firstRow, statusColIdx + 1, count, 1)
      .clearContent().setBackground(null).setFontWeight('normal');

    safeAlert('✅ Done', 'Status cleared. Ready to send again.');
  } catch (e) { _handleError('Clear Sent Status', e); }
}


// ══════════════════════════════════════════════════════════════
//  7. CLEAR TEMPLATE
// ══════════════════════════════════════════════════════════════
function emptyTemplate() {
  try {
    const sheet = _getSheet(CONFIG.SHEET_TEMPLATE);
    if (!safeConfirm('🧹 Clear Template',
      'Clear Subject and Body?\n• Formatting kept\n• Recipients sheet NOT affected'
    )) return;
    sheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE).clearContent();
    sheet.getRange(CONFIG.BODY_ROW,    CONFIG.COL_VALUE).clearContent();
    safeAlert('✅ Done', 'Template cleared.\nType new content, then re-parse to rebuild columns.');
  } catch (e) { _handleError('Clear Template', e); }
}


// ══════════════════════════════════════════════════════════════
//  PRIVATE UTILITIES
// ══════════════════════════════════════════════════════════════
function _writeStatus(sheet, rowIndex, statusColIdx, statusText) {
  const cell = sheet.getRange(rowIndex, statusColIdx + 1);
  cell.setValue(statusText).setFontWeight('bold');
  cell.setBackground(statusText === CONFIG.STATUS_SENT ? '#B7E1CD' : '#F4C7C3');
}

function _handleError(context, error) {
  Logger.log(`❌ [${context}]: ${error.message}`);
  try {
    SpreadsheetApp.getUi()
      .alert('❌ Error', `[${context}]:\n\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {}
}

function _buildPreviewHtml(recipient, subject, body, fieldRows) {
  return `<!DOCTYPE html>
<html><head><style>
  body  { font-family:Google Sans,Arial,sans-serif; background:#f1f3f4; padding:12px; font-size:13px; margin:0; }
  .card { background:#fff; padding:14px; margin-bottom:10px; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,.12); }
  .lbl  { font-size:10px; font-weight:700; color:#5f6368; text-transform:uppercase; margin-bottom:6px; letter-spacing:.5px; }
  table { width:100%; border-collapse:collapse; }
  td    { padding:5px 4px; border-bottom:1px solid #f0f0f0; }
  td.lbl{ color:#555; width:42%; font-weight:normal; font-size:13px; text-transform:none; letter-spacing:0; }
  td.val{ text-align:right; font-weight:600; }
  .body-preview { white-space:pre-wrap; line-height:1.7; color:#333; }
</style></head><body>
  <div class="card">
    <div class="lbl">Recipient</div>
    <b>${recipient[CONFIG.REC_NAME]}</b><br>
    <span style="color:#5f6368">${recipient[CONFIG.REC_EMAIL]}</span>
  </div>
  <div class="card">
    <div class="lbl">Subject</div>
    <b>${subject}</b>
  </div>
  ${fieldRows ? `<div class="card">
    <div class="lbl">Variable Values</div>
    <table>${fieldRows}</table>
  </div>` : ''}
  <div class="card">
    <div class="lbl">Email Body Preview</div>
    <div class="body-preview">${body}</div>
  </div>
</body></html>`;
}
