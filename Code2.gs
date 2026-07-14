/**
 * ════════════════════════════════════════════════════════════════
 *  EXAM RESULTS EMAIL AUTOMATION
 *  v4.0.0  —  Draft-Style Template  ＋  Student Marks Mailer
 * ════════════════════════════════════════════════════════════════
 */

'use strict';

// ══════════════════════════════════════════════════════════════
//  CONFIG
// ══════════════════════════════════════════════════════════════
const CONFIG = Object.freeze({
  SHEET_TEMPLATE : 'Email_Template',
  SHEET_MARKS    : 'Student_Marks',
  SHEET_WRITEUP  : 'Write-up',

  SUBJECT_ROW : 2,
  BODY_ROW    : 3,
  COL_FIELD   : 1,
  COL_VALUE   : 2,

  STATUS_SENT   : 'Sent',
  STATUS_FAILED : 'FAILED',
  STATUS_INVALID: 'INVALID EMAIL',

  STU_NAME  : 'Student Name',
  STU_EMAIL : 'Email Address',
  STU_STATUS: 'Status',
  STU_TOTAL : 'Total Score',

  EMAIL_REGEX  : /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/,
  SEND_DELAY_MS: 250,
});


// ══════════════════════════════════════════════════════════════
//  MENU  — Flat, no submenus needed (one purpose only)
// ══════════════════════════════════════════════════════════════
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📧 Email Automation')
      .addItem('⚙️  Setup Sheets',          'setupSheets')
      .addSeparator()
      .addItem('📋 Show Placeholders',      'showPlaceholders')   // ← ADD THIS LINE
      .addSeparator()
      .addItem('✅ Validate Student Data',   'validateData')
      .addItem('👁️  Preview Next Email',    'previewNextEmail')
      .addItem('📧 Send Result Emails',     'sendResultEmails')
      .addItem('🔄 Clear Result Status',    'clearResultStatus')
      .addSeparator()
      .addItem('🧹 Clear Template',         'emptyTemplate')
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
//  Reads Subject (B2) and Body (B3) from Email_Template sheet.
// ══════════════════════════════════════════════════════════════
function _getTemplates() {
  const sheet   = _getSheet(CONFIG.SHEET_TEMPLATE);
  const subject = sheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE).getValue().toString().trim();
  const body    = sheet.getRange(CONFIG.BODY_ROW,    CONFIG.COL_VALUE).getValue().toString().trim();

  if (!subject) throw new Error(
    'Subject cell (B2) is empty in Email_Template.\n' +
    'Open that sheet and type your subject in the yellow cell.'
  );
  if (!body) throw new Error(
    'Body cell (B3) is empty in Email_Template.\n' +
    'Open that sheet and type your email body in the yellow cell.'
  );
  return { subject, body };
}


// ══════════════════════════════════════════════════════════════
//  PLACEHOLDER RENDERER
//  Replaces {{Column Header}} with matching value from student.
// ══════════════════════════════════════════════════════════════
function _render(template, data) {
  return template.replace(/\{\{\s*(.+?)\s*\}\}/g, (match, rawKey) => {
    const key   = rawKey.trim();
    const value = data[key];
    if (value === undefined) {
      Logger.log(`⚠️ {{${key}}} has no matching column — left unchanged.`);
      return match;
    }
    if (typeof value === 'number') return Number.isInteger(value) ? String(value) : value.toFixed(2);
    return String(value);
  });
}


// ══════════════════════════════════════════════════════════════
//  1. SETUP SHEETS
//
//  Email_Template  → Styled draft form. Yellow cells = editable.
//                    Column B values LEFT BLANK — user types here.
//                    Cell notes guide non-technical users on hover.
//  Student_Marks   → Headers only. Subject columns added by user.
//  Write-up        → Completely blank. Never touched by script.
// ══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ══ EMAIL_TEMPLATE — Draft Form ═══════════════════════════
  const tplSheet = _getOrCreate(ss, CONFIG.SHEET_TEMPLATE);

  tplSheet.setColumnWidth(CONFIG.COL_FIELD, 90);
  tplSheet.setColumnWidth(CONFIG.COL_VALUE, 520);

  // Row 1: Title banner — tells user exactly what to do
  tplSheet.getRange(1, 1, 1, 2).merge()
    .setValue('📧  Email Draft  —  Click the yellow cells (B2, B3) to type your Subject and Body')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#E8F0FE')
    .setFontColor('#1A73E8')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  tplSheet.setRowHeight(1, 38);

  // Row 2: Subject label (Column A) + editable cell (Column B)
  tplSheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_FIELD)
    .setValue('Subject')
    .setFontWeight('bold')
    .setBackground('#F1F3F4')
    .setFontColor('#444444')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  tplSheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE)
    .setBackground('#FFFDE7')                      // ← Yellow = editable
    .setBorder(true, true, true, true, null, null,
               '#FBBC04', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setNote(
      '✏️ Click here and type your email subject.\n\n' +
      'Use {{Column Name}} to personalise per student:\n' +
      '  {{Student Name}}   → replaced with student\'s name\n\n' +
      'Example:\n' +
      '  Your Mid-Term Results — {{Student Name}}\n' +
      '  Hello {{Student Name}}, your marks are ready!'
    );
  tplSheet.setRowHeight(CONFIG.SUBJECT_ROW, 38);

  // Row 3: Body label (Column A) + editable cell (Column B)
  tplSheet.getRange(CONFIG.BODY_ROW, CONFIG.COL_FIELD)
    .setValue('Body')
    .setFontWeight('bold')
    .setBackground('#F1F3F4')
    .setFontColor('#444444')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('top');

  tplSheet.getRange(CONFIG.BODY_ROW, CONFIG.COL_VALUE)
    .setBackground('#FFFDE7')                      // ← Yellow = editable
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, null, null,
               '#FBBC04', SpreadsheetApp.BorderStyle.SOLID)
    .setNote(
      '✏️ Click here and type your full email body.\n\n' +
      'Use {{Column Header}} to personalise (must match\n' +
      'exact column names in Student_Marks sheet):\n\n' +
      '  {{Student Name}}   → student\'s name\n' +
      '  {{Email Address}}  → student\'s email\n' +
      '  {{Math}}           → Math marks\n' +
      '  {{Science}}        → Science marks\n' +
      '  {{Total Score}}    → auto-calculated total\n\n' +
      'HTML tags are supported:\n' +
      '  <b>bold</b>   <i>italic</i>   <br> = new line\n\n' +
      '── Example body ──────────────────\n' +
      'Dear {{Student Name}},\n\n' +
      'Here are your mid-term results:\n' +
      '  Math    : {{Math}}\n' +
      '  Science : {{Science}}\n' +
      '  Total   : {{Total Score}}\n\n' +
      'Regards,\nClass Representative'
    );
  tplSheet.setRowHeight(CONFIG.BODY_ROW, 240);

  // Row 4: Tip footer
  tplSheet.getRange(4, 1, 1, 2).merge()
    .setValue('💡  Tip: Hover over B2 or B3 for full placeholder guide  |  Use {{exact column header}} from Student_Marks to personalise each email')
    .setFontStyle('italic')
    .setFontSize(10)
    .setFontColor('#888888')
    .setBackground('#FAFAFA')
    .setHorizontalAlignment('center');
  tplSheet.setRowHeight(4, 26);

  Logger.log('Email_Template: draft form ready.');

  // ══ STUDENT_MARKS ══════════════════════════════════════════
  const marksSheet = _getOrCreate(ss, CONFIG.SHEET_MARKS);

  // Only create headers if the sheet is brand new (empty row 1)
  if (!marksSheet.getRange(1, 1).getValue()) {
    _styleHeader(marksSheet.getRange(1, 1, 1, 3)
      .setValues([[CONFIG.STU_NAME, CONFIG.STU_EMAIL, CONFIG.STU_STATUS]]));

    // Header cell notes guide the user without cluttering the sheet
    marksSheet.getRange(1, 1)
      .setNote('Student\'s full name.\nExample: Ravi Kumar');
    marksSheet.getRange(1, 2)
      .setNote(
        'Student\'s email address.\n\n' +
        '💡 To add exam subjects, insert new columns\n' +
        'between this column and the "Status" column.\n' +
        'Example column headers: Math, Science, English\n\n' +
        'The script will auto-detect them and calculate\n' +
        'Total Score automatically.'
      );
    marksSheet.getRange(1, 3)
      .setNote(
        '⚙️ Auto-managed by the script.\n' +
        'Do NOT type in this column.\n\n' +
        'Possible values:\n' +
        '  Sent          → email delivered ✅\n' +
        '  FAILED        → send error ❌\n' +
        '  INVALID EMAIL → bad email format ⚠️'
      );

    marksSheet.autoResizeColumns(1, 3);
    Logger.log('Student_Marks: created with headers.');
  } else {
    Logger.log('Student_Marks: already has data — left untouched.');
  }

  // ══ WRITE-UP ═══════════════════════════════════════════════
  // Script NEVER writes anything here — left for personal writing
  _getOrCreate(ss, CONFIG.SHEET_WRITEUP);
  Logger.log('Write-up: ready (blank).');

  // Open Email_Template so user can start drafting immediately
  ss.setActiveSheet(tplSheet);

  safeAlert('✅ Setup Complete',
    '3 sheets are ready:\n\n' +
    '• Email_Template  — your email draft form\n' +
    '  Click the yellow B2 cell → type Subject\n' +
    '  Click the yellow B3 cell → type Body\n' +
    '  (Hover over yellow cells for placeholder guide 💡)\n\n' +
    '• Student_Marks   — add student rows below the header\n' +
    '  Insert subject columns (Math, Science…) between\n' +
    '  "Email Address" and "Status"\n\n' +
    '• Write-up        — your personal reflection (fill manually)'
  );
}

// ══════════════════════════════════════════════════════════════
//  📋 SHOW PLACEHOLDERS
//  Reads actual column headers from Student_Marks and shows
//  the exact {{placeholder}} names the user must type in the
//  Email_Template body. Non-technical friendly.
// ══════════════════════════════════════════════════════════════
function showPlaceholders() {
  try {
    const sheet     = _getSheet(CONFIG.SHEET_MARKS);
    const allValues = sheet.getDataRange().getValues();

    if (allValues.length < 1) {
      safeAlert('⚠️ Empty Sheet',
        'Student_Marks has no headers yet.\nRun ⚙️ Setup Sheets first.'
      );
      return;
    }

    const headers = allValues[0].map(h => h.toString().trim()).filter(h => h);

    // Reserved non-placeholder columns
    const reserved = [CONFIG.STU_STATUS];

    // Build the list of usable placeholders
    const placeholders = headers
      .filter(h => !reserved.includes(h))
      .map(h => `  {{${h}}}`);

    // Always include Total Score at the end
    placeholders.push(`  {{${CONFIG.STU_TOTAL}}}`);

    const message =
      'Copy these exactly into your Email_Template body (B3):\n\n' +
      placeholders.join('\n') +
      '\n\n──────────────────────────────────\n' +
      '⚠️  Rules:\n' +
      '• Spelling and spaces must match exactly\n' +
      '• {{Total Score}} is always available (auto-calculated)\n' +
      '• Copy-paste from here to avoid typos\n\n' +
      '── Example body ───────────────────\n' +
      `Hello {{${CONFIG.STU_NAME}}},\n\n` +
      'Your marks are ready!\n' +
      headers
        .filter(h => ![CONFIG.STU_NAME, CONFIG.STU_EMAIL, CONFIG.STU_STATUS].includes(h))
        .map(h => `${h.padEnd(12)}: {{${h}}}`)
        .join('\n') +
      `\nTotal Score : {{${CONFIG.STU_TOTAL}}}\n\n` +
      'Regards,\nClass Representative';

    safeAlert('📋 Available Placeholders', message);

  } catch (e) { _handleError('Show Placeholders', e); }
}
// ══════════════════════════════════════════════════════════════
//  2. VALIDATE STUDENT DATA
// ══════════════════════════════════════════════════════════════
function validateData() {
  try {
    const { students, subjectCols } = _getStudentData();
    const issues = [];

    students.forEach(s => {
      const row   = `Row ${s.__rowIndex}`;
      const name  = String(s[CONFIG.STU_NAME]  || '').trim() || '(no name)';
      const email = String(s[CONFIG.STU_EMAIL] || '').trim();

      if (!String(s[CONFIG.STU_NAME] || '').trim())
        issues.push(`${row}: Missing Student Name.`);

      if (!email) {
        issues.push(`${row} [${name}]: Missing Email Address.`);
      } else if (!CONFIG.EMAIL_REGEX.test(email)) {
        issues.push(`${row} [${name}]: Invalid email → "${email}".`);
      }

      subjectCols.forEach(subj => {
        const val = s[subj];
        if (val === '' || val === null || val === undefined) {
          issues.push(`${row} [${name}]: "${subj}" mark is empty.`);
        } else if (isNaN(parseFloat(val))) {
          issues.push(`${row} [${name}]: "${subj}" is not a number → "${val}".`);
        }
      });
    });

    if (issues.length === 0) {
      safeAlert('✅ Validation Passed',
        `All ${students.length} student record(s) are valid and ready!\n\n` +
        `Subjects detected: ${subjectCols.length
          ? subjectCols.join(', ')
          : '(none found — add subject columns to Student_Marks)'}`
      );
    } else {
      safeAlert('⚠️ Issues Found',
        `Found ${issues.length} issue(s):\n\n` +
        issues.slice(0, 20).map(i => `• ${i}`).join('\n') +
        (issues.length > 20 ? `\n\n...and ${issues.length - 20} more.` : '')
      );
    }
  } catch (e) { _handleError('Validate Data', e); }
}


// ══════════════════════════════════════════════════════════════
//  3. PREVIEW NEXT EMAIL  (Sidebar)
// ══════════════════════════════════════════════════════════════
function previewNextEmail() {
  try {
    const { students, subjectCols } = _getStudentData();
    const { subject, body }         = _getTemplates();

    const candidate = students.find(s =>
      String(s[CONFIG.STU_STATUS] || '').trim() !== CONFIG.STATUS_SENT
    );

    if (!candidate) {
      safeAlert('📭 All Done', 'Every student is already marked Sent.');
      return;
    }

    const scoreRows = subjectCols.map(subj =>
      `<tr>
         <td class="lbl">${subj}</td>
         <td class="val">${candidate[subj] !== '' ? candidate[subj] : '—'}</td>
       </tr>`
    ).join('');

    SpreadsheetApp.getUi().showSidebar(
      HtmlService
        .createHtmlOutput(
          _buildPreviewHtml(
            candidate,
            _render(subject, candidate),
            _render(body,    candidate),
            scoreRows
          )
        )
        .setTitle('👁️ Email Preview')
        .setWidth(460)
    );
  } catch (e) { _handleError('Preview Next Email', e); }
}


// ══════════════════════════════════════════════════════════════
//  4. SEND RESULT EMAILS
// ══════════════════════════════════════════════════════════════
function sendResultEmails() {
  try {
    const { students, statusColIdx, sheet } = _getStudentData();
    const { subject: subjectTpl, body: bodyTpl } = _getTemplates();

    const pending = students.filter(s =>
      String(s[CONFIG.STU_STATUS] || '').trim() !== CONFIG.STATUS_SENT
    );

    if (pending.length === 0) {
      safeAlert('ℹ️ Nothing to Send', 'All students are already marked Sent.');
      return;
    }

    if (!safeConfirm('📧 Confirm Send',
      `Send result emails to ${pending.length} student(s)?\n\nAlready-sent students will be skipped automatically.`
    )) return;

    let sent = 0, failed = 0;
    const errors = [];

    pending.forEach((student, idx) => {
      const email = String(student[CONFIG.STU_EMAIL] || '').trim();
      const name  = String(student[CONFIG.STU_NAME]  || '').trim() || `Row ${student.__rowIndex}`;

      if (!email || !CONFIG.EMAIL_REGEX.test(email)) {
        _writeStatus(sheet, student.__rowIndex, statusColIdx, CONFIG.STATUS_INVALID);
        errors.push(`• Row ${student.__rowIndex} [${name}]: Invalid or missing email.`);
        failed++;
        return;
      }

      try {
        GmailApp.sendEmail(
          email,
          _render(subjectTpl, student),
          '',
          { htmlBody: _render(bodyTpl, student), name: 'Class Representative' }
        );
        _writeStatus(sheet, student.__rowIndex, statusColIdx, CONFIG.STATUS_SENT);
        sent++;
      } catch (err) {
        _writeStatus(sheet, student.__rowIndex, statusColIdx, CONFIG.STATUS_FAILED);
        errors.push(`• Row ${student.__rowIndex} [${name}]: ${err.message}`);
        failed++;
      }

      if (CONFIG.SEND_DELAY_MS > 0 && idx < pending.length - 1)
        Utilities.sleep(CONFIG.SEND_DELAY_MS);
    });

    SpreadsheetApp.flush();

    let summary = `✅ Sent:   ${sent}\n❌ Failed: ${failed}`;
    if (errors.length) summary += '\n\nFailures:\n' + errors.join('\n');
    safeAlert('📬 Results Dispatched', summary);

  } catch (e) { _handleError('Send Result Emails', e); }
}


// ══════════════════════════════════════════════════════════════
//  5. CLEAR RESULT STATUS
// ══════════════════════════════════════════════════════════════
function clearResultStatus() {
  try {
    const { students, statusColIdx, sheet } = _getStudentData();
    if (students.length === 0) { safeAlert('ℹ️ Empty', 'No student data found.'); return; }

    if (!safeConfirm('🔄 Clear Result Status',
      'Reset Status column in Student_Marks?\n\nNames, emails, and marks will NOT be affected.'
    )) return;

    // Single range clear — efficient even for large datasets
    const firstRow = students[0].__rowIndex;
    const count    = students[students.length - 1].__rowIndex - firstRow + 1;
    sheet.getRange(firstRow, statusColIdx + 1, count, 1)
      .clearContent().setBackground(null).setFontWeight('normal');

    safeAlert('✅ Done', 'All result statuses cleared. Ready to resend.');
  } catch (e) { _handleError('Clear Result Status', e); }
}


// ══════════════════════════════════════════════════════════════
//  6. CLEAR TEMPLATE
//  Wipes only Subject (B2) and Body (B3) values.
//  All formatting, labels, and layout are preserved.
// ══════════════════════════════════════════════════════════════
function emptyTemplate() {
  try {
    const sheet = _getSheet(CONFIG.SHEET_TEMPLATE);
    if (!safeConfirm('🧹 Clear Template',
      'Clear the Subject and Body content?\n\n' +
      '• Yellow cells will be emptied\n' +
      '• Labels and formatting will be kept\n' +
      '• Student_Marks will NOT be affected'
    )) return;

    sheet.getRange(CONFIG.SUBJECT_ROW, CONFIG.COL_VALUE).clearContent();
    sheet.getRange(CONFIG.BODY_ROW,    CONFIG.COL_VALUE).clearContent();
    safeAlert('✅ Done',
      'Template cleared.\n\nClick the yellow B2 cell to type a new Subject.\nClick the yellow B3 cell to type a new Body.'
    );
  } catch (e) { _handleError('Clear Template', e); }
}


// ══════════════════════════════════════════════════════════════
//  STUDENT DATA READER  (shared by all exam functions)
// ══════════════════════════════════════════════════════════════
function _getStudentData() {
  const sheet     = _getSheet(CONFIG.SHEET_MARKS);
  const allValues = sheet.getDataRange().getValues();

  if (allValues.length < 2) throw new Error(
    'Student_Marks has no data rows yet.\n' +
    'Add student rows below the header row first.'
  );

  const headers = allValues[0].map(h => h.toString().trim());

  // Auto-append Status column if user forgot to add it
  let statusColIdx = headers.indexOf(CONFIG.STU_STATUS);
  if (statusColIdx === -1) {
    statusColIdx = headers.length;
    headers.push(CONFIG.STU_STATUS);
    sheet.getRange(1, statusColIdx + 1)
      .setValue(CONFIG.STU_STATUS).setFontWeight('bold');
    SpreadsheetApp.flush();
  }

  // Detect subject columns dynamically
  const reserved    = [CONFIG.STU_NAME, CONFIG.STU_EMAIL, CONFIG.STU_STATUS];
  const subjectCols = headers.filter(h => h && !reserved.includes(h));

  const students = allValues.slice(1).map((row, i) => {
    const student = {};
    headers.forEach((h, ci) => { student[h] = row[ci] ?? ''; });

    // Auto-calculate total from whatever subject columns exist
    student[CONFIG.STU_TOTAL] = subjectCols.reduce((sum, subj) => {
      const v = parseFloat(student[subj]);
      return sum + (isNaN(v) ? 0 : v);
    }, 0);

    student.__rowIndex = i + 2;
    return student;
  });

  return { students, headers, subjectCols, statusColIdx, sheet };
}


// ══════════════════════════════════════════════════════════════
//  PRIVATE UTILITIES
// ══════════════════════════════════════════════════════════════

/** Colour-coded status writer. Green = Sent, Red = error. */
function _writeStatus(sheet, rowIndex, statusColIdx, statusText) {
  const cell = sheet.getRange(rowIndex, statusColIdx + 1);
  cell.setValue(statusText).setFontWeight('bold');
  cell.setBackground(statusText === CONFIG.STATUS_SENT ? '#B7E1CD' : '#F4C7C3');
}

/** Central error handler — shows alert + logs. */
function _handleError(context, error) {
  Logger.log(`❌ [${context}]: ${error.message}`);
  try {
    SpreadsheetApp.getUi()
      .alert('❌ Error', `[${context}]:\n\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {}
}

/** Builds the preview sidebar HTML. */
function _buildPreviewHtml(student, subject, body, scoreRows) {
  return `<!DOCTYPE html>
<html><head><style>
  body { font-family: Google Sans, Arial, sans-serif; background: #f1f3f4; padding: 12px; font-size: 13px; margin: 0; }
  .card { background: #fff; padding: 14px; margin-bottom: 10px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.12); }
  .label { font-size: 10px; font-weight: 700; color: #5f6368; text-transform: uppercase; margin-bottom: 6px; letter-spacing: .5px; }
  table { width: 100%; border-collapse: collapse; }
  td { padding: 5px 4px; border-bottom: 1px solid #f0f0f0; }
  td:last-child { text-align: right; }
  .total { font-weight: 700; color: #1a73e8; }
  .body-preview { white-space: pre-wrap; line-height: 1.7; color: #333; }
</style></head><body>
  <div class="card">
    <div class="label">Recipient</div>
    <b>${student[CONFIG.STU_NAME]}</b><br>
    <span style="color:#5f6368">${student[CONFIG.STU_EMAIL]}</span>
  </div>
  <div class="card">
    <div class="label">Subject</div>
    ${subject}
  </div>
  <div class="card">
    <div class="label">Score Breakdown</div>
    <table>
      ${scoreRows}
      <tr class="total"><td>Total Score</td><td>${student[CONFIG.STU_TOTAL]}</td></tr>
    </table>
  </div>
  <div class="card">
    <div class="label">Email Body Preview</div>
    <div class="body-preview">${body}</div>
  </div>
</body></html>`;
}
