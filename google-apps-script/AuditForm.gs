/**
 * Builds and manages the in-sheet Svara Audit Form.
 */

/**
 * Creates or refreshes the audit form sheet layout.
 */
function createAuditFormSheet() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.AUDIT_FORM_SHEET);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.clearNotes();
  }

  sheet.setTabColor('#4285f4');

  // Title
  sheet.getRange('B1').setValue('Svara QA Audit Form').setFontWeight('bold').setFontSize(14);
  sheet.getRange('B2').setValue('Use Unique Identity Number to load call data, then complete the audit.')
    .setFontColor('#666666');

  // UID block
  sheet.getRange(cfg.LABEL_COL, 4).setValue('Unique Identity Number (UID)').setFontWeight('bold');
  sheet.getRange(cfg.VALUE_COL, 4).setBackground('#fff9c4');
  sheet.getRange(cfg.CONTROL_COL, 4).setValue('← Enter UID in column C, then use menu: Audit → Fetch Call Data');

  // Permanent section header
  var permStartRow = 6;
  sheet.getRange(cfg.LABEL_COL, permStartRow).setValue('PERMANENT PARAMETERS (fixed)')
    .setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(cfg.VALUE_COL, permStartRow).setBackground('#e8f0fe');
  sheet.getRange(cfg.CONTROL_COL, permStartRow).setBackground('#e8f0fe');

  var row = permStartRow + 1;
  cfg.PERMANENT_HEADERS.forEach(function (header) {
    sheet.getRange(cfg.LABEL_COL, row).setValue(header);
    sheet.getRange(cfg.VALUE_COL, row).setBackground('#f3f3f3');
    row++;
  });

  // Rotating section
  var rotHeaderRow = row + 1;
  sheet.getRange(cfg.LABEL_COL, rotHeaderRow).setValue('ROTATING PARAMETERS (questionnaire — toggle via menu or sidebar)')
    .setFontWeight('bold').setBackground('#fce8e6');
  sheet.getRange(cfg.VALUE_COL, rotHeaderRow).setBackground('#fce8e6');
  sheet.getRange(cfg.CONTROL_COL, rotHeaderRow).setBackground('#fce8e6');

  var classified = getClassifiedParameters_();
  var questions = classified.rotating.length ? classified.rotating : getDefaultRotatingQuestions_();

  row = rotHeaderRow + 1;
  var rotatingStartRow = row;
  questions.forEach(function (question) {
    sheet.getRange(cfg.LABEL_COL, row).setValue(question);
    var cell = sheet.getRange(cfg.VALUE_COL, row);
    applyQuestionValidation_(cell);
    row++;
  });

  // Store metadata for toggling (hidden helper area column F)
  sheet.getRange('F1').setValue('ROTATING_START_ROW').setFontColor('#ffffff');
  sheet.getRange('G1').setValue(rotatingStartRow);
  sheet.getRange('F2').setValue('ROTATING_END_ROW').setFontColor('#ffffff');
  sheet.getRange('G2').setValue(row - 1);
  sheet.getRange('F3').setValue('ROTATING_VISIBLE').setFontColor('#ffffff');
  sheet.getRange('G3').setValue('TRUE');

  // Submit row
  row += 1;
  sheet.getRange(cfg.LABEL_COL, row).setValue('Submit audit via menu: Audit → Submit Audit Record')
    .setFontStyle('italic').setFontColor('#666666');

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(cfg.LABEL_COL, 320);
  sheet.setColumnWidth(cfg.VALUE_COL, 280);
  sheet.setColumnWidth(cfg.CONTROL_COL, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setFrozenRows(5);

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    'Audit form created on "' + cfg.AUDIT_FORM_SHEET + '".\n\n' +
    'Permanent fields: ' + cfg.PERMANENT_HEADERS.length + '\n' +
    'Rotating questions: ' + questions.length + '\n\n' +
    'Use Audit menu → Open Audit Sidebar for UID lookup and rotating toggle.'
  );
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Range} cell
 */
function applyQuestionValidation_(cell) {
  var cfg = SVARA_CONFIG;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cfg.QUESTION_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
}

/**
 * Fallback questions when classification sheet has no rotating rows yet.
 * @returns {string[]}
 */
function getDefaultRotatingQuestions_() {
  return [
    'Greeting & Introduction',
    'Purpose of Call Stated',
    'Identity Verification',
    'Active Listening',
    'Empathy & Tone',
    'Accurate Information Provided',
    'Objection Handling',
    'Call Closing & Summary',
    'Compliance Disclosure',
    'Bot Handoff Quality'
  ];
}

/**
 * Reads UID from form and populates permanent fields.
 */
function fetchCallDataToForm() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Run Setup → Create Audit Form Sheet first.');
    return;
  }

  var uid = normalizeUid_(sheet.getRange(cfg.VALUE_COL, 4).getValue());
  if (!uid) {
    SpreadsheetApp.getUi().alert('Enter a UID in cell C4 first.');
    return;
  }

  var result = fetchCallForAuditForm(uid);
  if (!result.found) {
    SpreadsheetApp.getUi().alert('No call found for UID: ' + uid);
    return;
  }

  var permStartRow = 7;
  cfg.PERMANENT_HEADERS.forEach(function (header, idx) {
    var value = result.fields[header];
    sheet.getRange(cfg.VALUE_COL, permStartRow + idx).setValue(value != null ? value : '');
  });

  // Default audit date to today if empty
  var auditDateIdx = cfg.PERMANENT_HEADERS.indexOf('Audit date');
  if (auditDateIdx >= 0) {
    var auditCell = sheet.getRange(cfg.VALUE_COL, permStartRow + auditDateIdx);
    if (!auditCell.getValue()) {
      auditCell.setValue(new Date());
    }
  }

  SpreadsheetApp.getUi().alert('Call data loaded for UID: ' + uid);
}

/**
 * Show or hide rotating parameter rows on the in-sheet form.
 * @param {boolean} visible
 */
function setRotatingParametersVisible(visible) {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!sheet) return;

  var startRow = Number(sheet.getRange('G1').getValue()) || cfg.ROTATING_SECTION_START_ROW;
  var endRow = Number(sheet.getRange('G2').getValue()) || startRow;
  if (endRow < startRow) return;

  sheet.showRows(startRow, endRow - startRow + 1);
  if (!visible) {
    sheet.hideRows(startRow, endRow - startRow + 1);
  }
  sheet.getRange('G3').setValue(visible ? 'TRUE' : 'FALSE');
}

/** Menu action: show rotating section */
function showRotatingParameters() {
  setRotatingParametersVisible(true);
}

/** Menu action: hide rotating section */
function hideRotatingParameters() {
  setRotatingParametersVisible(false);
}

/**
 * Collects current form values and appends to Audit Submissions.
 */
function submitAuditRecord() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!formSheet) {
    SpreadsheetApp.getUi().alert('Audit form sheet missing. Run Setup → Create Audit Form Sheet.');
    return;
  }

  var uid = normalizeUid_(formSheet.getRange(cfg.VALUE_COL, 4).getValue());
  if (!uid) {
    SpreadsheetApp.getUi().alert('Enter UID before submitting.');
    return;
  }

  var submissions = ensureSubmissionsSheet_();
  var classified = getClassifiedParameters_();
  var rotatingQuestions = classified.rotating.length
    ? classified.rotating
    : getDefaultRotatingQuestions_();

  var permStartRow = 7;
  var record = {
    'Submission Timestamp': new Date(),
    'UID': uid
  };

  cfg.PERMANENT_HEADERS.forEach(function (header, idx) {
    record[header] = formSheet.getRange(cfg.VALUE_COL, permStartRow + idx).getValue();
  });

  var rotStart = Number(formSheet.getRange('G1').getValue()) || cfg.ROTATING_SECTION_START_ROW;
  rotatingQuestions.forEach(function (q, idx) {
    record['Q: ' + q] = formSheet.getRange(cfg.VALUE_COL, rotStart + idx).getValue();
  });

  appendSubmissionRow_(submissions, record);
  SpreadsheetApp.getUi().alert('Audit submitted for UID: ' + uid);
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSubmissionsSheet_() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.AUDIT_SUBMISSIONS_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(cfg.AUDIT_SUBMISSIONS_SHEET);
  sheet.setTabColor('#34a853');
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} record
 */
function appendSubmissionRow_(sheet, record) {
  var headers = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) {
        return String(h || '').trim();
      })
    : [];

  Object.keys(record).forEach(function (key) {
    if (headers.indexOf(key) === -1) headers.push(key);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  var row = headers.map(function (h) {
    return record[h] !== undefined ? record[h] : '';
  });
  sheet.appendRow(row);
}

/**
 * Opens HTML sidebar audit form (recommended workflow).
 */
function openAuditSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('AuditFormSidebar')
    .setTitle('Svara Audit Form')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Sidebar submit handler.
 * @param {Object} payload
 * @returns {{success: boolean, message: string}}
 */
function submitAuditFromSidebar(payload) {
  if (!payload || !normalizeUid_(payload.uid)) {
    return { success: false, message: 'UID is required.' };
  }

  var cfg = SVARA_CONFIG;
  var submissions = ensureSubmissionsSheet_();
  var record = {
    'Submission Timestamp': new Date(),
    'UID': normalizeUid_(payload.uid),
    'Rotating Section Active': payload.rotatingActive ? 'Yes' : 'No'
  };

  (payload.permanent || []).forEach(function (item) {
    if (item && item.label) record[item.label] = item.value;
  });

  (payload.questions || []).forEach(function (q) {
    if (q && q.label) record['Q: ' + q.label] = q.value || '';
  });

  appendSubmissionRow_(submissions, record);
  return { success: true, message: 'Audit saved for UID ' + record['UID'] };
}

/**
 * Returns sidebar bootstrap data.
 */
function getSidebarBootstrapData() {
  var cfg = SVARA_CONFIG;
  var classified = getClassifiedParameters_();
  return {
    permanentHeaders: cfg.PERMANENT_HEADERS,
    rotatingQuestions: classified.rotating.length
      ? classified.rotating
      : getDefaultRotatingQuestions_(),
    questionOptions: cfg.QUESTION_OPTIONS
  };
}
