/**
 * SVARA QA AUDIT — paste into Code.gs
 * Also add HTML file named exactly: AuditForm (from PASTE_AuditForm.html)
 * Then run: installEverything
 * Sheet: https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w
 */

var SVARA_CONFIG = {
  CLASSIFICATION_SHEET: 'Svara Questions Classification',
  PARAM_COLUMN: 4,
  TYPE_COLUMN: 2,
  CATEGORY_COLUMN: 1,
  CLASSIFICATION_DATA_START_ROW: 2,
  CALLS_DATA_SHEET: 'Calls Data',
  AUDIT_SUBMISSIONS_SHEET: 'Audit Submissions',
  AUDIT_FORM_SHEET: 'Svara Audit Form',
  CALLS_UID_COLUMN: 1,
  PERMANENT_HEADERS: [
    'LAN', 'Call Date', 'Audit date', 'QA Name', 'Call Quality Assessment',
    'Call Duration', 'Penalty Timeline', 'Experiment Tag', 'Recording Link',
    'DPD', 'CX Preferred Language', 'Bot Spoken Language', 'Bot Disposition',
    'Right Disposition according to Conversation', 'Audit Type', 'Bot Activity'
  ],
  QUESTION_OPTIONS: ['Error', 'No error'],
  ROTATING_SECTION_START_ROW: 22,
  LABEL_COL: 2,
  VALUE_COL: 3,
  CONTROL_COL: 4
};

function onOpen() { buildAuditMenu_(); }
function buildAuditMenu() { buildAuditMenu_(); }

function buildAuditMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Audit')
    .addItem('Open Standalone Audit Form', 'openStandaloneAuditForm')
    .addItem('Open Audit Sidebar', 'openAuditSidebar')
    .addSeparator()
    .addItem('Fetch Call Data by UID (in-sheet form)', 'fetchCallDataToForm')
    .addItem('Show Rotating Parameters', 'showRotatingParameters')
    .addItem('Hide Rotating Parameters', 'hideRotatingParameters')
    .addSeparator()
    .addItem('Submit Audit Record (in-sheet form)', 'submitAuditRecord')
    .addItem('Re-install / Refresh Workbook', 'installEverything')
    .addToUi();
}

/** RUN THIS ONCE */
function installEverything() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildClassificationSheet_(ss);
  buildCallsDataSheet_(ss);
  buildParameterClassificationSheet_(ss);
  buildAuditSubmissionsSheet_(ss);
  createAuditFormSheet_();
  buildAuditMenu_();
  SpreadsheetApp.getUi().alert(
    'Svara QA Audit workbook ready.\n\n' +
    'Reload the sheet, then:\n' +
    'Audit → Open Standalone Audit Form\n' +
    'Test UID: SV-10001'
  );
}

function getRotatingQuestions_() {
  return [
    ['Greeting', 'Greeting & Introduction'],
    ['Greeting', 'Purpose of Call Stated'],
    ['Verification', 'Identity Verification'],
    ['Conversation', 'Active Listening'],
    ['Conversation', 'Empathy & Tone'],
    ['Accuracy', 'Accurate Information Provided'],
    ['Handling', 'Objection Handling'],
    ['Closing', 'Call Closing & Summary'],
    ['Compliance', 'Compliance Disclosure'],
    ['Bot Quality', 'Bot Handoff Quality'],
    ['Bot Quality', 'Language Switching Accuracy'],
    ['Bot Quality', 'Disposition Mapping Accuracy'],
    ['Process', 'Hold / Silence Handling'],
    ['Process', 'Repeat Request Handling'],
    ['Outcome', 'Promise / Next Step Captured']
  ];
}

function getPermanentClassRows_() {
  return [
    ['Metadata', 'LAN'], ['Metadata', 'Call Date'], ['Metadata', 'Call Duration'],
    ['Metadata', 'Recording Link'], ['Metadata', 'DPD'],
    ['Language', 'CX Preferred Language'], ['Language', 'Bot Spoken Language'],
    ['Outcome', 'Bot Disposition'],
    ['Outcome', 'Right Disposition according to Conversation'],
    ['Audit', 'Audit Type'], ['Audit', 'Bot Activity']
  ];
}

function upsertSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  else { sheet.clear(); sheet.clearConditionalFormatRules(); }
  return sheet;
}

function buildClassificationSheet_(ss) {
  var sheet = upsertSheet_(ss, SVARA_CONFIG.CLASSIFICATION_SHEET);
  sheet.setTabColor('#0f766e');
  sheet.getRange(1, 1, 1, 4).setValues([['Category', 'Type', 'Section', 'Parameter']])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  var rows = [];
  getPermanentClassRows_().forEach(function (r) {
    rows.push([r[0], 'Permanent', 'Fixed header', r[1]]);
  });
  getRotatingQuestions_().forEach(function (r) {
    rows.push([r[0], 'Rotating', 'Questionnaire', r[1]]);
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.setColumnWidths(1, 3, 160);
  sheet.setColumnWidth(4, 320);
}

function buildCallsDataSheet_(ss) {
  var sheet = upsertSheet_(ss, SVARA_CONFIG.CALLS_DATA_SHEET);
  sheet.setTabColor('#ea4335');
  var headers = ['UID'].concat(SVARA_CONFIG.PERMANENT_HEADERS);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  var samples = [
    ['SV-10001', 'LAN-1001', new Date(2026, 7, 1), '', '', '03:42', 'T+0', 'exp-a',
      'https://example.com/rec/SV-10001', 15, 'Hindi', 'Hindi', 'PTP', 'PTP', 'Regular', 'Active'],
    ['SV-10002', 'LAN-1002', new Date(2026, 7, 1), '', '', '05:11', 'T+1', 'exp-b',
      'https://example.com/rec/SV-10002', 30, 'English', 'English', 'RPC', 'RPC', 'Regular', 'Active'],
    ['SV-10003', 'LAN-1003', new Date(2026, 7, 2), '', '', '02:05', 'T+0', 'exp-a',
      'https://example.com/rec/SV-10003', 7, 'Kannada', 'Kannada', 'Follow-up', 'Follow-up', 'Calibration', 'Silent']
  ];
  sheet.getRange(2, 1, samples.length, headers.length).setValues(samples);
  sheet.setFrozenRows(1);
}

function buildParameterClassificationSheet_(ss) {
  var sheet = upsertSheet_(ss, 'Parameter Classification');
  sheet.setTabColor('#fbbc04');
  sheet.getRange(1, 1, 1, 4)
    .setValues([['Parameter (Col D)', 'Type', 'Category', 'Source Row']])
    .setFontWeight('bold').setBackground('#f59e0b').setFontColor('#ffffff');
  var rows = [];
  var src = 2;
  getPermanentClassRows_().forEach(function (r) {
    rows.push([r[1], 'Permanent', r[0], src++]);
  });
  getRotatingQuestions_().forEach(function (r) {
    rows.push([r[1], 'Rotating', r[0], src++]);
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function buildAuditSubmissionsSheet_(ss) {
  var sheet = upsertSheet_(ss, SVARA_CONFIG.AUDIT_SUBMISSIONS_SHEET);
  sheet.setTabColor('#34a853');
  var headers = ['Submission Timestamp', 'UID', 'Rotating Section Active']
    .concat(SVARA_CONFIG.PERMANENT_HEADERS)
    .concat(getRotatingQuestions_().map(function (r) { return 'Q: ' + r[1]; }));
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#16a34a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function createAuditFormSheet_() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = upsertSheet_(ss, cfg.AUDIT_FORM_SHEET);
  sheet.setTabColor('#4285f4');
  sheet.getRange('B1').setValue('Svara QA Audit Form').setFontWeight('bold').setFontSize(14);
  sheet.getRange('B2').setValue('Enter UID in C4, then Audit → Fetch Call Data (or use Standalone Form).')
    .setFontColor('#666666');
  sheet.getRange(cfg.LABEL_COL, 4).setValue('Unique Identity Number (UID)').setFontWeight('bold');
  sheet.getRange(cfg.VALUE_COL, 4).setValue('SV-10001').setBackground('#fff9c4');
  sheet.getRange(cfg.CONTROL_COL, 4).setValue('← Change UID, then Fetch');

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

  var rotHeaderRow = row + 1;
  sheet.getRange(cfg.LABEL_COL, rotHeaderRow)
    .setValue('ROTATING PARAMETERS (Error / No error)')
    .setFontWeight('bold').setBackground('#fce8e6');
  sheet.getRange(cfg.VALUE_COL, rotHeaderRow).setBackground('#fce8e6');
  sheet.getRange(cfg.CONTROL_COL, rotHeaderRow).setBackground('#fce8e6');

  var questions = getRotatingQuestions_().map(function (r) { return r[1]; });
  row = rotHeaderRow + 1;
  var rotatingStartRow = row;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cfg.QUESTION_OPTIONS, true).setAllowInvalid(false).build();
  questions.forEach(function (q) {
    sheet.getRange(cfg.LABEL_COL, row).setValue(q);
    sheet.getRange(cfg.VALUE_COL, row).setDataValidation(rule);
    row++;
  });

  sheet.getRange('F1').setValue('ROTATING_START_ROW').setFontColor('#ffffff');
  sheet.getRange('G1').setValue(rotatingStartRow);
  sheet.getRange('F2').setValue('ROTATING_END_ROW').setFontColor('#ffffff');
  sheet.getRange('G2').setValue(row - 1);
  sheet.getRange('F3').setValue('ROTATING_VISIBLE').setFontColor('#ffffff');
  sheet.getRange('G3').setValue('TRUE');

  sheet.getRange(cfg.LABEL_COL, row + 1)
    .setValue('Submit via Audit → Submit Audit Record')
    .setFontStyle('italic').setFontColor('#666666');

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(cfg.LABEL_COL, 320);
  sheet.setColumnWidth(cfg.VALUE_COL, 280);
  sheet.setFrozenRows(5);
  ss.setActiveSheet(sheet);
}

function getClassifiedParameters_() {
  var cfg = SVARA_CONFIG;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!sheet) return { permanent: [], rotating: getRotatingQuestions_().map(function (r) { return r[1]; }) };
  var lastRow = sheet.getLastRow();
  if (lastRow < cfg.CLASSIFICATION_DATA_START_ROW) {
    return { permanent: [], rotating: [] };
  }
  var numRows = lastRow - cfg.CLASSIFICATION_DATA_START_ROW + 1;
  var params = sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.PARAM_COLUMN, numRows, 1).getValues();
  var types = sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.TYPE_COLUMN, numRows, 1).getValues();
  var permanent = [];
  var rotating = [];
  var seen = {};
  for (var i = 0; i < numRows; i++) {
    var label = String(params[i][0] || '').trim();
    if (!label || seen[label]) continue;
    seen[label] = true;
    var t = String(types[i][0] || '').toLowerCase();
    if (t.indexOf('perman') !== -1) permanent.push(label);
    else rotating.push(label);
  }
  return { permanent: permanent, rotating: rotating };
}

function normalizeUid_(value) {
  return String(value == null ? '' : value).trim();
}

function fetchCallDataByUid(uid) {
  var cfg = SVARA_CONFIG;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.CALLS_DATA_SHEET);
  if (!sheet) throw new Error('Calls Data sheet missing. Run installEverything.');
  var searchUid = normalizeUid_(uid);
  if (!searchUid) throw new Error('Enter a UID.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var uidValues = sheet.getRange(2, cfg.CALLS_UID_COLUMN, lastRow - 1, 1).getValues();
  for (var r = 0; r < uidValues.length; r++) {
    if (normalizeUid_(uidValues[r][0]) === searchUid) {
      var rowValues = sheet.getRange(r + 2, 1, 1, lastCol).getValues()[0];
      var result = { _uid: searchUid, _row: r + 2 };
      for (var c = 0; c < headers.length; c++) {
        var key = String(headers[c] || '').trim();
        if (key) result[key] = rowValues[c];
      }
      return result;
    }
  }
  return null;
}

function mapCallDataToPermanentFields_(callData) {
  if (!callData) return {};
  var mapped = {};
  var keys = Object.keys(callData);
  SVARA_CONFIG.PERMANENT_HEADERS.forEach(function (header) {
    if (callData[header] !== undefined) { mapped[header] = callData[header]; return; }
    var lower = header.toLowerCase();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].charAt(0) === '_') continue;
      if (keys[i].toLowerCase() === lower) { mapped[header] = callData[keys[i]]; return; }
    }
    mapped[header] = '';
  });
  return mapped;
}

function fetchCallForAuditForm(uid) {
  var callData = fetchCallDataByUid(uid);
  if (!callData) return { found: false, uid: normalizeUid_(uid), fields: {} };
  return { found: true, uid: callData._uid, fields: mapCallDataToPermanentFields_(callData), raw: callData };
}

function fetchCallDataToForm() {
  var cfg = SVARA_CONFIG;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('Run installEverything first.'); return; }
  var uid = normalizeUid_(sheet.getRange(cfg.VALUE_COL, 4).getValue());
  if (!uid) { SpreadsheetApp.getUi().alert('Enter UID in C4.'); return; }
  var result = fetchCallForAuditForm(uid);
  if (!result.found) { SpreadsheetApp.getUi().alert('No call found for UID: ' + uid); return; }
  var permStartRow = 7;
  cfg.PERMANENT_HEADERS.forEach(function (header, idx) {
    var value = result.fields[header];
    sheet.getRange(cfg.VALUE_COL, permStartRow + idx).setValue(value != null ? value : '');
  });
  var auditDateIdx = cfg.PERMANENT_HEADERS.indexOf('Audit date');
  if (auditDateIdx >= 0) {
    var auditCell = sheet.getRange(cfg.VALUE_COL, permStartRow + auditDateIdx);
    if (!auditCell.getValue()) auditCell.setValue(new Date());
  }
  SpreadsheetApp.getUi().alert('Call data loaded for UID: ' + uid);
}

function setRotatingParametersVisible(visible) {
  var cfg = SVARA_CONFIG;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!sheet) return;
  var startRow = Number(sheet.getRange('G1').getValue()) || cfg.ROTATING_SECTION_START_ROW;
  var endRow = Number(sheet.getRange('G2').getValue()) || startRow;
  if (endRow < startRow) return;
  sheet.showRows(startRow, endRow - startRow + 1);
  if (!visible) sheet.hideRows(startRow, endRow - startRow + 1);
  sheet.getRange('G3').setValue(visible ? 'TRUE' : 'FALSE');
}
function showRotatingParameters() { setRotatingParametersVisible(true); }
function hideRotatingParameters() { setRotatingParametersVisible(false); }

function ensureSubmissionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SVARA_CONFIG.AUDIT_SUBMISSIONS_SHEET);
  if (sheet) return sheet;
  return buildAuditSubmissionsSheet_(ss) || ss.getSheetByName(SVARA_CONFIG.AUDIT_SUBMISSIONS_SHEET);
}

function appendSubmissionRow_(sheet, record) {
  var headers = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
        .map(function (h) { return String(h || '').trim(); })
    : [];
  Object.keys(record).forEach(function (key) {
    if (headers.indexOf(key) === -1) headers.push(key);
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.appendRow(headers.map(function (h) { return record[h] !== undefined ? record[h] : ''; }));
}

function submitAuditRecord() {
  var cfg = SVARA_CONFIG;
  var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.AUDIT_FORM_SHEET);
  if (!formSheet) { SpreadsheetApp.getUi().alert('Audit form missing.'); return; }
  var uid = normalizeUid_(formSheet.getRange(cfg.VALUE_COL, 4).getValue());
  if (!uid) { SpreadsheetApp.getUi().alert('Enter UID before submitting.'); return; }
  var submissions = ensureSubmissionsSheet_();
  var classified = getClassifiedParameters_();
  var rotatingQuestions = classified.rotating.length
    ? classified.rotating
    : getRotatingQuestions_().map(function (r) { return r[1]; });
  var record = { 'Submission Timestamp': new Date(), 'UID': uid };
  cfg.PERMANENT_HEADERS.forEach(function (header, idx) {
    record[header] = formSheet.getRange(cfg.VALUE_COL, 7 + idx).getValue();
  });
  var rotStart = Number(formSheet.getRange('G1').getValue()) || cfg.ROTATING_SECTION_START_ROW;
  rotatingQuestions.forEach(function (q, idx) {
    record['Q: ' + q] = formSheet.getRange(cfg.VALUE_COL, rotStart + idx).getValue();
  });
  appendSubmissionRow_(submissions, record);
  SpreadsheetApp.getUi().alert('Audit submitted for UID: ' + uid);
}

function getSidebarBootstrapData() {
  var classified = getClassifiedParameters_();
  return {
    permanentHeaders: SVARA_CONFIG.PERMANENT_HEADERS,
    rotatingQuestions: classified.rotating.length
      ? classified.rotating
      : getRotatingQuestions_().map(function (r) { return r[1]; }),
    questionOptions: SVARA_CONFIG.QUESTION_OPTIONS
  };
}

function submitAuditFromSidebar(payload) {
  if (!payload || !normalizeUid_(payload.uid)) {
    return { success: false, message: 'UID is required.' };
  }
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

function openAuditSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('AuditForm')
    .setTitle('Svara Audit Form')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openStandaloneAuditForm() {
  var html = HtmlService.createHtmlOutputFromFile('AuditForm')
    .setWidth(920)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Svara QA Audit Form');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('AuditForm')
    .setTitle('Svara QA Audit Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
