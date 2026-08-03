/**
 * =============================================================================
 * SVARA QA AUDIT — SINGLE-FILE INSTALLER
 * Sheet: https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w
 *
 * INSTALL (2 minutes):
 * 1. Open the sheet → Extensions → Apps Script
 * 2. Delete any default code in Code.gs
 * 3. Paste THIS ENTIRE FILE into Code.gs → Save (Ctrl/Cmd+S)
 * 4. Select installEverything → Run → Review permissions → Allow
 * 5. Reload the spreadsheet
 * 6. Menu: Audit → Open Standalone Audit Form
 * 7. Test UID: SV-10001
 * =============================================================================
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
    '✅ Svara QA Audit workbook ready.\n\n' +
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
  var html = HtmlService.createHtmlOutput(getInlineFormHtml_())
    .setTitle('Svara Audit Form').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openStandaloneAuditForm() {
  var html = HtmlService.createHtmlOutput(getInlineFormHtml_())
    .setWidth(920).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Svara QA Audit Form');
}

function doGet() {
  return HtmlService.createHtmlOutput(getInlineFormHtml_())
    .setTitle('Svara QA Audit Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Inline form HTML (base64) — avoids paste/escape syntax errors */
function getInlineFormHtml_() {
  var b64 = "PCFET0NUWVBFIGh0bWw+CjxodG1sPgo8aGVhZD4KICA8YmFzZSB0YXJnZXQ9Il90b3AiPgogIDxtZXRhIGNoYXJzZXQ9InV0Zi04" +
  "Ij4KICA8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEiPgog" +
  "IDxzdHlsZT4KICAgIDpyb290IHsKICAgICAgLS1pbms6ICMxYzE5MTc7IC0tbXV0ZWQ6ICM3ODcxNmM7IC0tbGluZTogI2U3ZTVl" +
  "NDsKICAgICAgLS1hY2NlbnQ6ICMwZjc2NmU7IC0tYWNjZW50LXNvZnQ6ICNjY2ZiZjE7CiAgICAgIC0td2FybjogIzlhMzQxMjsg" +
  "LS13YXJuLXNvZnQ6ICNmZmVkZDU7CiAgICAgIC0tb2s6ICMxNjY1MzQ7IC0tb2stc29mdDogI2RjZmNlNzsKICAgICAgLS1lcnI6" +
  "ICNiOTFjMWM7IC0tZXJyLXNvZnQ6ICNmZWUyZTI7CiAgICB9CiAgICAqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQogICAg" +
  "Ym9keSB7CiAgICAgIG1hcmdpbjogMDsgcGFkZGluZzogMTZweDsgY29sb3I6IHZhcigtLWluayk7CiAgICAgIGZvbnQtZmFtaWx5" +
  "OiAiU2Vnb2UgVUkiLCBBcmlhbCwgc2Fucy1zZXJpZjsKICAgICAgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KDE4MGRlZywg" +
  "I2VjZmRmNSAwJSwgI2ZhZmFmOSA0MCUpOwogICAgfQogICAgaDEgeyBtYXJnaW46IDAgMCA0cHg7IGZvbnQtc2l6ZTogMS41cmVt" +
  "OyB9CiAgICAubGVkZSB7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IG1hcmdpbjogMCAwIDE0cHg7IGZvbnQtc2l6ZTogMC45cmVtOyB9" +
  "CiAgICAucGFuZWwgewogICAgICBiYWNrZ3JvdW5kOiAjZmZmOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1saW5lKTsKICAgICAg" +
  "Ym9yZGVyLXJhZGl1czogMTJweDsgcGFkZGluZzogMTRweDsgbWFyZ2luLWJvdHRvbTogMTJweDsKICAgIH0KICAgIC50aXRsZSB7" +
  "CiAgICAgIGZvbnQtc2l6ZTogMC43MnJlbTsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IDAuMDZlbTsKICAgICAg" +
  "dGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgY29sb3I6IHZhcigtLWFjY2VudCk7IG1hcmdpbjogMCAwIDEwcHg7CiAgICB9CiAg" +
  "ICAucm90YXRpbmcgLnRpdGxlIHsgY29sb3I6IHZhcigtLXdhcm4pOyB9CiAgICBsYWJlbCB7IGRpc3BsYXk6IGJsb2NrOyBmb250" +
  "LXNpemU6IDAuOHJlbTsgZm9udC13ZWlnaHQ6IDYwMDsgbWFyZ2luOiAwIDAgNHB4OyB9CiAgICBpbnB1dCwgc2VsZWN0IHsKICAg" +
  "ICAgd2lkdGg6IDEwMCU7IHBhZGRpbmc6IDhweCAxMHB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1saW5lKTsKICAgICAgYm9y" +
  "ZGVyLXJhZGl1czogOHB4OyBmb250LXNpemU6IDAuOTJyZW07CiAgICB9CiAgICBpbnB1dFtyZWFkb25seV0geyBiYWNrZ3JvdW5k" +
  "OiAjZmFmYWY5OyB9CiAgICAuZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgZ2Fw" +
  "OiAxMHB4OyB9CiAgICAudWlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgYXV0bzsgZ2FwOiA4" +
  "cHg7IGFsaWduLWl0ZW1zOiBlbmQ7IH0KICAgIC5xIHsKICAgICAgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5z" +
  "OiAxLjRmciAxZnI7IGdhcDogOHB4OwogICAgICBhbGlnbi1pdGVtczogY2VudGVyOyBwYWRkaW5nOiA4cHggMDsgYm9yZGVyLWJv" +
  "dHRvbTogMXB4IHNvbGlkIHZhcigtLWxpbmUpOwogICAgfQogICAgLnE6bGFzdC1jaGlsZCB7IGJvcmRlci1ib3R0b206IG5vbmU7" +
  "IH0KICAgIGJ1dHRvbiB7CiAgICAgIGJvcmRlcjogbm9uZTsgYm9yZGVyLXJhZGl1czogOHB4OyBwYWRkaW5nOiA5cHggMTRweDsK" +
  "ICAgICAgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogICAgfQogICAgLnByaW1hcnkgeyBiYWNrZ3JvdW5kOiB2" +
  "YXIoLS1hY2NlbnQpOyBjb2xvcjogI2ZmZjsgfQogICAgLnNlY29uZGFyeSB7IGJhY2tncm91bmQ6IHZhcigtLWFjY2VudC1zb2Z0" +
  "KTsgY29sb3I6IHZhcigtLWFjY2VudCk7IH0KICAgIC50b2dnbGUgeyBiYWNrZ3JvdW5kOiB2YXIoLS13YXJuLXNvZnQpOyBjb2xv" +
  "cjogdmFyKC0td2Fybik7IH0KICAgIGJ1dHRvbjpkaXNhYmxlZCB7IG9wYWNpdHk6IDAuNTU7IH0KICAgIC5zdGF0dXMgewogICAg" +
  "ICBkaXNwbGF5OiBub25lOyBtYXJnaW4tdG9wOiA4cHg7IHBhZGRpbmc6IDhweCAxMHB4OwogICAgICBib3JkZXItcmFkaXVzOiA4" +
  "cHg7IGZvbnQtc2l6ZTogMC44NXJlbTsKICAgIH0KICAgIC5zdGF0dXMuc2hvdyB7IGRpc3BsYXk6IGJsb2NrOyB9CiAgICAub2sg" +
  "eyBiYWNrZ3JvdW5kOiB2YXIoLS1vay1zb2Z0KTsgY29sb3I6IHZhcigtLW9rKTsgfQogICAgLmVyciB7IGJhY2tncm91bmQ6IHZh" +
  "cigtLWVyci1zb2Z0KTsgY29sb3I6IHZhcigtLWVycik7IH0KICAgIC5pbmZvIHsgYmFja2dyb3VuZDogI2Y1ZjVmNDsgY29sb3I6" +
  "IHZhcigtLW11dGVkKTsgfQogICAgLmhpZGRlbiB7IGRpc3BsYXk6IG5vbmUgIWltcG9ydGFudDsgfQogICAgLmFjdGlvbnMgeyBk" +
  "aXNwbGF5OiBmbGV4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyBtYXJnaW4tdG9wOiA0cHg7IH0KICAgIC5oZWFkIHsKICAg" +
  "ICAgZGlzcGxheTogZmxleDsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczogY2VudGVyOwogICAg" +
  "ICBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyBtYXJnaW4tYm90dG9tOiA4cHg7CiAgICB9CiAgICBAbWVkaWEgKG1heC13aWR0" +
  "aDogNjQwcHgpIHsKICAgICAgLmdyaWQsIC5xLCAudWlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0KICAgIH0KICA8" +
  "L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PgogIDxoMT5TdmFyYSBRQSBBdWRpdCBGb3JtPC9oMT4KICA8cCBjbGFzcz0ibGVkZSI+RmV0" +
  "Y2ggYnkgVUlELiBQZXJtYW5lbnQgZmllbGRzIHN0YXkgZml4ZWQuIFRvZ2dsZSByb3RhdGluZyBFcnJvciAvIE5vIGVycm9yIHF1" +
  "ZXN0aW9ucyB3aGVuIG5lZWRlZC48L3A+CgogIDxkaXYgY2xhc3M9InBhbmVsIj4KICAgIDxkaXYgY2xhc3M9InRpdGxlIj5DYWxs" +
  "IExvb2t1cDwvZGl2PgogICAgPGRpdiBjbGFzcz0idWlkIj4KICAgICAgPGRpdj4KICAgICAgICA8bGFiZWwgZm9yPSJ1aWQiPlVu" +
  "aXF1ZSBJZGVudGl0eSBOdW1iZXIgKFVJRCk8L2xhYmVsPgogICAgICAgIDxpbnB1dCBpZD0idWlkIiB0eXBlPSJ0ZXh0IiBwbGFj" +
  "ZWhvbGRlcj0iZS5nLiBTVi0xMDAwMSI+CiAgICAgIDwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJwcmltYXJ5IiBpZD0iZmV0" +
  "Y2hCdG4iIHR5cGU9ImJ1dHRvbiI+RmV0Y2ggQ2FsbCBEYXRhPC9idXR0b24+CiAgICA8L2Rpdj4KICAgIDxkaXYgaWQ9ImZldGNo" +
  "U3RhdHVzIiBjbGFzcz0ic3RhdHVzIj48L2Rpdj4KICA8L2Rpdj4KCiAgPGRpdiBjbGFzcz0icGFuZWwiPgogICAgPGRpdiBjbGFz" +
  "cz0idGl0bGUiPlBlcm1hbmVudCBQYXJhbWV0ZXJzIChmaXhlZCk8L2Rpdj4KICAgIDxkaXYgaWQ9InBlcm1hbmVudEZpZWxkcyIg" +
  "Y2xhc3M9ImdyaWQiPjwvZGl2PgogIDwvZGl2PgoKICA8ZGl2IGNsYXNzPSJwYW5lbCByb3RhdGluZyI+CiAgICA8ZGl2IGNsYXNz" +
  "PSJoZWFkIj4KICAgICAgPGRpdiBjbGFzcz0idGl0bGUiIHN0eWxlPSJtYXJnaW46MCI+Um90YXRpbmcgUGFyYW1ldGVyczwvZGl2" +
  "PgogICAgICA8YnV0dG9uIGNsYXNzPSJ0b2dnbGUiIGlkPSJ0b2dnbGVCdG4iIHR5cGU9ImJ1dHRvbiI+SGlkZSBSb3RhdGluZyBR" +
  "dWVzdGlvbnM8L2J1dHRvbj4KICAgIDwvZGl2PgogICAgPGRpdiBpZD0icm90YXRpbmdCbG9jayI+CiAgICAgIDxkaXYgaWQ9InJv" +
  "dGF0aW5nUXVlc3Rpb25zIj48L2Rpdj4KICAgIDwvZGl2PgogIDwvZGl2PgoKICA8ZGl2IGNsYXNzPSJhY3Rpb25zIj4KICAgIDxi" +
  "dXR0b24gY2xhc3M9InByaW1hcnkiIGlkPSJzdWJtaXRCdG4iIHR5cGU9ImJ1dHRvbiI+U3VibWl0IEF1ZGl0PC9idXR0b24+CiAg" +
  "ICA8YnV0dG9uIGNsYXNzPSJzZWNvbmRhcnkiIGlkPSJjbGVhckJ0biIgdHlwZT0iYnV0dG9uIj5DbGVhcjwvYnV0dG9uPgogIDwv" +
  "ZGl2PgogIDxkaXYgaWQ9InN1Ym1pdFN0YXR1cyIgY2xhc3M9InN0YXR1cyI+PC9kaXY+CgogIDxzY3JpcHQ+CiAgICB2YXIgcm90" +
  "YXRpbmdWaXNpYmxlID0gdHJ1ZTsKCiAgICBkb2N1bWVudC5hZGRFdmVudExpc3RlbmVyKCdET01Db250ZW50TG9hZGVkJywgZnVu" +
  "Y3Rpb24gKCkgewogICAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmV0Y2hCdG4nKS5hZGRFdmVudExpc3RlbmVyKCdjbGlj" +
  "aycsIGZldGNoQ2FsbCk7CiAgICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2dnbGVCdG4nKS5hZGRFdmVudExpc3RlbmVy" +
  "KCdjbGljaycsIHRvZ2dsZVJvdGF0aW5nKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3N1Ym1pdEJ0bicpLmFkZEV2" +
  "ZW50TGlzdGVuZXIoJ2NsaWNrJywgc3VibWl0QXVkaXQpOwogICAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xlYXJCdG4n" +
  "KS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIHJlc2V0Rm9ybSk7CiAgICAgIGdvb2dsZS5zY3JpcHQucnVuCiAgICAgICAgLndp" +
  "dGhTdWNjZXNzSGFuZGxlcihpbml0KQogICAgICAgIC53aXRoRmFpbHVyZUhhbmRsZXIoZnVuY3Rpb24gKGUpIHsgc2hvdygnZmV0" +
  "Y2hTdGF0dXMnLCAnZXJyJywgZS5tZXNzYWdlKTsgfSkKICAgICAgICAuZ2V0U2lkZWJhckJvb3RzdHJhcERhdGEoKTsKICAgIH0p" +
  "OwoKICAgIGZ1bmN0aW9uIGluaXQoZGF0YSkgewogICAgICByZW5kZXJQZXJtKGRhdGEucGVybWFuZW50SGVhZGVycyB8fCBbXSk7" +
  "CiAgICAgIHJlbmRlclFzKGRhdGEucm90YXRpbmdRdWVzdGlvbnMgfHwgW10sIGRhdGEucXVlc3Rpb25PcHRpb25zIHx8IFsnRXJy" +
  "b3InLCAnTm8gZXJyb3InXSk7CiAgICB9CgogICAgZnVuY3Rpb24gcmVuZGVyUGVybShoZWFkZXJzKSB7CiAgICAgIHZhciBib3gg" +
  "PSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGVybWFuZW50RmllbGRzJyk7CiAgICAgIGJveC5pbm5lckhUTUwgPSAnJzsKICAg" +
  "ICAgaGVhZGVycy5mb3JFYWNoKGZ1bmN0aW9uIChoZWFkZXIpIHsKICAgICAgICB2YXIgZWRpdGFibGUgPSBoZWFkZXIgPT09ICdR" +
  "QSBOYW1lJyB8fCBoZWFkZXIgPT09ICdBdWRpdCBkYXRlJyB8fAogICAgICAgICAgaGVhZGVyID09PSAnQ2FsbCBRdWFsaXR5IEFz" +
  "c2Vzc21lbnQnIHx8IGhlYWRlciA9PT0gJ0F1ZGl0IFR5cGUnOwogICAgICAgIHZhciB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxl" +
  "bWVudCgnZGl2Jyk7CiAgICAgICAgdmFyIGxhYmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGFiZWwnKTsKICAgICAgICBs" +
  "YWJlbC50ZXh0Q29udGVudCA9IGhlYWRlcjsKICAgICAgICB2YXIgaW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1" +
  "dCcpOwogICAgICAgIGlucHV0LnR5cGUgPSAndGV4dCc7CiAgICAgICAgaW5wdXQuc2V0QXR0cmlidXRlKCdkYXRhLXBlcm1hbmVu" +
  "dCcsIGhlYWRlcik7CiAgICAgICAgaWYgKCFlZGl0YWJsZSkgaW5wdXQucmVhZE9ubHkgPSB0cnVlOwogICAgICAgIHdyYXAuYXBw" +
  "ZW5kQ2hpbGQobGFiZWwpOwogICAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoaW5wdXQpOwogICAgICAgIGJveC5hcHBlbmRDaGlsZCh3" +
  "cmFwKTsKICAgICAgfSk7CiAgICB9CgogICAgZnVuY3Rpb24gcmVuZGVyUXMocXVlc3Rpb25zLCBvcHRpb25zKSB7CiAgICAgIHZh" +
  "ciBib3ggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncm90YXRpbmdRdWVzdGlvbnMnKTsKICAgICAgYm94LmlubmVySFRNTCA9" +
  "ICcnOwogICAgICBxdWVzdGlvbnMuZm9yRWFjaChmdW5jdGlvbiAocSkgewogICAgICAgIHZhciByb3cgPSBkb2N1bWVudC5jcmVh" +
  "dGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ3EnOwogICAgICAgIHZhciBsYWJlbCA9IGRvY3VtZW50" +
  "LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGxhYmVsLnRleHRDb250ZW50ID0gcTsKICAgICAgICB2YXIgc2VsZWN0ID0g" +
  "ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICAgICAgc2VsZWN0LnNldEF0dHJpYnV0ZSgnZGF0YS1xdWVzdGlv" +
  "bicsIHEpOwogICAgICAgIHZhciBibGFuayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICAgIGJsYW5r" +
  "LnZhbHVlID0gJyc7CiAgICAgICAgYmxhbmsudGV4dENvbnRlbnQgPSAnLSBTZWxlY3QgLSc7CiAgICAgICAgc2VsZWN0LmFwcGVu" +
  "ZENoaWxkKGJsYW5rKTsKICAgICAgICBvcHRpb25zLmZvckVhY2goZnVuY3Rpb24gKG9wdCkgewogICAgICAgICAgdmFyIG8gPSBk" +
  "b2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgICAgIG8udmFsdWUgPSBvcHQ7CiAgICAgICAgICBvLnRleHRD" +
  "b250ZW50ID0gb3B0OwogICAgICAgICAgc2VsZWN0LmFwcGVuZENoaWxkKG8pOwogICAgICAgIH0pOwogICAgICAgIHJvdy5hcHBl" +
  "bmRDaGlsZChsYWJlbCk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKHNlbGVjdCk7CiAgICAgICAgYm94LmFwcGVuZENoaWxkKHJv" +
  "dyk7CiAgICAgIH0pOwogICAgfQoKICAgIGZ1bmN0aW9uIGZldGNoQ2FsbCgpIHsKICAgICAgdmFyIHVpZCA9IGRvY3VtZW50Lmdl" +
  "dEVsZW1lbnRCeUlkKCd1aWQnKS52YWx1ZS50cmltKCk7CiAgICAgIGlmICghdWlkKSB7CiAgICAgICAgc2hvdygnZmV0Y2hTdGF0" +
  "dXMnLCAnZXJyJywgJ0VudGVyIGEgVUlEIGZpcnN0LicpOwogICAgICAgIHJldHVybjsKICAgICAgfQogICAgICBkb2N1bWVudC5n" +
  "ZXRFbGVtZW50QnlJZCgnZmV0Y2hCdG4nKS5kaXNhYmxlZCA9IHRydWU7CiAgICAgIHNob3coJ2ZldGNoU3RhdHVzJywgJ2luZm8n" +
  "LCAnTG9hZGluZy4uLicpOwogICAgICBnb29nbGUuc2NyaXB0LnJ1bgogICAgICAgIC53aXRoU3VjY2Vzc0hhbmRsZXIoZnVuY3Rp" +
  "b24gKHJlc3VsdCkgewogICAgICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZldGNoQnRuJykuZGlzYWJsZWQgPSBmYWxz" +
  "ZTsKICAgICAgICAgIGlmICghcmVzdWx0LmZvdW5kKSB7CiAgICAgICAgICAgIHNob3coJ2ZldGNoU3RhdHVzJywgJ2VycicsICdO" +
  "byBjYWxsIGZvdW5kIGZvciBVSUQ6ICcgKyB1aWQpOwogICAgICAgICAgICByZXR1cm47CiAgICAgICAgICB9CiAgICAgICAgICBP" +
  "YmplY3Qua2V5cyhyZXN1bHQuZmllbGRzIHx8IHt9KS5mb3JFYWNoKGZ1bmN0aW9uIChrZXkpIHsKICAgICAgICAgICAgdmFyIGlu" +
  "cHV0ID0gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignW2RhdGEtcGVybWFuZW50PSInICsgY3NzRXNjKGtleSkgKyAnIl0nKTsKICAg" +
  "ICAgICAgICAgaWYgKCFpbnB1dCkgcmV0dXJuOwogICAgICAgICAgICB2YXIgdmFsID0gcmVzdWx0LmZpZWxkc1trZXldOwogICAg" +
  "ICAgICAgICBpbnB1dC52YWx1ZSA9IHZhbCBpbnN0YW5jZW9mIERhdGUgPyBmbXQodmFsKSA6ICh2YWwgIT0gbnVsbCA/IFN0cmlu" +
  "Zyh2YWwpIDogJycpOwogICAgICAgICAgfSk7CiAgICAgICAgICB2YXIgYXVkaXQgPSBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCdb" +
  "ZGF0YS1wZXJtYW5lbnQ9IkF1ZGl0IGRhdGUiXScpOwogICAgICAgICAgaWYgKGF1ZGl0ICYmICFhdWRpdC52YWx1ZSkgYXVkaXQu" +
  "dmFsdWUgPSBmbXQobmV3IERhdGUoKSk7CiAgICAgICAgICBzaG93KCdmZXRjaFN0YXR1cycsICdvaycsICdMb2FkZWQgVUlEICcg" +
  "KyB1aWQpOwogICAgICAgIH0pCiAgICAgICAgLndpdGhGYWlsdXJlSGFuZGxlcihmdW5jdGlvbiAoZSkgewogICAgICAgICAgZG9j" +
  "dW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZldGNoQnRuJykuZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgICAgIHNob3coJ2ZldGNoU3Rh" +
  "dHVzJywgJ2VycicsIGUubWVzc2FnZSk7CiAgICAgICAgfSkKICAgICAgICAuZmV0Y2hDYWxsRm9yQXVkaXRGb3JtKHVpZCk7CiAg" +
  "ICB9CgogICAgZnVuY3Rpb24gdG9nZ2xlUm90YXRpbmcoKSB7CiAgICAgIHJvdGF0aW5nVmlzaWJsZSA9ICFyb3RhdGluZ1Zpc2li" +
  "bGU7CiAgICAgIHZhciBibG9jayA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyb3RhdGluZ0Jsb2NrJyk7CiAgICAgIHZhciBi" +
  "dG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9nZ2xlQnRuJyk7CiAgICAgIGlmIChyb3RhdGluZ1Zpc2libGUpIHsKICAg" +
  "ICAgICBibG9jay5jbGFzc0xpc3QucmVtb3ZlKCdoaWRkZW4nKTsKICAgICAgICBidG4udGV4dENvbnRlbnQgPSAnSGlkZSBSb3Rh" +
  "dGluZyBRdWVzdGlvbnMnOwogICAgICAgIGJ0bi5jbGFzc05hbWUgPSAndG9nZ2xlJzsKICAgICAgfSBlbHNlIHsKICAgICAgICBi" +
  "bG9jay5jbGFzc0xpc3QuYWRkKCdoaWRkZW4nKTsKICAgICAgICBidG4udGV4dENvbnRlbnQgPSAnQWN0aXZhdGUgUm90YXRpbmcg" +
  "UXVlc3Rpb25zJzsKICAgICAgICBidG4uY2xhc3NOYW1lID0gJ3NlY29uZGFyeSc7CiAgICAgIH0KICAgIH0KCiAgICBmdW5jdGlv" +
  "biBzdWJtaXRBdWRpdCgpIHsKICAgICAgdmFyIHVpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd1aWQnKS52YWx1ZS50cmlt" +
  "KCk7CiAgICAgIGlmICghdWlkKSB7CiAgICAgICAgc2hvdygnc3VibWl0U3RhdHVzJywgJ2VycicsICdVSUQgaXMgcmVxdWlyZWQu" +
  "Jyk7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHZhciBwZXJtYW5lbnQgPSBbXTsKICAgICAgZG9jdW1lbnQucXVlcnlT" +
  "ZWxlY3RvckFsbCgnW2RhdGEtcGVybWFuZW50XScpLmZvckVhY2goZnVuY3Rpb24gKGVsKSB7CiAgICAgICAgcGVybWFuZW50LnB1" +
  "c2goeyBsYWJlbDogZWwuZ2V0QXR0cmlidXRlKCdkYXRhLXBlcm1hbmVudCcpLCB2YWx1ZTogZWwudmFsdWUgfSk7CiAgICAgIH0p" +
  "OwogICAgICB2YXIgcXVlc3Rpb25zID0gW107CiAgICAgIGlmIChyb3RhdGluZ1Zpc2libGUpIHsKICAgICAgICBkb2N1bWVudC5x" +
  "dWVyeVNlbGVjdG9yQWxsKCdbZGF0YS1xdWVzdGlvbl0nKS5mb3JFYWNoKGZ1bmN0aW9uIChlbCkgewogICAgICAgICAgcXVlc3Rp" +
  "b25zLnB1c2goeyBsYWJlbDogZWwuZ2V0QXR0cmlidXRlKCdkYXRhLXF1ZXN0aW9uJyksIHZhbHVlOiBlbC52YWx1ZSB9KTsKICAg" +
  "ICAgICB9KTsKICAgICAgfQogICAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc3VibWl0QnRuJykuZGlzYWJsZWQgPSB0cnVl" +
  "OwogICAgICBnb29nbGUuc2NyaXB0LnJ1bgogICAgICAgIC53aXRoU3VjY2Vzc0hhbmRsZXIoZnVuY3Rpb24gKHJlcykgewogICAg" +
  "ICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3N1Ym1pdEJ0bicpLmRpc2FibGVkID0gZmFsc2U7CiAgICAgICAgICBzaG93" +
  "KCdzdWJtaXRTdGF0dXMnLCByZXMuc3VjY2VzcyA/ICdvaycgOiAnZXJyJywgcmVzLm1lc3NhZ2UpOwogICAgICAgIH0pCiAgICAg" +
  "ICAgLndpdGhGYWlsdXJlSGFuZGxlcihmdW5jdGlvbiAoZSkgewogICAgICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3N1" +
  "Ym1pdEJ0bicpLmRpc2FibGVkID0gZmFsc2U7CiAgICAgICAgICBzaG93KCdzdWJtaXRTdGF0dXMnLCAnZXJyJywgZS5tZXNzYWdl" +
  "KTsKICAgICAgICB9KQogICAgICAgIC5zdWJtaXRBdWRpdEZyb21TaWRlYmFyKHsKICAgICAgICAgIHVpZDogdWlkLAogICAgICAg" +
  "ICAgcGVybWFuZW50OiBwZXJtYW5lbnQsCiAgICAgICAgICBxdWVzdGlvbnM6IHF1ZXN0aW9ucywKICAgICAgICAgIHJvdGF0aW5n" +
  "QWN0aXZlOiByb3RhdGluZ1Zpc2libGUKICAgICAgICB9KTsKICAgIH0KCiAgICBmdW5jdGlvbiByZXNldEZvcm0oKSB7CiAgICAg" +
  "IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd1aWQnKS52YWx1ZSA9ICcnOwogICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxs" +
  "KCdbZGF0YS1wZXJtYW5lbnRdLCBbZGF0YS1xdWVzdGlvbl0nKS5mb3JFYWNoKGZ1bmN0aW9uIChlbCkgewogICAgICAgIGVsLnZh" +
  "bHVlID0gJyc7CiAgICAgIH0pOwogICAgICBzaG93KCdmZXRjaFN0YXR1cycsICcnLCAnJyk7CiAgICAgIHNob3coJ3N1Ym1pdFN0" +
  "YXR1cycsICcnLCAnJyk7CiAgICB9CgogICAgZnVuY3Rpb24gc2hvdyhpZCwgdHlwZSwgbXNnKSB7CiAgICAgIHZhciBlbCA9IGRv" +
  "Y3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICAgICAgZWwuY2xhc3NOYW1lID0gJ3N0YXR1cycgKyAobXNnID8gJyBzaG93ICcg" +
  "KyB0eXBlIDogJycpOwogICAgICBlbC50ZXh0Q29udGVudCA9IG1zZyB8fCAnJzsKICAgIH0KCiAgICBmdW5jdGlvbiBmbXQoZCkg" +
  "ewogICAgICBpZiAoIShkIGluc3RhbmNlb2YgRGF0ZSkpIGQgPSBuZXcgRGF0ZShkKTsKICAgICAgcmV0dXJuIGQudG9Mb2NhbGVE" +
  "YXRlU3RyaW5nKCdlbi1JTicsIHsgeWVhcjogJ251bWVyaWMnLCBtb250aDogJ3Nob3J0JywgZGF5OiAnMi1kaWdpdCcgfSk7CiAg" +
  "ICB9CgogICAgZnVuY3Rpb24gY3NzRXNjKHMpIHsKICAgICAgcmV0dXJuIFN0cmluZyhzKS5yZXBsYWNlKC9cXC9nLCAnXFxcXCcp" +
  "LnJlcGxhY2UoLyIvZywgJ1xcIicpOwogICAgfQogIDwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K";
  return Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();
}
