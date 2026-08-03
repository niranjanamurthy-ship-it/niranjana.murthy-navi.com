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

/** Inline form HTML so only Code.gs is required */
function getInlineFormHtml_() {
  return '' +
'<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1">' +
'<style>' +
':root{--ink:#1c1917;--muted:#78716c;--line:#e7e5e4;--accent:#0f766e;--accent-soft:#ccfbf1;--warn:#9a3412;--warn-soft:#ffedd5;--ok:#166534;--ok-soft:#dcfce7;--err:#b91c1c;--err-soft:#fee2e2;}' +
'*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--ink);background:linear-gradient(180deg,#ecfdf5 0%,#fafaf9 40%);padding:16px}' +
'h1{margin:0 0 4px;font-size:1.5rem}.lede{color:var(--muted);margin:0 0 14px;font-size:.9rem}' +
'.panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px;margin-bottom:12px}' +
'.title{font-size:.72rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--accent);margin:0 0 10px}' +
'.rotating .title{color:var(--warn)}label{display:block;font-size:.8rem;font-weight:600;margin:0 0 4px}' +
'input,select{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:8px;font-size:.92rem}' +
'input[readonly]{background:#fafaf9}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
'@media(max-width:640px){.grid{grid-template-columns:1fr}}' +
'.uid{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:end}' +
'.q{display:grid;grid-template-columns:1.4fr 1fr;gap:8px;align-items:center;padding:8px 0;border-bottom:1px solid var(--line)}' +
'.q:last-child{border-bottom:none}@media(max-width:640px){.q{grid-template-columns:1fr}}' +
'button{border:none;border-radius:8px;padding:9px 14px;font-weight:600;cursor:pointer}' +
'.primary{background:var(--accent);color:#fff}.secondary{background:var(--accent-soft);color:var(--accent)}' +
'.toggle{background:var(--warn-soft);color:var(--warn)}button:disabled{opacity:.55}' +
'.status{display:none;margin-top:8px;padding:8px 10px;border-radius:8px;font-size:.85rem}' +
'.status.show{display:block}.ok{background:var(--ok-soft);color:var(--ok)}.err{background:var(--err-soft);color:var(--err)}.info{background:#f5f5f4;color:var(--muted)}' +
'.hidden{display:none!important}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px}' +
'.head{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:8px}' +
'</style></head><body>' +
'<h1>Svara QA Audit Form</h1>' +
'<p class="lede">Fetch by UID · permanent fields fixed · toggle rotating Error/No error questions</p>' +
'<div class="panel"><div class="title">Call Lookup</div><div class="uid"><div>' +
'<label for="uid">Unique Identity Number (UID)</label>' +
'<input id="uid" type="text" placeholder="e.g. SV-10001"></div>' +
'<button class="primary" id="fetchBtn" onclick="fetchCall()">Fetch Call Data</button></div>' +
'<div id="fetchStatus" class="status"></div></div>' +
'<div class="panel"><div class="title">Permanent Parameters (fixed)</div><div id="permanentFields" class="grid"></div></div>' +
'<div class="panel rotating"><div class="head"><div class="title" style="margin:0">Rotating Parameters</div>' +
'<button class="toggle" id="toggleBtn" onclick="toggleRotating()">Hide Rotating Questions</button></div>' +
'<div id="rotatingBlock"><div id="rotatingQuestions"></div></div></div>' +
'<div class="actions"><button class="primary" id="submitBtn" onclick="submitAudit()">Submit Audit</button>' +
'<button class="secondary" onclick="resetForm()">Clear</button></div>' +
'<div id="submitStatus" class="status"></div>' +
'<script>' +
'var rotatingVisible=true;' +
'document.addEventListener("DOMContentLoaded",function(){' +
'google.script.run.withSuccessHandler(init).withFailureHandler(function(e){show("fetchStatus","err",e.message)}).getSidebarBootstrapData();});' +
'function init(d){renderPerm(d.permanentHeaders||[]);renderQs(d.rotatingQuestions||[],d.questionOptions||["Error","No error"]);}' +
'function renderPerm(headers){var b=document.getElementById("permanentFields");b.innerHTML="";headers.forEach(function(h){' +
'var edit=h==="QA Name"||h==="Audit date"||h==="Call Quality Assessment"||h==="Audit Type";' +
'var d=document.createElement("div");d.innerHTML="<label>"+esc(h)+"</label><input type=\\"text\\" data-permanent=\\""+attr(h)+"\\""+(edit?"":" readonly")+">";b.appendChild(d);});}' +
'function renderQs(qs,opts){var b=document.getElementById("rotatingQuestions");b.innerHTML="";qs.forEach(function(q){' +
'var o=opts.map(function(x){return "<option value=\\""+attr(x)+"\\">"+esc(x)+"</option>";}).join("");' +
'var r=document.createElement("div");r.className="q";r.innerHTML="<div>"+esc(q)+"</div><select data-question=\\""+attr(q)+"\\"><option value=\\"\\">— Select —</option>"+o+"</select>";b.appendChild(r);});}' +
'function fetchCall(){var uid=document.getElementById("uid").value.trim();if(!uid){return show("fetchStatus","err","Enter a UID first.");}' +
'document.getElementById("fetchBtn").disabled=true;show("fetchStatus","info","Loading…");' +
'google.script.run.withSuccessHandler(function(res){document.getElementById("fetchBtn").disabled=false;' +
'if(!res.found){return show("fetchStatus","err","No call found for UID: "+uid);}' +
'Object.keys(res.fields||{}).forEach(function(k){var el=document.querySelector("[data-permanent=\\""+css(k)+"\\"]");if(!el)return;' +
'var v=res.fields[k];el.value=v instanceof Date?fmt(v):(v!=null?String(v):"");});' +
'var a=document.querySelector("[data-permanent=\\"Audit date\\"]");if(a&&!a.value)a.value=fmt(new Date());' +
'show("fetchStatus","ok","Loaded UID "+uid);}).withFailureHandler(function(e){document.getElementById("fetchBtn").disabled=false;show("fetchStatus","err",e.message);}).fetchCallForAuditForm(uid);}' +
'function toggleRotating(){rotatingVisible=!rotatingVisible;var b=document.getElementById("rotatingBlock"),btn=document.getElementById("toggleBtn");' +
'if(rotatingVisible){b.classList.remove("hidden");btn.textContent="Hide Rotating Questions";btn.className="toggle";}' +
'else{b.classList.add("hidden");btn.textContent="Activate Rotating Questions";btn.className="secondary";}}' +
'function submitAudit(){var uid=document.getElementById("uid").value.trim();if(!uid){return show("submitStatus","err","UID is required.");}' +
'var permanent=[],questions=[];document.querySelectorAll("[data-permanent]").forEach(function(el){permanent.push({label:el.getAttribute("data-permanent"),value:el.value});});' +
'if(rotatingVisible){document.querySelectorAll("[data-question]").forEach(function(el){questions.push({label:el.getAttribute("data-question"),value:el.value});});}' +
'document.getElementById("submitBtn").disabled=true;' +
'google.script.run.withSuccessHandler(function(r){document.getElementById("submitBtn").disabled=false;show("submitStatus",r.success?"ok":"err",r.message);})' +
'.withFailureHandler(function(e){document.getElementById("submitBtn").disabled=false;show("submitStatus","err",e.message);})' +
'.submitAuditFromSidebar({uid:uid,permanent:permanent,questions:questions,rotatingActive:rotatingVisible});}' +
'function resetForm(){document.getElementById("uid").value="";document.querySelectorAll("[data-permanent],[data-question]").forEach(function(el){el.value="";});show("fetchStatus","","");show("submitStatus","","");}' +
'function show(id,t,m){var el=document.getElementById(id);el.className="status"+(m?" show "+t:"");el.textContent=m||"";}' +
'function fmt(d){if(!(d instanceof Date))d=new Date(d);return d.toLocaleDateString("en-IN",{year:"numeric",month:"short",day:"2-digit"});}' +
'function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}' +
'function attr(s){return esc(s).replace(/"/g,"&quot;");}function css(s){return String(s).replace(/\\\\/g,"\\\\\\\\").replace(/"/g,"\\\\\\"");} ' +
'</script></body></html>';
}
