/**
 * ONE-CLICK INSTALLER for your Svara Google Sheet
 * Spreadsheet: https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ
 *
 * HOW TO RUN:
 * 1. Open that sheet → Extensions → Apps Script
 * 2. Paste this file as Installer.gs (and the other google-apps-script files)
 * 3. Select installSvaraAuditWorkbook → Run → Authorize
 * 4. Reload the sheet → Audit menu appears
 */

function installSvaraAuditWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildClassificationSheet_(ss);
  buildCallsDataSheet_(ss);
  buildParameterClassificationSheet_(ss);
  buildAuditSubmissionsSheet_(ss);
  // Uses AuditForm.gs + Classification.gs + Config.gs
  createAuditFormSheet();
  buildAuditMenu_();
  SpreadsheetApp.getUi().alert(
    'Svara QA Audit workbook installed.\n\n' +
    'Next: Audit → Open Standalone Audit Form\n' +
    'Test UID: SV-10001'
  );
}

function getPermanentHeaders_() {
  return SVARA_CONFIG.PERMANENT_HEADERS;
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

function getPermanentCategory_(header) {
  if (['LAN', 'Call Date', 'Call Duration', 'Recording Link', 'DPD', 'Experiment Tag'].indexOf(header) >= 0) {
    return 'Metadata';
  }
  if (header.indexOf('Language') >= 0) return 'Language';
  if (header.indexOf('Disposition') >= 0) return 'Outcome';
  if (['Audit date', 'QA Name', 'Call Quality Assessment', 'Penalty Timeline', 'Audit Type', 'Bot Activity'].indexOf(header) >= 0) {
    return 'Audit';
  }
  return 'Metadata';
}

function getPermanentClassRows_() {
  return SVARA_CONFIG.PERMANENT_HEADERS.map(function (h) {
    return [getPermanentCategory_(h), h];
  });
}

function upsertSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  else sheet.clear();
  return sheet;
}

function buildClassificationSheet_(ss) {
  var name = SVARA_CONFIG.CLASSIFICATION_SHEET;
  var sheet = upsertSheet_(ss, name);
  sheet.setTabColor('#0f766e');
  sheet.getRange(1, 1, 1, 4)
    .setValues([['Category', 'Type', 'Section', 'Parameter']])
    .setFontWeight('bold')
    .setBackground('#0f766e')
    .setFontColor('#ffffff');

  var rows = [];
  getPermanentClassRows_().forEach(function (r) {
    rows.push([r[0], 'Permanent', 'Fixed header', r[1]]);
  });
  getRotatingQuestions_().forEach(function (r) {
    rows.push([r[0], 'Rotating', 'Questionnaire', r[1]]);
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.setColumnWidths(1, 4, 180);
  sheet.setColumnWidth(4, 320);
}

function buildCallsDataSheet_(ss) {
  var sheet = upsertSheet_(ss, SVARA_CONFIG.CALLS_DATA_SHEET);
  sheet.setTabColor('#ea4335');
  var headers = ['UID'].concat(getPermanentHeaders_());
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#0f766e')
    .setFontColor('#ffffff');

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
    .setFontWeight('bold')
    .setBackground('#f59e0b')
    .setFontColor('#ffffff');

  var rows = [];
  var src = 2;
  getPermanentClassRows_().forEach(function (r) {
    rows.push([r[1], 'Permanent', r[0], src++]);
  });
  getRotatingQuestions_().forEach(function (r) {
    rows.push([r[1], 'Rotating', r[0], src++]);
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.setColumnWidths(1, 4, 220);
}

function buildAuditSubmissionsSheet_(ss) {
  var sheet = upsertSheet_(ss, SVARA_CONFIG.AUDIT_SUBMISSIONS_SHEET);
  sheet.setTabColor('#34a853');
  var headers = ['Submission Timestamp', 'UID', 'Rotating Section Active']
    .concat(getPermanentHeaders_())
    .concat(getRotatingQuestions_().map(function (r) { return 'Q: ' + r[1]; }));
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#16a34a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}
