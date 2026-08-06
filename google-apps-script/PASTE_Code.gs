/**
 * SVARA QA AUDIT — paste into Code.gs
 * Also add HTML file named exactly: AuditForm (from PASTE_AuditForm.html)
 * Then run: installEverything
 * Sheet: https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ
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
    .addItem('Fix All Sheet Issues', 'fixAllSheetIssues')
    .addItem('Fix Timestamp Circular Dependency', 'fixTimestampCircularDependency')
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

/** Converts 1-based column index to A1 letter(s). */
function columnLetter_(col) {
  var temp = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

/**
 * Finds a header column by name (case-insensitive).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string|string[]} names
 * @returns {number} 1-based column index, or 0 if not found
 */
function findHeaderColumn_(sheet, names) {
  var aliases = Array.isArray(names) ? names : [names];
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    for (var a = 0; a < aliases.length; a++) {
      if (h === String(aliases[a] || '').trim().toLowerCase()) return c + 1;
    }
  }
  return 0;
}

/** True when a formula references the given column letter (e.g. self-reference loop). */
function formulaReferencesColumn_(formula, colLetter) {
  if (!formula) return false;
  var col = String(colLetter || '').toUpperCase();
  if (!col) return false;
  return new RegExp('\\$?' + col + '(\\$?\\d+|\\$?:)', 'i').test(String(formula));
}

/**
 * Repairs Timestamp / Week_Num formulas that reference their own column (#REF circular dependency).
 * Timestamp is rebuilt from Audit date, then Call Date — never from Timestamp itself.
 */
function fixCircularTimestampColumns_(ss, fixes) {
  ss.getSheets().forEach(function (sheet) {
    var tsCol = findHeaderColumn_(sheet, ['Timestamp']);
    if (!tsCol) return;

    var auditCol = findHeaderColumn_(sheet, ['Audit date', 'Audit Date']);
    var callCol = findHeaderColumn_(sheet, ['Call Date', 'Call date']);
    var weekCol = findHeaderColumn_(sheet, ['Week_Num', 'Week Num', 'WeekNum']);
    var tsLetter = columnLetter_(tsCol);
    var lastRow = Math.max(sheet.getLastRow(), 2);
    var tsCell = sheet.getRange(2, tsCol);
    var tsFormula = tsCell.getFormula();
    var tsDisplay = String(tsCell.getDisplayValue() || '');
    var needsFix = tsDisplay.indexOf('#') === 0
      || (tsFormula && formulaReferencesColumn_(tsFormula, tsLetter));

    if (!needsFix && lastRow > 2) {
      var sample = sheet.getRange(2, tsCol, Math.min(lastRow - 1, 6), 1);
      var formulas = sample.getFormulas();
      var displays = sample.getDisplayValues();
      for (var i = 0; i < formulas.length; i++) {
        if (String(displays[i][0] || '').indexOf('#') === 0) { needsFix = true; break; }
        if (formulas[i][0] && formulaReferencesColumn_(formulas[i][0], tsLetter)) { needsFix = true; break; }
      }
    }

    if (!needsFix) return;

    sheet.getRange(2, tsCol, Math.max(lastRow - 1, 1), 1).clearContent();

    if (!auditCol && !callCol) {
      fixes.push('Cleared circular Timestamp formulas on "' + sheet.getName()
        + '" — add Audit date or Call Date columns, then re-run this fix');
      return;
    }

    var fAudit = auditCol ? '$' + columnLetter_(auditCol) + '2:$' + columnLetter_(auditCol) : '';
    var fCall = callCol ? '$' + columnLetter_(callCol) + '2:$' + columnLetter_(callCol) : '';
    var rebuiltFormula;
    if (auditCol && callCol) {
      rebuiltFormula = '=ARRAYFORMULA(IF(LEN(' + fAudit + '), ' + fAudit + ', IF(LEN(' + fCall + '), ' + fCall + ', )))';
    } else if (auditCol) {
      rebuiltFormula = '=ARRAYFORMULA(IF(LEN(' + fAudit + '), ' + fAudit + ', ))';
    } else {
      rebuiltFormula = '=ARRAYFORMULA(IF(LEN(' + fCall + '), ' + fCall + ', ))';
    }
    tsCell.setFormula(rebuiltFormula);
    fixes.push('Fixed circular Timestamp dependency on "' + sheet.getName() + '"');

    if (weekCol && callCol) {
      var weekCell = sheet.getRange(2, weekCol);
      var weekLetter = columnLetter_(weekCol);
      var weekFormula = weekCell.getFormula();
      if (!weekFormula
        || formulaReferencesColumn_(weekFormula, tsLetter)
        || formulaReferencesColumn_(weekFormula, weekLetter)
        || String(weekCell.getDisplayValue() || '').indexOf('#') === 0) {
        var fCallWeek = '$' + columnLetter_(callCol) + '2:$' + columnLetter_(callCol);
        weekCell.setFormula('=ARRAYFORMULA(IF(LEN(' + fCallWeek + '), WEEKNUM(' + fCallWeek + '), ))');
        fixes.push('Rebuilt Week_Num from Call Date on "' + sheet.getName() + '"');
      }
    }
  });
}

/** Standalone repair for Timestamp #REF / circular dependency errors. */
function fixTimestampCircularDependency() {
  var fixes = [];
  fixCircularTimestampColumns_(SpreadsheetApp.getActiveSpreadsheet(), fixes);
  SpreadsheetApp.getUi().alert(
    fixes.length
      ? 'Fixed ' + fixes.length + ' issue(s):\n\n• ' + fixes.join('\n• ') + '\n\nReload the sheet.'
      : 'No circular Timestamp formulas found.'
  );
}

/** Repairs common sheet issues without wiping Calls Data or submissions. */
function fixAllSheetIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fixes = [];

  fixCircularTimestampColumns_(ss, fixes);
  fixClassificationSheet_(ss, fixes);
  fixCallsDataHeaders_(ss, fixes);
  buildParameterClassificationSheet_(ss);
  fixes.push('Refreshed Parameter Classification summary');
  createAuditFormSheet_();
  fixes.push('Refreshed Svara Audit Form layout and dropdowns');
  fixAuditSubmissionsHeaders_(ss, fixes);
  buildAuditMenu_();
  fixes.push('Rebuilt Audit menu');

  SpreadsheetApp.getUi().alert(
    'Fixed ' + fixes.length + ' issue(s):\n\n• ' + fixes.join('\n• ') +
    '\n\nReload the sheet, then test Audit → Open Standalone Audit Form with UID SV-10001.'
  );
}

function fixClassificationSheet_(ss, fixes) {
  var cfg = SVARA_CONFIG;
  var sheet = ss.getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!sheet) {
    buildClassificationSheet_(ss);
    fixes.push('Created missing "' + cfg.CLASSIFICATION_SHEET + '" tab');
    return;
  }

  var header = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (String(header[0] || '') !== 'Category' || String(header[3] || '') !== 'Parameter') {
    sheet.getRange(1, 1, 1, 4).setValues([['Category', 'Type', 'Section', 'Parameter']])
      .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    fixes.push('Corrected classification sheet headers (A1:D1)');
  }

  var lastRow = Math.max(sheet.getLastRow(), cfg.CLASSIFICATION_DATA_START_ROW - 1);
  var numRows = lastRow - cfg.CLASSIFICATION_DATA_START_ROW + 1;
  if (numRows < 1) numRows = 0;

  var existing = {};
  if (numRows > 0) {
    var params = sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.PARAM_COLUMN, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      var label = String(params[i][0] || '').trim();
      if (label) existing[label.toLowerCase()] = cfg.CLASSIFICATION_DATA_START_ROW + i;
    }
  }

  var missing = [];
  cfg.PERMANENT_HEADERS.forEach(function (h) {
    if (!existing[h.toLowerCase()]) missing.push(h);
  });

  if (missing.length > 0) {
    var insertAt = lastRow + 1;
    var newRows = missing.map(function (h) {
      return [getPermanentCategory_(h), 'Permanent', 'Fixed header', h];
    });
    sheet.getRange(insertAt, 1, newRows.length, 4).setValues(newRows);
    fixes.push('Added ' + missing.length + ' missing permanent parameter(s): ' + missing.join(', '));
  }

  // Re-mark known permanent headers that were wrongly typed as Rotating
  if (numRows > 0) {
    var types = sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.TYPE_COLUMN, numRows, 1).getValues();
    var params2 = sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.PARAM_COLUMN, numRows, 1).getValues();
    var retyped = 0;
    for (var j = 0; j < numRows; j++) {
      var p = String(params2[j][0] || '').trim();
      if (!p || cfg.PERMANENT_HEADERS.indexOf(p) === -1) continue;
      var t = String(types[j][0] || '').toLowerCase();
      if (t.indexOf('perman') === -1) {
        sheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW + j, cfg.TYPE_COLUMN).setValue('Permanent');
        retyped++;
      }
    }
    if (retyped > 0) {
      fixes.push('Re-typed ' + retyped + ' misclassified permanent parameter(s) in column B');
    }
  }
}

function fixCallsDataHeaders_(ss, fixes) {
  var cfg = SVARA_CONFIG;
  var sheet = ss.getSheetByName(cfg.CALLS_DATA_SHEET);
  if (!sheet) {
    buildCallsDataSheet_(ss);
    fixes.push('Created missing "' + cfg.CALLS_DATA_SHEET + '" tab with sample UIDs');
    return;
  }

  var expected = ['UID'].concat(cfg.PERMANENT_HEADERS);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });

  var missingHeaders = expected.filter(function (h) {
    return headers.indexOf(h) === -1;
  });

  if (missingHeaders.length > 0) {
    var newHeaders = headers.slice();
    missingHeaders.forEach(function (h) { newHeaders.push(h); });
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders])
      .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    fixes.push('Added missing Calls Data column(s): ' + missingHeaders.join(', '));
  }

  if (sheet.getLastRow() < 2) {
    buildCallsDataSheet_(ss);
    fixes.push('Added sample call rows (SV-10001, SV-10002, SV-10003)');
  }
}

function fixAuditSubmissionsHeaders_(ss, fixes) {
  var cfg = SVARA_CONFIG;
  var sheet = ss.getSheetByName(cfg.AUDIT_SUBMISSIONS_SHEET);
  if (!sheet) {
    buildAuditSubmissionsSheet_(ss);
    fixes.push('Created missing "' + cfg.AUDIT_SUBMISSIONS_SHEET + '" tab');
    return;
  }

  var expected = ['Submission Timestamp', 'UID', 'Rotating Section Active']
    .concat(cfg.PERMANENT_HEADERS)
    .concat(getRotatingQuestions_().map(function (r) { return 'Q: ' + r[1]; }));

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); })
    : [];

  var missing = expected.filter(function (h) { return headers.indexOf(h) === -1; });
  if (headers.length === 0 || missing.length > 0) {
    var merged = headers.slice();
    missing.forEach(function (h) { if (merged.indexOf(h) === -1) merged.push(h); });
    if (merged.length === 0) merged = expected;
    sheet.getRange(1, 1, 1, merged.length).setValues([merged])
      .setFontWeight('bold').setBackground('#16a34a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    fixes.push('Updated Audit Submissions headers (' + missing.length + ' column(s) added)');
  }
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
