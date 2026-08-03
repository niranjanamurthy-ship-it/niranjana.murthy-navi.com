/**
 * One-time workbook initialization and classification helpers.
 */

/**
 * Full setup: Calls Data template, refresh classification summary, create audit form.
 */
function initializeAuditWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCallsDataSheet_(ss);
  createClassificationSummarySheet_(ss);
  createAuditFormSheet();
}

/**
 * Adds Rotating / Permanent markers on classification sheet if missing.
 * Reads column D parameters and writes summary to "Parameter Classification".
 */
function segregateParametersFromClassificationSheet() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!source) {
    SpreadsheetApp.getUi().alert(
      'Cannot find "' + cfg.CLASSIFICATION_SHEET + '". Rename your tab to match or update Config.gs.'
    );
    return;
  }

  createClassificationSummarySheet_(ss);
  var classified = getClassifiedParameters_();

  SpreadsheetApp.getUi().alert(
    'Parameters segregated from column D:\n\n' +
    'Permanent: ' + classified.permanent.length + '\n' +
    'Rotating: ' + classified.rotating.length + '\n\n' +
    'See sheet "Parameter Classification" for the full list.'
  );
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createClassificationSummarySheet_(ss) {
  var cfg = SVARA_CONFIG;
  var name = 'Parameter Classification';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }

  sheet.getRange('A1:D1').setValues([['Parameter (Col D)', 'Type', 'Category', 'Source Row']])
    .setFontWeight('bold').setBackground('#d9ead3');

  var source = ss.getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!source) return;

  var lastRow = source.getLastRow();
  if (lastRow < cfg.CLASSIFICATION_DATA_START_ROW) return;

  var numRows = lastRow - cfg.CLASSIFICATION_DATA_START_ROW + 1;
  var rows = [];
  for (var i = 0; i < numRows; i++) {
    var r = cfg.CLASSIFICATION_DATA_START_ROW + i;
    var param = String(source.getRange(r, cfg.PARAM_COLUMN).getValue() || '').trim();
    if (!param) continue;
    var typeVal = String(source.getRange(r, cfg.TYPE_COLUMN).getValue() || '').trim();
    var typeLower = typeVal.toLowerCase();
    var type;
    if (typeLower.indexOf('perman') !== -1) type = 'Permanent';
    else if (typeLower.indexOf('rotat') !== -1) type = 'Rotating';
    else type = 'Rotating (default)';

    var category = String(source.getRange(r, cfg.CATEGORY_COLUMN).getValue() || '').trim();
    rows.push([param, type, category, r]);
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  sheet.autoResizeColumns(1, 4);
  sheet.setTabColor('#fbbc04');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function ensureCallsDataSheet_(ss) {
  var cfg = SVARA_CONFIG;
  var sheet = ss.getSheetByName(cfg.CALLS_DATA_SHEET);
  if (sheet) return;

  sheet = ss.insertSheet(cfg.CALLS_DATA_SHEET);
  var headers = ['UID'].concat(cfg.PERMANENT_HEADERS);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setTabColor('#ea4335');

  // Sample row for testing
  var sample = ['SV-10001'];
  cfg.PERMANENT_HEADERS.forEach(function (h, i) {
    if (h === 'LAN') sample.push('LAN-001');
    else if (h === 'Call Date') sample.push(new Date());
    else if (h === 'Audit date') sample.push('');
    else if (h === 'QA Name') sample.push('');
    else if (h === 'Call Quality Assessment') sample.push('');
    else if (h === 'Call Duration') sample.push('03:42');
    else if (h === 'Recording Link') sample.push('https://example.com/rec/SV-10001');
    else if (h === 'DPD') sample.push('15');
    else sample.push('Sample ' + (i + 1));
  });
  sheet.getRange(2, 1, 1, headers.length).setValues([sample]);
}

/** Repairs common sheet issues without wiping Calls Data or submissions. */
function fixAllSheetIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = SVARA_CONFIG;
  var fixes = [];

  var classSheet = ss.getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!classSheet) {
    SpreadsheetApp.getUi().alert(
      'Cannot find "' + cfg.CLASSIFICATION_SHEET + '". Rename your tab or run Initialize Audit Workbook first.'
    );
    return;
  }

  var lastRow = Math.max(classSheet.getLastRow(), cfg.CLASSIFICATION_DATA_START_ROW - 1);
  var numRows = lastRow - cfg.CLASSIFICATION_DATA_START_ROW + 1;
  var existing = {};
  if (numRows > 0) {
    var params = classSheet.getRange(cfg.CLASSIFICATION_DATA_START_ROW, cfg.PARAM_COLUMN, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      var label = String(params[i][0] || '').trim();
      if (label) existing[label.toLowerCase()] = true;
    }
  }

  var missing = cfg.PERMANENT_HEADERS.filter(function (h) { return !existing[h.toLowerCase()]; });
  if (missing.length > 0) {
    var insertAt = lastRow + 1;
    var newRows = missing.map(function (h) {
      var cat = 'Audit';
      if (['LAN', 'Call Date', 'Call Duration', 'Recording Link', 'DPD', 'Experiment Tag'].indexOf(h) >= 0) {
        cat = 'Metadata';
      } else if (h.indexOf('Language') >= 0) {
        cat = 'Language';
      } else if (h.indexOf('Disposition') >= 0) {
        cat = 'Outcome';
      }
      return [cat, 'Permanent', 'Fixed header', h];
    });
    classSheet.getRange(insertAt, 1, newRows.length, 4).setValues(newRows);
    fixes.push('Added missing permanent parameters: ' + missing.join(', '));
  }

  ensureCallsDataSheet_(ss);
  fixes.push('Verified Calls Data tab and sample UID row');

  createClassificationSummarySheet_(ss);
  fixes.push('Refreshed Parameter Classification summary');

  createAuditFormSheet();
  fixes.push('Refreshed Svara Audit Form layout');

  var submissions = ss.getSheetByName(cfg.AUDIT_SUBMISSIONS_SHEET);
  if (!submissions) {
    submissions = ss.insertSheet(cfg.AUDIT_SUBMISSIONS_SHEET);
    submissions.setTabColor('#34a853');
    fixes.push('Created Audit Submissions tab');
  }

  buildAuditMenu_();
  fixes.push('Rebuilt Audit menu');

  SpreadsheetApp.getUi().alert(
    'Fixed ' + fixes.length + ' issue(s):\n\n• ' + fixes.join('\n• ') +
    '\n\nReload the sheet, then test with UID SV-10001.'
  );
}
