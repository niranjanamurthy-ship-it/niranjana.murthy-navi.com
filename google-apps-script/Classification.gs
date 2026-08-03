/**
 * Reads rotating / permanent parameters from the classification sheet.
 */

/**
 * @returns {{permanent: string[], rotating: string[]}}
 */
function getClassifiedParameters_() {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.CLASSIFICATION_SHEET);
  if (!sheet) {
    throw new Error(
      'Sheet not found: "' + cfg.CLASSIFICATION_SHEET + '". Run Setup → Initialize Audit Workbook.'
    );
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < cfg.CLASSIFICATION_DATA_START_ROW) {
    return { permanent: [], rotating: [] };
  }

  var numRows = lastRow - cfg.CLASSIFICATION_DATA_START_ROW + 1;
  var paramRange = sheet.getRange(
    cfg.CLASSIFICATION_DATA_START_ROW,
    cfg.PARAM_COLUMN,
    numRows,
    1
  ).getValues();
  var typeRange = sheet.getRange(
    cfg.CLASSIFICATION_DATA_START_ROW,
    cfg.TYPE_COLUMN,
    numRows,
    1
  ).getValues();

  var permanent = [];
  var rotating = [];
  var seen = {};

  for (var i = 0; i < numRows; i++) {
    var label = String(paramRange[i][0] || '').trim();
    if (!label || seen[label]) continue;
    seen[label] = true;

    var typeText = String(typeRange[i][0] || '').trim().toLowerCase();
    if (typeText.indexOf('rotat') !== -1) {
      rotating.push(label);
    } else if (typeText.indexOf('perman') !== -1) {
      permanent.push(label);
    } else {
      // Default: treat as rotating questionnaire item when type is blank
      rotating.push(label);
    }
  }

  return { permanent: permanent, rotating: rotating };
}

/**
 * Exposed for HTML sidebar — returns classification as JSON.
 */
function getClassificationForForm() {
  return getClassifiedParameters_();
}
