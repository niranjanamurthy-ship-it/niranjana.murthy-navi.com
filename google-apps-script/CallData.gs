/**
 * Fetch call metadata by unique identity number (UID).
 */

/**
 * Builds a header → column index map from Calls Data row 1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, number>}
 */
function buildCallsHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h) map[h] = c + 1;
  }
  return map;
}

/**
 * Normalizes UID for comparison.
 * @param {*} value
 * @returns {string}
 */
function normalizeUid_(value) {
  return String(value == null ? '' : value).trim();
}

/**
 * @param {string} uid Unique identity number
 * @returns {Object|null} Row data keyed by header name, or null if not found
 */
function fetchCallDataByUid(uid) {
  var cfg = SVARA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.CALLS_DATA_SHEET);
  if (!sheet) {
    throw new Error(
      'Sheet not found: "' + cfg.CALLS_DATA_SHEET + '". Run Setup → Initialize Audit Workbook.'
    );
  }

  var searchUid = normalizeUid_(uid);
  if (!searchUid) {
    throw new Error('Please enter a Unique Identity Number (UID).');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var headerMap = buildCallsHeaderMap_(sheet);
  var uidCol = cfg.CALLS_UID_COLUMN;
  var uidValues = sheet.getRange(2, uidCol, lastRow - 1, 1).getValues();

  for (var r = 0; r < uidValues.length; r++) {
    if (normalizeUid_(uidValues[r][0]) === searchUid) {
      var rowNum = r + 2;
      var lastCol = sheet.getLastColumn();
      var rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
      var result = { _uid: searchUid, _row: rowNum };
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var c = 0; c < headers.length; c++) {
        var key = String(headers[c] || '').trim();
        if (key) result[key] = rowValues[c];
      }
      return result;
    }
  }
  return null;
}

/**
 * Maps call row onto permanent form fields (fuzzy header match).
 * @param {Object} callData
 * @returns {Object.<string, *>}
 */
function mapCallDataToPermanentFields_(callData) {
  if (!callData) return {};
  var cfg = SVARA_CONFIG;
  var mapped = {};
  var keys = Object.keys(callData);

  cfg.PERMANENT_HEADERS.forEach(function (header) {
    if (callData[header] !== undefined) {
      mapped[header] = callData[header];
      return;
    }
    var lowerHeader = header.toLowerCase();
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (k.charAt(0) === '_') continue;
      if (k.toLowerCase() === lowerHeader) {
        mapped[header] = callData[k];
        return;
      }
    }
    mapped[header] = '';
  });

  return mapped;
}

/**
 * Used by sidebar and in-sheet fetch control.
 * @param {string} uid
 * @returns {Object}
 */
function fetchCallForAuditForm(uid) {
  var callData = fetchCallDataByUid(uid);
  if (!callData) {
    return { found: false, uid: normalizeUid_(uid), fields: {} };
  }
  return {
    found: true,
    uid: callData._uid,
    fields: mapCallDataToPermanentFields_(callData),
    raw: callData
  };
}
