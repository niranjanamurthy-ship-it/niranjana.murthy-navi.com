/**
 * MINIMAL FIX — Timestamp #REF circular dependency
 *
 * Paste this entire file into Extensions → Apps Script (Code.gs is fine),
 * then Run → fixTimestampCircularDependency → Allow → reload sheet.
 *
 * Sheet: https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ
 */

function columnLetter_(col) {
  var temp = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

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

function formulaReferencesColumn_(formula, colLetter) {
  if (!formula) return false;
  var col = String(colLetter || '').toUpperCase();
  if (!col) return false;
  return new RegExp('\\$?' + col + '(\\$?\\d+|\\$?:)', 'i').test(String(formula));
}

function fixTimestampCircularDependency() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fixes = [];

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
      fixes.push('Cleared circular Timestamp on "' + sheet.getName()
        + '" — add Audit date or Call Date columns, then re-run');
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
    fixes.push('Fixed Timestamp on "' + sheet.getName() + '"');

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
        fixes.push('Fixed Week_Num on "' + sheet.getName() + '"');
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    fixes.length
      ? 'Fixed:\n\n• ' + fixes.join('\n• ') + '\n\nReload the sheet.'
      : 'No circular Timestamp formulas found.'
  );
}
