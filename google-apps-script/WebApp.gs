/**
 * Standalone Svara Audit Form — full-page web app (separate from sidebar).
 */

/**
 * Serves the standalone audit form when deployed as a Web App.
 * Deploy: Deploy → New deployment → Web app → Execute as: Me → Anyone with link
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('StandaloneAuditForm')
    .setTitle('Svara QA Audit Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Opens the standalone form in a large modal dialog (inside the spreadsheet).
 */
function openStandaloneAuditForm() {
  var html = HtmlService.createHtmlOutputFromFile('StandaloneAuditForm')
    .setWidth(920)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Svara QA Audit Form');
}
