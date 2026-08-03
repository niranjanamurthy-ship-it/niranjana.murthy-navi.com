/**
 * Svara QA Audit Form — menu entry point
 */

/**
 * Runs when the spreadsheet is opened — adds custom Audit menu.
 */
function onOpen() {
  buildAuditMenu_();
}

/**
 * Manual menu refresh (run once after pasting scripts).
 */
function buildAuditMenu() {
  buildAuditMenu_();
}

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
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Setup')
        .addItem('Fix All Sheet Issues', 'fixAllSheetIssues')
        .addItem('Initialize Audit Workbook', 'initializeAuditWorkbook')
        .addItem('Create / Refresh Audit Form Sheet', 'createAuditFormSheet')
        .addItem('Segregate Parameters (Col D → summary)', 'segregateParametersFromClassificationSheet')
    )
    .addToUi();
}
