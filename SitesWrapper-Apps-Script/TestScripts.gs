/*
 * Displays "PASSED" in the sidebar panel with a return button.
 */
function testPassed() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('TestPassed'));
}

/*
 * Displays "FAILED" in the sidebar panel with a return button.
 */
function testFailed() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('TestFailed'));
}

/*
 * SitesWrapper_GAS_1
 *
 * Bind and initialize the spreadsheet and update the initial configuration in the datastore.
 */
function SitesWrapper_GAS_1() {
  document.insertSheet('Empty', 0);
  document.setActiveSheet(document.getSheets()[0]);
  var numSheets = 6;
  for (sheetNum = 0; sheetNum < numSheets; sheetNum++) {
    document.deleteSheet(document.getSheets()[1]);
  }
  if (SpreadsheetApp.getActiveSpreadsheet().getSheets().length == 1) {
    testPassed();
  } else {
    testFailed();
  }
}

/*
 * 
 */
function SitesWrapper_GAS_2() {
  testPassed();
}

/*
 * 
 */
function SitesWrapper_GAS_3() {
  testFailed();
}

