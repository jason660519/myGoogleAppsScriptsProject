function sortABC12() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 3, ascending: true}]);
  range.setFontSize(12);
  sheet.setActiveRange(sheet.getRange(lastRow, 1));
}


function sortBCD12() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 2, ascending: true}, {column: 3, ascending: true}, {column: 4, ascending: true}]);
  range.setFontSize(12);
  sheet.setActiveRange(sheet.getRange(lastRow, 2));
}


function sortCDE12() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 3, ascending: true}, {column: 4, ascending: true}, {column: 5, ascending: true}]);
  range.setFontSize(12);
  sheet.setActiveRange(sheet.getRange(lastRow, 3));
}


function sortDEF12() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 4, ascending: true}, {column: 5, ascending: true}, {column: 6, ascending: true}]);
  range.setFontSize(12);
  sheet.setActiveRange(sheet.getRange(lastRow, 4));
}


function sortEFG12() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 5, ascending: true}, {column: 6, ascending: true}, {column: 7, ascending: true}]);
  range.setFontSize(12);
  sheet.setActiveRange(sheet.getRange(lastRow, 5));
}


