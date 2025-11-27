function deleteBlankRowsAndMoveCursor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var values = range.getValues();
  var numRowsDeleted = 0;

  for (var i = lastRow - 1; i >= 0; i--) {
    var row = values[i];
    var isBlank = true;

    for (var j = 0; j < lastColumn; j++) {
      if (row[j] !== "") {
        isBlank = false;
        break;
      }
    }

    if (isBlank) {
      sheet.deleteRow(i + 1);
      numRowsDeleted++;
    }
  }

  // Move cursor to the first cell of the last row
  var lastRowNew = sheet.getLastRow();
  var firstCell = sheet.getRange(lastRowNew, 1);
  sheet.setActiveRange(firstCell);

  // Show a message with the number of deleted rows
  SpreadsheetApp.getUi().alert("Deleted " + numRowsDeleted + " row(s) with all blank cells.");
}

