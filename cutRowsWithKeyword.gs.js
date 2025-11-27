// 定義一個函數，用於顯示輸入對話框並獲取用戶輸入的關鍵字
function showInputDialog() {
  // 調用getKeywordFromUser函數，並將提示信息作為參數傳入
  var keywordObj = getKeywordFromUser("Enter the keyword to search:");
  // 如果用戶輸入了有效的關鍵字（非空），則調用cutRowsWithKeyword函數進行處理
  if (keywordObj !== null) {
    cutRowsWithKeyword(keywordObj);
  }
}

// 定義一個函數，用於從用戶那裡獲取輸入的關鍵字
function getKeywordFromUser(promptText) {
  var ui = SpreadsheetApp.getUi();
  
  // 顯示對話框
  var response = ui.prompt(promptText, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var keyword = response.getResponseText();
    
    // 確認是否考慮大小寫
    var caseResponse = ui.alert('Do you want the search to be case sensitive?', ui.ButtonSet.YES_NO);
    var isCaseSensitive = caseResponse === ui.Button.YES;
    
    if (keyword.trim() !== "") {
      return { keyword: keyword, isCaseSensitive: isCaseSensitive };
    } else {
      ui.alert("Keyword cannot be empty.");
      return null;
    }
  } else {
    ui.alert("Keyword input canceled.");
    return null;
  }
}


// 定義一個函數，用於處理符合關鍵字的行
function cutRowsWithKeyword(keywordObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var data = dataRange.getValues();
  var targetRows = [];
  
  var keyword = keywordObj.keyword;
  var isCaseSensitive = keywordObj.isCaseSensitive;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var foundKeyword = false;
    for (var j = 0; j < row.length; j++) {
      var cellValue = row[j].toString();
      if (!isCaseSensitive) {
        keyword = keyword.toLowerCase();
        cellValue = cellValue.toLowerCase();
      }
      if (cellValue.indexOf(keyword) !== -1) {
        foundKeyword = true;
        break;
      }
    }
    if (foundKeyword) {
      targetRows.push(i + 2); // +2 because we start from the second row
    }
  }

  if (targetRows.length > 0) {
    var newLastRow = sheet.getLastRow() + 1;
    for (var i = 0; i < targetRows.length; i++) {
      var sourceRange = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn());
      var targetRange = sheet.getRange(newLastRow, 1);
      sourceRange.copyTo(targetRange, { contentsOnly: false });
      newLastRow++;
    }
    for (var i = targetRows.length - 1; i >= 0; i--) {
      sheet.deleteRow(targetRows[i]);
    }
  } else {
    var newKeywordObj = getKeywordFromUser("No matching rows found. Enter the keyword to search again:");
    if (newKeywordObj !== null) {
      cutRowsWithKeyword(newKeywordObj);
    }
  }
}
