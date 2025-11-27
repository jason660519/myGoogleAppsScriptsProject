function removeNewlinesInMultipleCells() {
  var spreadsheet = SpreadsheetApp.getActive();
  var selection = spreadsheet.getActiveRange(); // 获取当前选中的单元格区域
  var values = selection.getValues(); // 获取选中单元格的所有值

  // 遍历选中区域的每个单元格，对每个字符串值移除换行符
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        values[i][j] = values[i][j].replace(/\n/g, ''); // 移除换行符
      }
    }
  }

  // 将处理过的值设置回相同的单元格范围
  selection.setValues(values);
}
