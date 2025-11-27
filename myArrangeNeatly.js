function my_Arrange_neatly() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  // 激活整个工作表的实际用到的范围
  var dataRange = sheet.getDataRange();
  dataRange.activate();

  // 应用格式化设置
  spreadsheet.getActiveRangeList().setFontSize(12)
    .setFontFamily(null) // 如果您想设置特定字体，可以替换 'null' 为字体名称，如 'Arial'
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    .setFontColor(null)
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);

  // 获取实际使用的最后一行
  var lastRow = sheet.getLastRow();

  // 设置所有行的高度为 22 (覆盖之前的设置)
  sheet.setRowHeights(1, lastRow, 22);

  // 将鼠标移动到最后一行的第5格
  sheet.setActiveRange(sheet.getRange(lastRow, 5));
}