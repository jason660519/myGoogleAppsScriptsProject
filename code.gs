/**
 * 這是一個「連接器腳本」。
 * 它負責建立選單，並將按鈕動作轉發給 MyTools Library。
 *
 * 重要：
 * 1. 使用者必須在「資源」->「程式庫」中加入你的 Script ID。
 * 2. 識別碼 (Identifier) 必須設定為: MyTools
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var my_menu = ui.createMenu("My Custom Tools"); // 選單名稱


  // 建立 Sort 子選單
  var my_submenu_1 = ui.createMenu("Sort");
  my_submenu_1.addItem("SortABC12", "call_sortABC12");
  my_submenu_1.addSeparator();
  my_submenu_1.addItem("SortBCD12", "call_sortBCD12");
  my_submenu_1.addSeparator();
  my_submenu_1.addItem("SortCDE12", "call_sortCDE12");
  my_submenu_1.addSeparator();
  my_submenu_1.addItem("SortDEF12", "call_sortDEF12");
  my_submenu_1.addSeparator();
  my_submenu_1.addItem("SortEFG12", "call_sortEFG12");


  my_menu.addSubMenu(my_submenu_1);


  // 一般功能
  my_menu.addItem("Cut Rows with Keyword", "call_showInputDialog");
  my_menu.addItem("Del Blank Rows", "call_deleteBlankRows");
  my_menu.addItem("Remove Newlines", "call_removeNewlines");
  my_menu.addItem("Arrange Neatly", "call_arrangeNeatly");


  my_menu.addSeparator();


  // AI / 爬蟲功能
  my_menu.addItem("Crawl GitHub Info", "call_crawlGithubInfo");
  my_menu.addItem("LeetCode Explain", "call_explainLeetCode");


  my_menu.addToUi();
}


// --- 以下是轉接函數 (Wrapper Functions) ---
// 這些函數單純是用來呼叫 MyTools Library 裡面的對應功能


function call_sortABC12() {
  MyTools.sortABC12();
}


function call_sortBCD12() {
  MyTools.sortBCD12();
}


function call_sortCDE12() {
  MyTools.sortCDE12();
}


function call_sortDEF12() {
  MyTools.sortDEF12();
}


function call_sortEFG12() {
  MyTools.sortEFG12();
}


function call_showInputDialog() {
  MyTools.showInputDialog();
}


function call_deleteBlankRows() {
  MyTools.deleteBlankRowsAndMoveCursor();
}


function call_removeNewlines() {
  MyTools.removeNewlinesInMultipleCells();
}


function call_arrangeNeatly() {
  MyTools.my_Arrange_neatly();
}


function call_explainLeetCode() {
  MyTools.explainLeetCodeSelection();
}


// 修正後的 GitHub 爬蟲呼叫函數
function call_crawlGithubInfo() {
  // 這裡原本寫錯成 crawlGithubInfo，已修正為 crawlGithubSelection
  if (typeof MyTools.crawlGithubSelection === "function") {
    MyTools.crawlGithubSelection();
  } else {
    SpreadsheetApp.getUi().alert(
      "Library 中找不到 crawlGithubSelection 函數，請確認 Library 版本是否已更新。"
    );
  }
}


