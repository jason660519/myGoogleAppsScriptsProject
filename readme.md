# a0405_GoogleSheetTools — 安裝與使用指南

本文件說明如何在 Google Sheets 專案中安裝並使用你的 Library，以及如何設定必要的專案屬性（API Keys）與操作說明。

## 專案資訊
- 專案標題（Project title）：`a0405_GoogleSheetTools`
- Library Script ID：`1qQxlTXWrRpbowAvq6ya_WrO_Z9BJYrglp7eoBePm_fw2c_8LDpfM1xdE`
- 建議識別碼（Identifier）：`MyTools`

## 安裝步驟（將 Library 導入你的 Google Sheet）
1. 在 Google Sheets 中，開啟「擴充功能」→「Apps Script」。
2. 在 Apps Script 編輯器中，開啟「資源」→「程式庫」（Libraries）。
3. 貼上 Library Script ID：`1qQxlTXWrRpbowAvq6ya_WrO_Z9BJYrglp7eoBePm_fw2c_8LDpfM1xdE`。
4. 設定識別碼（Identifier）為 `MyTools`，選擇最新版本並新增。
5. 新增檔案 `code.gs`，將本倉庫的 `code.gs` 內容完整複製貼上到你的專案中。
6. 儲存後重載試算表，於選單列應出現「My Custom Tools」。

> 備註：若你使用新版 Apps Script UI，Libraries 入口可能位於「Project Settings」或介面不同，請依 UI 指示找到 Libraries 功能並加入 Script ID。

## 需要設定的 Script Properties（API 金鑰）
若要啟用 AI／爬蟲等功能，請在「Project Settings」→「Script properties」中新增以下鍵值：
- `OPENAI_API_KEY`：OpenAI 金鑰
- `DEEPSEEK_API_KEY`：DeepSeek 金鑰
- （如有使用）`GEMINI_API_KEY`：Google Gemini 金鑰

> 建議：避免將金鑰硬編碼在程式中。使用 `PropertiesService.getScriptProperties()` 於程式執行期讀取，以提升安全性與可維護性。

## 使用方式（功能一覽）
`code.gs` 會在開啟試算表時建立自訂選單，並把選單動作轉接至 Library `MyTools`。

- 選單建立：`onOpen()` 於 `code.gs:11` 建立「My Custom Tools」選單並掛載各功能。
- 排序功能（Sort）：
  - `call_sortABC12()` → `MyTools.sortABC12()`（`code.gs:55`）
  - `call_sortBCD12()` → `MyTools.sortBCD12()`（`code.gs:60`）
  - `call_sortCDE12()` → `MyTools.sortCDE12()`（`code.gs:65`）
  - `call_sortDEF12()` → `MyTools.sortDEF12()`（`code.gs:70`）
  - `call_sortEFG12()` → `MyTools.sortEFG12()`（`code.gs:75`）
- 一般工具：
  - `call_showInputDialog()` → `MyTools.showInputDialog()`（`code.gs:80`）
  - `call_deleteBlankRows()` → `MyTools.deleteBlankRowsAndMoveCursor()`（`code.gs:85`）
  - `call_removeNewlines()` → `MyTools.removeNewlinesInMultipleCells()`（`code.gs:90`）
  - `call_arrangeNeatly()` → `MyTools.my_Arrange_neatly()`（`code.gs:95`）
- AI／爬蟲：
  - `call_explainLeetCode()` → `MyTools.explainLeetCodeSelection()`（`code.gs:100`）
  - `call_crawlGithubInfo()` 會呼叫 `MyTools.crawlGithubSelection()`（`code.gs:105`）。若 Library 還未提供該函式，將跳出提示。

## 常見問題（FAQ）
- 看不到「My Custom Tools」選單
  - 重新載入試算表或在 Apps Script 中執行一次 `onOpen()`。
- 出現 `MyTools is not defined`
  - 請確認已在「程式庫」加入 Script ID，識別碼為 `MyTools`，且已選擇有效版本。
- 提示「Library 中找不到 crawlGithubSelection」
  - 請更新 Library 至最新版本，或與維護者確認該函式是否已發佈。
- 呼叫 AI 功能失敗
  - 請確認 `OPENAI_API_KEY` 或 `DEEPSEEK_API_KEY` 已於 Script Properties 設定且有效。
- 權限問題
  - 首次執行可能會要求授權（例如使用 `SpreadsheetApp`、外部 HTTP 請求等），請按提示授予權限。

## 進階：以 clasp 管理專案（選用）
若你希望在本機端管理 Apps Script 專案，`clasp` 可讓你在本地開發並推送到雲端：

1. 安裝：`npm install -g @google/clasp`
2. 啟用 Apps Script API：在 `https://script.google.com/home/usersettings` 開啟。
3. 登入：`clasp login`
4. 拉取／推送：
   - `clasp pull`：從雲端拉取最新程式
   - `clasp push`：推送本地變更到雲端
5. 忽略規則：在 `.claspignore` 中加入不需要推送的檔案，例如：
   ```
   code.gs
   readme.md
   ```

## 更新與版本管理建議
- 當 Library 更新後，請在「程式庫」調整版本為最新，避免介面不一致造成錯誤。
- 如需自訂選單名稱或顯示項目，可編輯 `code.gs` 中的 `onOpen()`（`code.gs:11`）與對應的 `call_*` 函式。

## 聯絡維護者
- 如需新增功能或回報問題，請於此 GitHub 倉庫建立 Issue。

