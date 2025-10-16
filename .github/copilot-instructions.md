# Copilot Instructions for PMS

## 專案架構與主要模組
- 本專案為「治具管理與SOP系統」，以多個 Python 檔案模組化，涵蓋 SOP、治具、帳號、紀錄、資料庫等功能。
- 主要檔案：
  - `PMS.py`：主程式入口，整合各分頁與功能模組。
  - `fixture_tabs.py`、`fixture_helper.py`、`fixture_bom_tab.py`、`fixture_logs_tab.py`：治具管理、BOM、紀錄等分頁與邏輯。
  - `sop_build_tab.py`、`changelog_tab.py`：SOP 與改版歷程分頁。
  - `account_management_tab.py`：帳號與權限管理。
  - `db_helper.py`、`schema_helper.py`：資料庫存取、結構維護。
  - `utils.py`、`fixture_logger.py`：共用工具與日誌。
  - `PMS_Tool/`：資料庫維護、備份、遷移、檢查等工具腳本。

## 關鍵開發流程
- **啟動**：以 `PMS_launcher.bat` 或直接執行 `PMS.py` 啟動主程式。
- **資料庫**：預設使用 SQLite，所有存取請透過 `db_helper.py` 的 get_conn()，避免多重連線分歧。
- **資料表 schema**：schema 由 `schema_helper.py` 定義與自動 ensure，勿於其他模組硬編 schema。
- **日誌/紀錄**：所有異動紀錄請呼叫 `fixture_logger.py` 或 `fixture_logs_tab.py` 相關函式。
- **SOP/治具操作**：分頁邏輯集中於對應 *_tab.py 檔案，UI 事件與資料操作分離。
- **工具腳本**：`PMS_Tool/` 目錄下腳本僅供維護、遷移、備份、檢查資料庫用，勿於主程式直接 import。

## 專案慣例與注意事項
- **資料庫連線**：統一用 `get_conn()`，勿自行開啟 sqlite3 連線。
- **欄位命名**：資料表欄位、主鍵、外鍵命名請參考 `schema_helper.py`，保持一致。
- **分頁刷新**：資料異動後，務必呼叫對應 refresh_all() 以同步 UI。
- **防呆/防重複**：重要操作（如新增、入庫、轉倉）需加 @safe_button_action decorator，避免重複觸發。
- **版本管理**：改版歷程記錄於 `changelog_tab.py`，版本號於 `version.txt`。
- **多執行緒/鎖**：如需跨執行緒存取 DB，請參考現有 lock 實作，避免 deadlock。

## 測試與除錯
- **測試腳本**：目前無自動化測試，請以手動操作 UI 驗證。
- **資料庫備份/還原**：請用 `PMS_Tool/db_backup.py`、`db_cleanup.py` 等腳本。
- **常見問題**：DB lock、資料不同步，請檢查連線與 refresh 流程。

## 其他
- **重要檔案**：`README.md`（功能說明、更新紀錄）、`PMS_mindmap.md`（架構心智圖）、`PMS.drawio`（流程圖）。
- **外部依賴**：主要為 Python 標準庫與 sqlite3，部分功能需 openpyxl、tkinter。

---
如需新增功能，請先檢查是否已有對應模組，並遵循現有架構與命名慣例。
