# PMS專案

---

## 更新重點（截至 2025-10-07）

* **SOP 模組**
  * [x] SOP 上傳/生成流程完成：拼圖選取、排序、合併、生成 PDF
  * [x] SOP 套用功能完成：來源檔案選取、批量套用、全選/勾選
  * [x] SOP 套用清單 UX 優化（來源清單移除 X 軸，套用清單支援 X 軸與滑鼠滾輪）
  * [x] SOP 更新時間戳可單獨修改（不必更動內容）
  * [x] SOP 上傳按鈕加上 @safe_button_action 防止重複點擊
  * [x] 統一套用 fixture_helper.get_conn() 管理 SQLite 連線（v1.7.7 fix2）
  * [x] 移除重複 PRAGMA 設定，統一管理 checkpoint
  * [x] SOP 拼圖生成模組驗證完成（PDF 合併、檔案路徑安全性）

* **治具管理模組**
  * [x] 建立治具時自動生成四倉別資料
  * [x] 新增料號驗證（長度 8 或 12 碼、必須為數字）
  * [x] 刪除後再新增相同料號，不會遺留舊數據
  * [x] 入庫 / 調撥 / 消耗功能完成，數量不得為負
  * [x] 入庫倉別依 Notebook 分頁自動判斷，不再用下拉選單
  * [x] 儲位規則：1-1-1 ~ 9-4-50，不補零；僅虹堡倉有安庫與儲位
  * [x] 四倉分頁 TreeView 獨立顯示，並於下方顯示總數統計
  * [x] 倉別加總顯示不重複料號
  * [x] 匯出 Excel 時新增預估請購量/金額與總金額統計
  * [x] 四大操作按鈕（新增 / 入庫 / 調撥 / 消耗）全面套用 @safe_button_action（utils.py）
  * [x] 新增防重複觸發裝飾器 safe_button_action
  * [x] 治具室與外倉（上齊 / 睿均 / 不良品）數量一致性檢查 ensure_stock_consistency()
  * [x] fixture_helper 完成 transfer_stock / consume_stock / get_overview_by_warehouse 三函式整合
  * [x] fixture_helper 介面同步 PMS 主系統 schema 初始化
  * [x] fixture_boms 已完成介面層與資料表，待應用層整合
  * [ ] Backlog：治具 BOM 表單綁定 fixture_helper 呼叫、BOM 匯出功能

* **治具紀錄模組**
  * [x] 分頁完成：查詢條件（料號 / 使用者）、TreeView 顯示、刪除紀錄
  * [x] 與 fixture_logger 整合（建立 / 入庫 / 調撥 / 消耗自動寫入）
  * [x] 單號格式更新：
    * [x] C：建立治具
    * [x] D：刪除治具
    * [x] U：修改資料
    * [x] I：入庫
    * [x] T：調撥
    * [x] X：消耗
    * [ ]（預留 A：申請、R：退回、M：保養）
    * [ ] Backlog：日期查詢、匯出 Excel、分頁顯示

* **帳號與權限模組**
  * [x] 帳號管理、密碼修改與權限設定完成
  * [x] 欄位完整對應 users 表 can_add / can_delete / active
  * [x] 新增使用者權限防呆（重複帳號禁止新增）
  * [ ] Backlog：治具管理權限模組化

* **登入與使用者管理**
  * [x] 登入顯示使用者名稱
  * [x] 登出同步回寫
  * [x] 一機一開機制
  * [x] Idle Timeout 閒置自動登出
  * [x] Enter Key 雙重觸發問題修正

* **資料庫與穩定性**
  * [x] SQLite WAL 模式啟用
  * [x] busy_timeout 與 Zombie lock 檢查
  * [x] checkpoint 策略模組化，避免 DB lock
  * [x] schema_helper / fixture_helper 初始化統一
  * [x] 定義不同資料變數數量 usable_qty
  * [x] 正式 DB 路徑整合至 config.apply_db_path()
  * [x] migrate_db.py 重構（issues → sop_information、changelog → change_log）
  * [x] migrate_db_local.py 測試模式支援
  * [x] 新增 db_backup / sync_back_to_server() 雲端回寫安全策略
  * [x] PMS.py 初始化流程整合 ensure_schemas()，防止 DB_PATH 為 None

* **改版歷程**
  * [x] 改版紀錄表完成：版本、日期、內容
  * [x] UI：狀態提示不再遮擋輸入框
  * [x] export_readme.py：自動匯出 changelog 至 README.md

---

## 設計

* [x] 開發流程圖
  * [x] GitHub
    * [x] Drawio
  * [x] MS To Do
* [x] 系統架構設計
  * [x] GitHub
    * [x] Drawio
  * [x] MS To Do

---

## 介面層

* [x] 圖形介面框架 Tkinter
  * [x] SOP資訊 / SOP生成 / SOP紀錄 / 治具管理 / 治具紀錄 / 帳號管理 / 改版歷程
  * [x] 新增安全防呆裝飾器 safe_button_action（防重複點擊）
  * [x] 治具管理
    * [x] 儲位驗證規則（1-1-1 至 9-4-50）
    * [x] 安全庫存僅限虹堡倉有效
    * [x] 各倉別分頁下方顯示治具總數
  * [x] SOP生成模組統一使用 get_conn()
  * [x] 改版歷程頁面：雙擊編輯、版本自動生成
  * [ ] Backlog：治具申請 / 治具損耗

---

## 應用層

* [x] 統一 DB 連線介面（get_conn）
* [x] 新增 db_locks 管理登入資訊
* [x] 加強資料驗證：料號重複、數量不得負、帳號唯一
* [x] safe_button_action 套用於所有 GUI 操作
* [x] 自動登出模組整合主迴圈監聽
* [x] migrate_db 重構支援本地與雲端切換
* [x] export_readme.py 匯出版本說明

---

## 資料庫層

* **activity_logs**
  * activity_log_id (INTEGER), activity_log_username (TEXT), activity_log_action (TEXT), activity_log_filename (TEXT), activity_log_timestamp (TEXT), activity_log_module (TEXT)
* **change_log**
  * version (TEXT), date (TEXT), content (TEXT)
* **consumption_logs**
  * consumption_log_id (TEXT), consumption_log_part_no (TEXT), consumption_log_warehouse (TEXT), consumption_log_qty (INTEGER), consumption_log_user (TEXT), consumption_log_timestamp (TEXT)
* **fixture_boms**
  * fixture_bom_id (TEXT), fixture_bom_parent_no (TEXT), fixture_bom_child_no (TEXT), fixture_bom_qty (INTEGER), fixture_bom_timestamp (TEXT)
* **fixture_logs**
  * fixture_log_id (TEXT), fixture_log_part_no (TEXT), fixture_log_action (TEXT), fixture_log_qty (INTEGER), fixture_log_from_wh (TEXT), fixture_log_to_wh (TEXT), fixture_log_user (TEXT), fixture_log_timestamp (TEXT)
* **fixtures**
  * part_no (TEXT), part_name (TEXT), part_spec (TEXT), part_group (TEXT), unit_price_ntd (REAL), unit_price_usd (REAL), safety_stock (INTEGER), storage_location (TEXT), created_at (TEXT)
* **sop_information**
  * product_code (TEXT), product_name (TEXT), dip_sop (TEXT), assembly_sop (TEXT), test_sop (TEXT), packaging_sop (TEXT), oqc_checklist (TEXT), created_by (TEXT), created_at (TEXT), dip_sop_bypass (INTEGER), assembly_sop_bypass (INTEGER), test_sop_bypass (INTEGER), packaging_sop_bypass (INTEGER), oqc_checklist_bypass (INTEGER)
* **sqlite_sequence**
  * name (), seq ()
* **transfer_logs**
  * transfer_log_id (TEXT), transfer_log_part_no (TEXT), transfer_log_from_wh (TEXT), transfer_log_to_wh (TEXT), transfer_log_qty (INTEGER), transfer_log_user (TEXT), transfer_log_timestamp (TEXT)
* **users**
  * username (TEXT), password (TEXT), role (TEXT), specialty (TEXT), can_view_logs (INTEGER), can_delete_logs (INTEGER), can_upload_sop (INTEGER), can_view_sop_info (INTEGER), can_manage_users (INTEGER), can_add (INTEGER), can_delete (INTEGER), active (INTEGER)
* **warehouse_stock**
  * part_no (TEXT), warehouse (TEXT), usable_qty (INTEGER), safety_stock (INTEGER)
* **db_locks**
  * username (TEXT), hostname (TEXT), login_time (TEXT)

---