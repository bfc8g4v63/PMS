# PMS專案

---

## 更新重點（截至 2025-10-02）

* **SOP 模組**
  * [x] SOP 上傳/生成流程完成：拼圖選取、排序、合併、生成 PDF
  * [x] SOP 套用功能完成：來源檔案選取、批量套用、全選/勾選
  * [x] SOP 套用清單 UX 優化（來源清單移除 X 軸，套用清單支援 X 軸與滑鼠滾輪）
  * [x] SOP 更新時間戳可單獨修改（不必更動內容）
  * [x] SOP 上傳按鈕加上 @safe_button_action 防止重複點擊

* **治具管理模組**
  * [x] 建立治具時自動生成四倉別資料
  * [x] 新增料號驗證（長度 8 或 12 碼、必須為數字）
  * [x] 刪除後再新增相同料號，不會遺留舊數據
  * [x] 入庫 / 調撥 / 消耗功能完成，數量不得為負
  * [x] 入庫倉別依 Notebook 分頁自動判斷，不再用下拉選單
  * [x] 儲位規則：1-1-1 ~ 9-4-50，不補零；僅虹堡倉有安庫與儲位
  * [x] 四倉分頁 TreeView 獨立顯示，並於下方顯示總數統計
  * [x] 倉別加總顯示不重複料號
  * [x] 匯出 Excel 時新增運算輸出 預估請購量/金額與總金額統計
  * [x] 四大操作按鈕（新增/入庫/調撥/消耗）尚待全面套用 @safe_button_action
  * [x] 治具 BOM 分頁 UI 規劃中（fixture_boms 已完成 介面層，資料表已建立）

* **治具紀錄模組**
  * [x] 分頁完成：查詢條件（料號 / 使用者）、TreeView 顯示、刪除紀錄
  * [x] 與 fixture_logger 整合（建立/入庫/調撥/消耗自動寫入）
  * [ ] Backlog：日期查詢、匯出 Excel

* **帳號與權限模組**
  * [x] 帳號管理
    * [x] 新增使用者
      * [x] 帳號
      * [x] 密碼
      * [x] 角色
        * [x] admin
        * [x] engineer
        * [x] leader
      * [x] 專業
        * [x] dip
        * [x] assembly
        * [x] test
        * [x] packaging
        * [x] oqc
      * [x] 新增
      * [x] 刪除
      * [x] 啟用
      * [x] 可見SOP紀錄
      * [x] 刪除SOP紀錄
      * [x] 刪除SOP
      * [x] 可見SOP資訊
      * [x] 帳號管理
    * [x] 修改權限
      * [x] 更新權限
        * [x] 帳號
        * [x] 密碼
        * [x] 角色
        * [x] 專長
      * [x] 刪除帳號
    * [x] 密碼修改（admin、engineer）

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
  * [x] 正式 DB 路徑整合
  * [x] migrate_db.py 重構（issues → sop_information、changelog → change_log）

* **改版歷程**
  * [x] 改版紀錄表完成：版本、日期、內容
  * [x] UI：狀態提示不再遮擋輸入框

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
  * [x] 登入介面
    * [x] 輸入框
      * [x] 帳號
      * [x] 密碼
        * [x] tab快捷切換
    * [x] 按鈕
      * [x] 登入（支援 Enter 快捷）
        * [x] 主介面
          * [x] SOP資訊
            * [x] 欄位：料號 / 品名 / DIP / 組裝 / 測試 / 包裝 / OQC
            * [x] 修正：SOP 更新時間戳、表格置中與欄寬調整
          * [x] SOP生成
            * [x] SOP生成：上傳拼圖、選用拼圖、調整序列、執行生成
              * [x] 上傳按鈕已套用 @safe_button_action
            * [x] SOP套用
              * [x] 功能：搜尋、來源選擇、勾選/全選、批量套用
              * [x] 修正：來源清單移除 X 軸；套用清單支援 X 軸與滑鼠滾輪；檢索更精準
          * [x] SOP紀錄：帳號 / 料號 / 動作 / 時間
          * [x] 治具管理
            * [x] 基本資料：料號、品名、規格、類群、單價（NTD/USD）、總價、倉別總價
            * [x] 數量：安全庫存、可用數量、倉別總數、分頁顯示合計
            * [x] 修正：料號/品名驗證、儲位規則（1-1-1 至 9-4-50）、刪除後重建不殘留、虹堡倉限制
            * [x] 操作功能：建立治具 / 刪除治具 / 修改資料 / 入庫（依分頁倉別） / 調撥（來源倉雙擊快捷預設、目標倉） / 消耗 / 匯出 EXCEL（倉內資料、料件外幣單價、料件本幣單價、料件本幣總價，倉別總價、倉別總數、預估請購量/金額/總額）
          * [x] 治具紀錄
            * [x] 查詢條件：料號 / 使用者
            * [x] TreeView 顯示：單號 / 治具料號 / 治具規格 / 動作 / 異動數量 / 來源倉 / 目的倉 / 治具操作人 / 時間
            * [x] 操作功能：查詢 / 刪除紀錄
            * [x] 後端整合：fixture_logger，自動寫入入庫 / 調撥 / 消耗 / 建立治具紀錄
            * [ ] Backlog（匯出 Excel）
          * [ ] 治具 BOM
            * [x] 後端支援：fixture_boms 資料表與函式（get_bom_by_part, add_bom_item,delete_bom_item）
            * [x] UI 表單：新增 / 刪除 / TreeView 顯示
            * [ ] Backlog（UI 與後端整合、BOM 匯出）
          * [ ] Backlog(治具申請)
          * [ ] Backlog(治具損耗)
          * [x] 帳號管理：新增 / 刪除 / 變更 / 密碼修改（Eng+）、啟用 / 停用
          * [x] 改版歷程：版本 / 日期 / 內容
            * [x] 自動產生版本
            * [x] 雙擊編輯快捷
            * [x] 自定義各內容

---

## 應用層

* [x] 版控：Git
* [x] 資料驗證
  * [x] 共用規則：時間戳 yyyy-MM-dd HH:mm:ss、數值不得 NaN/None
  * [x] SOP 模組：名稱 = 料號+品名、格式限制（.pdf/.xlsx）
  * [x] 帳號管理：帳號唯一、密碼 6–12 位、狀態啟用/停用
  * [x] 治具管理：料號驗證、數量不得為負、調撥來源倉≠目的倉且數量足夠、單價必須為正
  * [x] SOP紀錄：帳號存在且啟用、動作屬於允許集合、時間戳有效
  * [x] 檔案與路徑：UNC 可寫、檔名不得含非法字元
* [x] 登入：顯示使用者 UI、登出同步回寫、一機一開、Enter key 修正
* [x] 閒置登出：Idle Timeout
* [x] 版本管理：檢查 / 自動更新 / 自動部屬
* [x] 系統：初始化方式調整、統一使用 with get_conn、移除多餘 checkpoint

---

## 資料庫層（inspect .177, 2025-10-02）

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