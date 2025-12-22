# PMS專案

---

## 更新重點（截至 2025-12-16）

* **SOP 模組**
  * [x] SOP 資訊
    * [x] SOP bypass 欄位邏輯完成（dip / assembly / test / packaging / oqc）
    * [x] SOP 欄位統計邏輯調整為「唯一料號數量」而非筆數
    * [x] SOP 更新時間戳可單獨修改（不必更動內容）
    * [x] SOP 上傳流程統一使用 get_conn()，避免 WAL 重複設定
    * [x] SOP 檔案實體路徑與 DB 記錄一致性檢查
  * [x] SOP 生成
    * [x] SOP 上傳/生成流程完成：拼圖選取、排序、合併、生成 PDF
    * [x] SOP 生成套用功能完成：來源檔案選取、批量套用、全選/勾選
    * [x] SOP 套用清單 UX 優化（來源清單移除 X 軸，套用清單支援 X 軸與滑鼠滾輪）
    * [x] SOP 上傳按鈕加上 @safe_button_action 防止重複點擊
    * [x] SOP 套用搜尋原則修正
    * [x] SOP 套用流程異常時不中斷整批操作（單筆失敗不中斷）

* **治具管理模組**
  * [x] 建立治具時自動生成十二倉別資料（虹堡總倉、上齊、睿均、捷暉、立榮、華勤、上貿、麥博、信利、GC、虹堡工程、不良品）
  * [x] 新增治具時自動建立所有既有倉別資料（非寫死倉名，支援後續擴充）
  * [x] 建立治具驗證（長度 8 或 12 碼、必須為數字）
  * [x] 生成料號（依類別前綴 930–939，確保不重複）
  * [x] 刪除後再新增相同料號，不會遺留舊數據
  * [x] 入庫 / 調撥流程完成，來源倉與目的倉不可相同，數量不得為負
  * [x] 入庫倉別依 Notebook 分頁自動判斷，不再使用下拉選單
  * [x] 調撥來源倉可由 TreeView 雙擊自動帶入
  * [x] 儲位規則：1-1-1 ~ 9-4-50，不補零；僅虹堡倉有安全庫存與儲位
  * [x] 各倉分頁 TreeView 獨立顯示，倉別加總不重複料號，並於下方顯示總數統計
  * [x] 各倉別 safety_stock 與治具主檔 safety_stock 職責分離
  * [x] 倉別數量欄位統一使用 usable_qty（歷史欄位清理）
  * [x] 匯出 Excel 新增預估請購量 / 預估請購金額 / 預估請購總金額
  * [x] 匯出 Excel 運算統一由資料層處理，UI 不重複計算
  * [x] 三大操作按鈕（新增 / 入庫 / 調撥）全面套用 @safe_button_action
  * [x] 消耗功能已移除（不再包含消耗倉）

* **治具紀錄模組**
  * [x] 分頁完成：查詢條件（料號 / 使用者）、TreeView 顯示、刪除紀錄
  * [x] 與 fixture_logger 整合（建立 / 入庫 / 調撥 自動寫入）
  * [x] fixture_logs 移除 remark 欄位，統一以 action、user、timestamp 記錄
  * [x] Backlog：日期查詢、匯出 Excel
  * [x] fixture_logs 與 transfer_logs 分工完成（操作紀錄 / 調撥紀錄）
  * [x] 刪除紀錄僅刪除紀錄本身，不回滾庫存
  * [x] 刪除紀錄需權限 can_delete_fixture_logs
  * [x] 預設查詢筆數限制（避免一次載入全部紀錄）

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
    * [x] 權限管理
      * [x] 修改權限
        * [x] 新增SOP列
        * [x] 刪除SOP列
        * [x] 帳號啟用
        * [x] 可見SOP紀錄
        * [x] 刪除SOP紀錄
        * [x] 上傳SOP
        * [x] 可見SOP資訊
        * [x] 帳號管理
        * [x] 可見治具管理
        * [x] 可編輯治具
        * [x] 可調帳治具
        * [x] 可見治具紀錄
        * [x] 刪除治具紀錄
      * [x] 修改帳密
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
  * [x]  登出、Idle Timeout 皆確保狀態回寫
  * [x] 多視窗重複啟動防呆（同帳號）

* **治具 BOM**
   * [x] BOM 資料表初始化納入 ensure_schema 流程
   * [x] BOM 主子料不可相同檢查
   * [x] BOM 刪除僅影響 BOM，不影響治具本體
   * [ ] Backlog：BOM 與庫存連動試算

* **資料庫與穩定性**
  * [x] SQLite WAL 模式啟用
  * [x] busy_timeout 與 Zombie lock 檢查
  * [x] checkpoint 策略模組化，避免 DB lock
  * [x] schema_helper / fixture_helper 初始化統一
  * [x] 定義不同資料變數數量 usable_qty
  * [x] 正式 DB 路徑整合
  * [x] migrate_db.py 重構（issues → sop_information、changelog → change_log）
  * [x] 所有寫入流程統一使用 with get_conn()
  * [x] 移除各模組自行設定 PRAGMA 的行為
  * [x] fixture_helper / schema_helper 責任邊界釐清
  * [x] 正式 DB 與測試 DB 切換集中於啟動階段
  * [x] WAL checkpoint 僅保留單一策略入口

* **改版歷程**
  * [x] 改版紀錄表完成：版本、日期、內容
  * [x] UI：狀態提示不再遮擋輸入框
  * [x] 改版內容支援多行描述
  * [x] 版本編輯後即時刷新 TreeView
  * [x] 版本資料表不存在時自動補齊

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
            * [x] 欄位：料號 / 品名 / DIP / 組裝 / 測試 / 包裝 / OQC / SOP建立人 / 時間
            * [x] 修正：SOP 更新時間戳、表格置中與欄寬調整
          * [x] SOP生成
            * [x] SOP生成：上傳拼圖、選用拼圖、調整序列、執行生成
              * [x] 上傳按鈕已套用 @safe_button_action
            * [x] SOP套用
              * [x] 功能：搜尋、來源選擇、勾選/全選、批量套用綁定資料列
              * [x] 修正：來源清單移除 X 軸；套用清單支援 X 軸與滑鼠滾輪；檢索更精準
              * [x] 修正：搜尋品號邏輯問題
          * [x] SOP紀錄：SOP建立人 / 動作 / 檔案名稱 / 時間
          * [x] 治具管理
            * [x] 治具操作
              * [x] 基本資料：治具料號、治具品名、治具規格、治具類群、治具單價（NTD/USD）、安全庫存、儲位、入庫數量
                * [x] 建立治具：治具單價、安全庫存(非必須資料；其餘需驗證)
                * [x] 刪除治具：雙擊該筆資料列後讀取該筆資料數據並指向，點選刪除治具後刪除該筆資料列於資料庫
                * [x] 生成料號：以選擇治具類群為依據&資料庫不重複&最小號碼填入治具料號
                * [x] 生成儲位：
                  * [x] 儲位規則（1-1-1 至 9-4-50）
                    * [x] 例A: "1-1" → "1-1-1~50"系統分配未分配最小位置或自訂義
                    * [x] 例B: "9" → "9-1~4-1~50"系統分配未分配最小層數&位置或自訂義
                    * [x] 刪除後重建不殘留、虹堡倉限制
                * [x] 修改資料：雙擊該筆資料列後讀出該筆資料列資料於基本資料；自訂義完成後點擊修改後回寫該筆資料列資料
                * [x] 入庫：入庫數量內數字寫入資料庫（依分頁倉別決策入庫倉別）
                * [x] 調撥（快捷 → 雙擊資料列讀取資料+來源倉預設、人為決策目標倉）→ 入庫數量決策撥出量
                * [x] 匯出 EXCEL（治具總覽輸出、料件外幣單價、料件本幣單價、料件本幣總價，倉別總價、倉別總數、預估請購量/金額/總額）
                * [x] 操作按鈕全面套用 @safe_button_action 避免重複觸發
          * [x] 治具紀錄
            * [x] 查詢條件：料號 / 使用者
            * [x] TreeView 顯示：單號 / 治具料號 / 治具規格 / 動作 / 異動數量 / 來源倉 / 目的倉 / 治具操作人 / 調帳原因 / 時間
            * [x] 操作功能：查詢 / 刪除紀錄
            * [x] 後端整合：fixture_logger，自動寫入 入庫 / 調撥 / 建立治具紀錄
          * [ ] 治具 BOM
            * [x] 後端支援：fixture_boms 資料表與函式（get_bom_by_part, add_bom_item, delete_bom_item）
            * [x] UI 表單：新增 / 刪除 / TreeView 顯示
            * [ ] Backlog（UI 與後端整合、BOM 匯出）
          * [ ] Backlog(治具申請)
          * [ ] Backlog(治具損耗)
          * [x] 帳號管理：新增SOP列、刪除SOP列、帳號啟用、可見SOP紀錄、刪除SOP紀錄、上傳SOP、可見SOP資訊、可見治具管理、可編輯治具、可調帳治具、可見治具紀錄、刪除治具紀錄
          * [x] 改版歷程：版本 / 日期 / 內容
            * [x] 產生版本
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
* [x] 調帳功能
  * [x] 調帳類型：庫存修正（盤點差異 / 資料修正）
  * [x] 調帳倉別依來源紀錄自動判斷
  * [x] 調帳數量支援正負調整，不得造成庫存為負
  * [x] 調帳原因必填，作為稽核與追蹤依據
  * [x] 調帳權限控管（限授權角色）
  * [x] 調帳不回滾歷史資料，僅新增紀錄
  * [x] 調帳完成後同步更新庫存與紀錄
  * [x] 調帳行為寫入 fixture_logs，動作標示為 ADJUST
  * [x] 模組初始化順序固定化（DB → schema → UI）
  * [x] 移除未使用欄位與變數（避免誤判為死碼）
  * [x] 所有路徑存取集中處理，避免散落硬編碼
  * [x] 日誌與操作紀錄欄位命名全面前綴化（避免歧義）
---

## 資料庫層（inspect .177, 2025-12-16）

* [x] 主資料庫架構
  * [x] SQLite
    * [x] WAL 模式
    * [x] 定時備份
    * [x] 地端與雲端切換
    * [x] 資料庫標準化；遷移；重構地端表
    * [x] 數量欄位命名一致（usable_qty）
    * [x] checkpoint 模組化；降低 DB lock 機率
    * [x] 雲/地端版本同步
    * [x] 正式 DB 路徑

* [x] 模組初始化工具
  * [x] schema_helper（含 ensure_changelog_schema）
  * [x] fixture_helper（fixtures / warehouse_stock / transfer_logs / fixture_boms）

* [x] SQLite 強化策略
  * [x] journal_mode 保持同步（WAL）
  * [x] timeout 設定與 lock 預防
  * [x] Zombie lock 檢查機制
  * [x] 多層備份
  * [x] fixture_logs、transfer_logs、activity_logs 職責明確區分
  * [x] sqlite_sequence 明確保留，不做清理
  * [x] 表結構自動補欄位（add_col_if_missing）
  * [x] DB lock 問題來源已定位至重複 PRAGMA 與多連線寫入

* [x] 表結構現況
  * **activity_logs**
    * activity_log_id (INTEGER), activity_log_username (TEXT), activity_log_action (TEXT), activity_log_filename (TEXT), activity_log_timestamp (TEXT), activity_log_module (TEXT)
  * **change_log**
    * version (TEXT), date (TEXT), content (TEXT)
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

* **Upcoming Features**
* [x] 測試BOM (schema-fixture_boms)
  * [ ] bom_model_name
    * [ ] part_no (TEXT)
    * [ ] part_name (TEXT)
    * [ ] part_spec (TEXT)
    * [ ] part_group (TEXT)
    * [ ] environment_bom_qty
    * [ ] environment_bom_created_at (TEXT)
