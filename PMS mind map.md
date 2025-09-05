# SOP資訊平台開發進度 (Production Information Platform)

---

## 使用者認證 (User Authentication)
- [x] 登入功能（Login）
- [x] 角色權限控管（Role-Based Access Control）
- [x] 密碼變更（Password Change）
  - 修正視窗初始化問題（Tkinter `Toplevel`）
  - 提升 UNC 路徑環境相容性（Improve UNC compatibility）

---

## 資料庫設計 (Database Design)
- [x] `issues` 表：料號、品名、DIP、組裝、測試、包裝 SOP 與 OQC 欄位
- [x] `users` 表：帳號、密碼、角色、權限欄位
- [x] `activity_logs` 表：操作紀錄（上傳、生成、套用、帳號修改等）
- [x] `fixtures` 表：治具料號、品名、規格、類群、單價USD、安庫量
- [x] `warehouse_stock` 表：各倉別存量
- [x] `transfer_logs` 表：倉庫調撥紀錄
- [x] `fixture_boms` 表：治具 BOM 結構
- [x] 欄位自動補齊機制（Auto Schema Update）

---

## SOP 管理功能 (SOP Management)
- [x] SOP 上傳（依據使用者專長限制 Upload by Specialty）
- [x] SOP 生成（拼圖式 PDF 合併，命名格式：料號_品名_時間戳）
- [x] SOP 批次套用（Batch Apply to `issues` records）
- [x] SOP 欄位右鍵啟用／停用（Bypass Field Control）
- [x] 雙擊 SOP 欄位直接開啟 PDF（Double-Click to Open）

---

## 治具管理功能 (Fixture Management)
- [x] 新增/刪除治具料號
- [x] 入庫（⏫）
- [x] 轉倉（🔁）
- [x] 匯出 Excel（📤）
- [x] 多倉別分頁顯示（治具室、上齊、睿均、捷暉、立榮、華勤、上貿、麥博、信利、GC、工程、不良品）
- [x] 欄位顯示：
  - 治具料號、品名、規格、類群
  - 單價USD、單價NTD、總價NTD
  - 在庫量
  - 安庫量（僅治具室）
- [x] Excel 匯出：
  - 顯示「缺少數量」「總價NTD」
  - 計算「預估採購總價NTD」
- [ ] UI 即時安庫不足警告（TreeView 紅色高亮）
- [ ] 與 BOM 結合的安庫檢查（需求 vs. 供給）

---

## GUI 介面 (GUI Interface)
- [x] 使用者介面設計（Tkinter Layout）
- [x] 分頁設計（Tab-Based Interface）
  - SOP資訊（Production Info）
  - SOP 生成與套用（SOP Generation & Application）
  - 帳號管理（User Management）
  - 操作紀錄（Activity Logs）
  - 治具管理（Fixture Management）
- [x] 顯示目前使用者帳號與角色資訊（User & Role Display）
- [x] 密碼變更按鈕與視窗（Password Change Button）

---

## 檔案處理 (File Handling)
- [x] 檔名規則：料號_品名_時間戳（Naming: code_name_timestamp.pdf）
- [x] 按專長分類儲存於子目錄（DIP, Assembly, Test, Packaging, OQC）
- [x] 雙擊檔名開啟 PDF（PDF File Opener）
- [x] `.ico` 圖示在 UNC 環境中無法顯示（Icon not supported on UNC path）

---

## 操作紀錄 (Operation Logs)
- [x] 紀錄使用者所有重要操作（Log Activities: Upload, Generate, Apply SOP, Change Permissions）
- [x] 支援 SOP 更新、權限修改等動作追蹤
- [x] 搜尋功能與升降序排序（Search & Sort）
- [x] 雙擊紀錄可開啟對應檔案（Open Related Files）

---

## 部署與執行機制 (Deployment & Execution)
- [x] 本地端虛擬環境測試（Local Venv for Validation）
- [x] 打包為 EXE 執行檔（PyInstaller Packaging）
- [x] 解決 UNC 路徑下打包相容性問題（UNC Path Compatibility）
- [x] 自動檢查版本、下載新版本至本機並啟動（Auto Versioning, Sync & Execute）

---

## 備份與復原機制 (Backup & Resilience)
- [x] 設置定時排程備份（Scheduled DB Backup）