# PMS專案

---

## 更新重點（截至 2025-09-19）

* **SOP 模組**

  * [x] SOP 上傳/生成流程完成：拼圖選取、排序、合併、生成 PDF
  * [x] SOP 套用功能完成：來源檔案選取、批量套用、全選/勾選
  * [x] SOP 套用清單 UX 優化（來源清單移除 X 軸，套用清單支援 X 軸與滑鼠滾輪）
  * [x] SOP 更新時間戳可單獨修改（不必更動內容）
  * [x] SOP 上傳按鈕加上 @safe\_button\_action 防止重複點擊

* **治具管理模組**

  * [x] 建立治具時自動生成四倉別資料
  * [x] 新增料號驗證（長度 8 或 12 碼、必須為數字）
  * [x] 刪除後再新增相同料號，不會遺留舊數據
  * [x] 入庫/調撥/消耗功能完成，數量不得為負
  * [x] 入庫倉別依 Notebook 分頁自動判斷，不再用下拉選單
  * [x] 儲位規則：1-1-1 \~ 9-4-50，不補零；僅虹堡倉有安庫與儲位
  * [x] 四倉分頁 TreeView 獨立顯示，並於下方顯示總數統計
  * [x] 倉別加總顯示不重複料號
  * [x] 匯出 Excel 時新增預估請購量/金額與總金額統計
  * [ ] 四大操作按鈕（新增/入庫/調撥/消耗）尚待全面套用 @safe\_button\_action
  * [ ] 治具 BOM 分頁 UI 規劃中（後端 DB 已完成）

* **帳號與權限模組**

  * [x] 帳號管理完成：新增/刪除/變更/停用/啟用
  * [x] 密碼規則：6–12 位，Eng+ 格式
  * [x] 欄位顯示與操作表格置中

* **登入與使用者管理**

  * [x] 登入顯示使用者名稱
  * [x] 登出同步回寫
  * [x] 一機一開機制
  * [x] Idle Timeout 閒置自動登出
  * [x] Enter Key 雙重觸發問題修正

* **資料庫與穩定性**

  * [x] SQLite WAL 模式啟用
  * [x] busy\_timeout 與 Zombie lock 檢查
  * [x] checkpoint 策略模組化，避免 DB lock
  * [x] schema\_helper / fixture\_helper 初始化統一
  * [x] 定義不同資料變數數量 usable\_qty
  * [x] 正式 DB 路徑整合

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
        * [x] 登入
          * [x] 快捷Enter登入

    * [x] 主介面

      * [x] SOP資訊

        * [x] 欄位

          * [x] 料號
          * [x] 品名
          * [x] DIP
          * [x] 組裝
          * [x] 測試
          * [x] 包裝
          * [x] OQC
        * [x] 修正

          * [x] SOP 更新時間戳
          * [x] 表格置中與欄寬調整（料號 60、品名 180、其他 120）
      * [x] SOP生成

        * [x] SOP生成

          * [x] 上傳拼圖
          * [x] 選用拼圖
          * [x] 調整序列
          * [x] 執行生成
          * [x] 上傳按鈕套用 @safe\_button\_action 防重複點擊
        * [x] SOP套用

          * [x] 功能

            * [x] 搜尋
            * [x] 選擇來源
            * [x] 勾選/全選
            * [x] 批量套用
          * [x] 修正

            * [x] 來源清單移除 X 軸滾軸；套用清單支援 X 軸滾軸與滑鼠滾輪
            * [x] 優化套用清單檢索邏輯；查詢更精準
      * [x] 治具管理

        * [x] 基本資料

          * [x] 治具料號
          * [x] 治具品名
          * [x] 治具規格
          * [x] 治具類群
          * [x] 治具單價

            * [x] 金額：正數、允許小數點 3 位

              * [x] 治具單價（NTD）
              * [x] 治具單價（USD）
              * [x] 料件總價（NTD）
              * [x] 倉別總價（NTD）
          * [x] 數量

            * [x] 安全庫存
            * [x] 在庫數量
            * [x] 倉別總數
        * [x] 修正

          * [x] 建立治具時料號/品名/規格/類群驗證更嚴格
          * [x] 入庫功能數量輸入框修正
          * [x] 儲位原則判定調整（1-1-1 至 9-4-50，不補零）
          * [x] 刪除治具後新增相同料號不再遺留舊資料
          * [x] 修改資料快捷讀取資料原則與驗證回寫
          * [x] 安全庫存與儲位僅對虹堡倉有效，其他倉別為空白/0
          * [x] 儲位可自動生成不重複值；亦可自定義
          * [x] 儲位顯示於「安庫量」左側
          * [x] 入庫功能：倉別不再用下拉選單；依 Notebook 分頁自動判斷
          * [x] 四倉分頁各自 TreeView 與下方總數統計；倉別加總不重複料號
        * [x] 操作功能

          * [x] 修改基本資料

            * [x] 治具料號
            * [x] 治具品名
            * [x] 治具規格
            * [x] 治具類群
            * [x] 治具單價
            * [x] 安全庫存
            * [x] 儲位
          * [x] 建立治具
          * [x] 刪除治具
          * [x] 調撥
          * [x] 入庫（依分頁倉別）
          * [x] 消耗登記
          * [x] 匯出 EXCEL

            * [x] 倉內資料
            * [x] 倉別總價
            * [x] 預估請購量
            * [x] 預估請購金額
            * [x] 預估請購總額
        * [x] 改版歷程

          * [x] 功能

            * [x] 版本
            * [x] 日期
            * [x] 內容
          * [x] 修正

            * [x] 僅修改時間戳亦可觸發「儲存修改」
        * [ ] Backlog（待辦）

          * [ ] 治具料號 Bypass
          * [ ] 停產
          * [ ] 停售
          * [ ] 替代
          * [ ] 可交割數
      * [ ] Backlog（待辦）

        * [ ] 治具紀錄
        * [ ] 治具申請
        * [ ] 治具損耗
        * [ ] db\_locks 登入追蹤 UI
      * [x] 帳號管理

        * [x] 操作

          * [x] 新增帳號
          * [x] 刪除帳號
          * [x] 變更帳號
          * [x] 變更密碼（僅 Eng+）
          * [x] 帳號啟用
          * [x] 帳號停用
      * [x] SOP紀錄

        * [x] 欄位

          * [x] 帳號
          * [x] 料號
          * [x] 動作
          * [x] 時間
      * [x] 開發日誌

        * [x] 欄位

          * [x] 版本
          * [x] 日期
          * [x] 內容
      * [x] 其他 UI 修正

        * [x] 狀態提示不再遮擋內容輸入框（雙擊讀取版本時）

---

## 應用層

* [x] 版控

  * [x] Git
* [x] 資料驗證

  * [x] 共用規則

    * [x] 時間戳格式統一（yyyy-MM-dd HH\:mm\:ss）
    * [x] 數值欄位不得為 NaN / None
  * [x] SOP 模組

    * [x] 名稱規則：SOP 名稱 = 料號 + 品名
    * [x] 檔案格式：僅允許 .pdf / .xlsx
  * [x] 帳號管理

    * [x] 帳號唯一（不可重複）
    * [x] 密碼長度 6\~12 位英數
    * [x] 帳號狀態：啟用 / 停用
  * [x] 治具管理

    * [x] 料號：長度 8 或 12 碼，且為整數數字
    * [x] 入庫：依 Notebook 分頁自動判斷目標倉
    * [x] 調撥：來源倉 ≠ 目的倉，來源倉數量須足夠
    * [x] 數量：任何操作後不得為負
    * [x] 單價：必須為正數
    * [x] 安全庫存：必須 ≥ 0（僅虹堡倉有效）
  * [x] SOP紀錄

    * [x] 帳號必須存在且為啟用狀態
    * [x] 動作必須屬於允許集合
    * [x] 時間為有效時間戳
  * [x] 檔案與路徑

    * [x] UNC 路徑合法且可寫
    * [x] 檔名不得含非法字元
* [x] 登入

  * [x] 顯示登入使用者 UI
  * [x] 登出同步回寫
  * [x] 登入介面鎖分配原則
  * [x] Enter key Return 雙重觸發修正
* [x] 一機一開
* [x] 閒置登出

  * [x] Idle timeout
* [x] 版本管理

  * [x] 檢查
  * [x] 自動更新
  * [x] 自動部屬
* [x] 系統

  * [x] 修改

    * [x] 初始化方式調整
    * [x] 統一使用 get\_conn 管理 SQLite 連線
    * [x] 移除多餘 checkpoint，優化策略以降低 DB lock

---

## 資料庫層

* [x] 主資料庫架構

  * [x] SQLite

    * [x] WAL 模式
    * [x] 定時備份
    * [x] 地端與雲端切換
    * [x] 資料庫標準化；遷移；重構地端表
    * [x] 數量欄位命名一致（usable\_qty）
    * [x] checkpoint 模組化；降低 DB lock 機率
    * [x] 雲/地端版本同步
    * [x] 正式 DB 路徑
* [x] 模組初始化工具

  * [x] schema\_helper（含 ensure\_changelog\_schema）
  * [x] fixture\_helper（fixtures / warehouse\_stock / transfer\_logs / consumption\_logs / fixture\_boms）
* [x] SQLite 強化策略

  * [x] journal\_mode 保持同步（WAL）
  * [x] timeout 設定與 lock 預防
  * [x] Zombie lock 檢查機制
  * [x] 多層備份
* [x] 表結構現況

  * [x] fixtures：part\_no, part\_name, part\_spec, part\_group, unit\_price\_ntd, unit\_price\_usd, safety\_stock, storage\_location
  * [x] warehouse\_stock：part\_no, warehouse, usable\_qty, safety\_stock
  * [x] SOP：product\_code, product\_name, dip\_sop, assembly\_sop, test\_sop, packaging\_sop, oqc\_checklist
  * [x] users：username, password, role, specialty, can\_view\_logs, can\_delete\_logs, can\_upload\_sop, can\_view\_SOP, can\_manage\_users
  * [x] activity\_logs：id, username, action, filename, timestamp, module
  * [x] transfer\_logs：id, part\_no, from\_wh, to\_wh, transfer\_qty, user, remark, created\_at
  * [x] consumption\_logs：id, part\_no, warehouse, consume\_qty, user, remark, created\_at
  * [x] fixture\_boms：bom\_id, parent\_part\_no, child\_part\_no, qty, remark, created\_at（暫未於 UI 使用）
  * [x] changelog：version, date, content

---

