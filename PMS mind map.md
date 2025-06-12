# 生產資訊平台開發進度 (Production Information Platform)

## 使用者認證 (User Authentication)
- 登入功能 (Login)
- 角色權限控管 (Role-based Access Control)
- 密碼變更 (Password Change)
  - 視窗初始化修正（Tkinter Toplevel）
  - 提升 UNC 環境兼容性 (Increase UNC compatibility )

## 資料庫設計 (Database Design)
- 料號、品名、DIP、組、測、包SOP、OQC (Issues schema)
- 帳號、密碼、權限欄位(Users schema)
- 操作紀錄表 (Activity Logs)
- 欄位自動更新 (Auto Schema Update)

## SOP 管理功能 (SOP Management)
- SOP 上傳 (Upload with specialty)
- SOP 生成 (拼圖式 PDF Merge，命名格式: 料號_品名_時間戳)
- SOP 批次套用 (Batch apply and update to Issues schema)
- SOP 欄位右鍵停用／啟用
- 雙擊 SOP 開啟檔案功能 (Double-click to Open)

## GUI 介面 (GUI Interface)
- 使用者介面設計 (Tkinter Layout)
- 分頁管理 (Tabs) 
- 生產資訊 (Production Info)
- SOP 生成 (SOP gernerate、apply)
- 帳號管理 (User Management)
- 操作紀錄 (Activity Logs)
- 顯示目前使用者帳號／角色資訊 (Account info、character info)
- 密碼變更按鈕 (Password change)

## 檔案處理 (File Handling)
- 儲存檔案依據料號_品名_時間戳 (Save file naming rule)
- 專長分目錄儲存 (DIP, 組裝, 測試, 包裝, OQC，specilty)
- 雙擊可開啟 PDF 文件 (Open file for link)
- UNC 結構下不支援 icon 顯示問題（Except`.ico`）

## 操作紀錄 (Operation Logs)
- 紀錄使用者操作 (Log Activity: Upload, Update, Apply SOP)
- SOP 新增、生成、套用、權限變更紀錄 (SOP addition、gernerate、apply、account limit change) 
- 搜尋與排序控制 (Serch、sort)
- 雙擊開啟相關檔案 (Open file from operation trigger)

## 部署機制 (Batch Execution and Deployment)
- 本地端 Venv 測試環境 (Vitual-FCT)
- EXE 打包 (pyinstaller)
- UNC 結構打包相容性調整（UNC compatibility adjust）
- 自動檢查版本、安裝/啟動 (Auto version check、setup local、excuteing)

## 備份機制 (Backup and Resilience)
- 交叉兩支備份程序已啟用（Backup thread）
