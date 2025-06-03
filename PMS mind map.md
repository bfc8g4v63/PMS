# 生產資訊平台開發進度 (Production Information Platform)

## 使用者認證 (User Authentication)
- 登入功能 (Login)
- 角色權限控管 (Role-based Access Control)
- 密碼變更 (Password Change)

## 資料庫設計 (Database Design)
- Issues 資料表 (料號、SOP、OQC 等)
- Users 資料表 (帳號、密碼、權限欄位)
- Activity Logs (操作紀錄表)
- Schema 欄位自動更新 (Auto Schema Update)

## SOP 管理功能 (SOP Management)
- SOP 上傳 (Upload with specialty)
- SOP 生成 (拼圖式 PDF Merge，命名格式: 料號_品名_時間戳)
- SOP 批次套用 (Batch Apply，並更新生產資訊表)
- 雙擊 SOP 開啟檔案功能 (Double-click to Open)

## GUI 介面 (GUI Interface)
- Tkinter Layout
- Tabs 分頁管理
- 生產資訊 (Production Info)
- SOP 生成 (SOP Management)
- 帳號管理 (User Management)
- 操作紀錄 (Activity Logs)

## 檔案處理 (File Handling)
- 儲存檔案依據料號_品名_時間戳 (Save File Naming Rule)
- 專長分目錄儲存 (DIP, 組裝, 測試, 包裝, OQC)
- 雙擊可開啟 PDF 文件

## 操作紀錄 (Operation Logs)
- 紀錄使用者操作 (Log Activity: Upload, Update, Apply SOP)
- SOP 生成紀錄

## 部署與批次鏡像 (Batch Execution and Deployment)
- 本地端 Venv 測試環境
- EXE 打包 (pyinstaller)
- 上傳 EXE 至 177 NAS
- 使用 .bat 批次檔鏡像到使用者本地端
