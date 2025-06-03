@echo off
chcp 65001 >nul
setlocal

REM === 定義變數 ===
set "LOCK=%TEMP%\PMS.lock"
set "SOURCE=\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.exe"
set "DEST=%TEMP%\PMS_temp.exe"

REM === 掛載共享資料夾為 Z:（避免 UNC 路徑問題）===
net use Z: /delete >nul 2>&1
net use Z: "\\192.120.100.177\工程部\生產管理\生產資訊平台" /persistent:no >nul 2>&1
if errorlevel 1 (
    echo 無法連線 PMS 主機，請確認網路與共用權限。
    pause
    exit /b
)

REM === 防重複開啟：判斷 PMS.lock 是否存在 + 程序是否存活 ===
if exist "%LOCK%" (
    tasklist | findstr /I PMS_temp.exe >nul
    if not errorlevel 1 (
        echo PMS 已啟動，請勿重複開啟。
        timeout /t 2 >nul
        net use Z: /delete >nul 2>&1
        exit /b
    )
    del "%LOCK%" >nul 2>&1
)

REM === 複製 PMS.exe 到 TEMP 跑 ===
echo 正在複製 PMS 執行檔...
copy /Y "Z:\PMS.exe" "%DEST%" >nul
if errorlevel 1 (
    echo 複製失敗，請確認來源檔案與權限。
    pause
    net use Z: /delete >nul 2>&1
    exit /b
)

REM === 啟動 PMS 程式（背景模式）===
echo PMS 系統啟動中...
start "" /B "%DEST%"
exit /b
