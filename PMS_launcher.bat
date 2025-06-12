@echo off
chcp 65001 >nul
setlocal

REM === 路徑定義 ===
set "SOURCE_DIR=\\192.120.100.177\工程部\生產管理\生產資訊平台"
set "SOURCE_EXE=%SOURCE_DIR%\PMS.exe"
set "SOURCE_VER=%SOURCE_DIR%\version.txt"
set "DEST_EXE=%TEMP%\PMS_temp.exe"
set "DEST_VER=%TEMP%\PMS_version.txt"
set "LOCK=%TEMP%\PMS.lock"
set "LOG=%TEMP%\PMS_start.log"

echo [啟動器] 初始化中，請稍候...

REM === Zombie lock 防呆處理 ===
if exist "%LOCK%" (
    tasklist | findstr /I PMS_temp.exe >nul
    if errorlevel 1 (
        echo [警告] 偵測到 Zombie Lock，自動清除。
        del "%LOCK%" >nul 2>&1
    ) else (
        echo [提示] PMS_temp.exe 執行中，請勿重複開啟。
        timeout /t 2 >nul
        exit /b
    )
)

REM === 掛載 Z: 解決 UNC 路徑問題 ===
net use Z: /delete >nul 2>&1
net use Z: "%SOURCE_DIR%" /persistent:no >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 無法連線伺服器，請檢查網路。
    pause
    exit /b
)

REM === 讀取版本資訊（只擷取版本號前段）===
set "REMOTE_VER="
if exist "%SOURCE_VER%" (
    for /f "tokens=1" %%v in (%SOURCE_VER%) do set "REMOTE_VER=%%v"
)

set "LOCAL_VER="
if exist "%DEST_VER%" (
    for /f "tokens=1" %%v in (%DEST_VER%) do set "LOCAL_VER=%%v"
)

REM === 比對版本，決定是否鏡像 ===
if "%REMOTE_VER%"=="%LOCAL_VER%" (
    echo [啟動器] 版本一致（%REMOTE_VER%），直接執行。
) else (
    echo [啟動器] 發現新版本（%REMOTE_VER%），部屬中...
    copy /Y "%SOURCE_EXE%" "%DEST_EXE%" >nul
    if errorlevel 1 (
        echo [錯誤] 部屬失敗，請洽管理員。
        pause
        net use Z: /delete >nul 2>&1
        exit /b
    )
    echo %REMOTE_VER% > "%DEST_VER%"
)

REM === 建立 Lock 並啟動 ===
echo %time% > "%LOCK%"
echo [啟動器] 啟動 PMS_temp.exe...
start "" /B "%DEST_EXE%"

REM === 可選：記錄啟動時間
echo [%date% %time%] 啟動 PMS 版本 %REMOTE_VER% >> "%LOG%"

REM === 取消掛載 Z:
net use Z: /delete >nul 2>&1

exit /b
