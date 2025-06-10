@echo off
chcp 65001 >nul
echo 版本檔寫入中...
python write_version.py

echo 開始打包 PMS.exe ...
pyinstaller --noconfirm --onefile --windowed --hidden-import=fitz --hidden-import=PyMuPDF --icon=PMS.ico PMS.py

echo 打包完成，請檢查 dist\PMS.exe
pause
