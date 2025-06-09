@echo off
pyinstaller --onefile --noconsole main.py
move /Y dist\main.exe .
