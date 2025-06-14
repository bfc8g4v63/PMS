import sqlite3
import os
import subprocess
import sys
from tkinter import messagebox
from datetime import datetime

def open_file(filepath):
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.call(["open", filepath])
        else:
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("錯誤", f"無法開啟檔案: {e}")

ACTION_MAP = {
    "add_user": "新增使用者",
    "update_user": "修改使用者",
    "delete_user": "刪除使用者",
    "upload": "新增 SOP",
    "generate_sop": "生成 SOP",
    "apply_sop": "套用 SOP",
    "delete": "刪除紀錄",
    "login": "登入系統",
    "logout": "登出系統",
    "change_password": "變更密碼",
    "更新SOP": "更新 SOP",
}

def log_activity(db_name, user, action, filename, module=None):
    action_display = ACTION_MAP.get(action, action)
    with sqlite3.connect(db_name, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO activity_logs (username, action, filename, timestamp, module)
            VALUES (?, ?, ?, ?, ?)
        """, (user, action, filename, datetime.now().strftime("%Y%m%dT%H%M%S"), module))
        conn.commit()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")