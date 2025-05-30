import sqlite3
from datetime import datetime
import os
import subprocess
import sys
from tkinter import messagebox

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

def log_activity(db_name, user, action, filename, module=None):
    with sqlite3.connect(db_name) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO activity_logs (username, action, filename, timestamp, module)
            VALUES (?, ?, ?, ?, ?)
        """, (user, action, filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), module))
        conn.commit()
