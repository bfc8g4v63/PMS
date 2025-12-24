#$ utils.py
#% activity_logs 共用工具、輔助函式

import os
import subprocess
import sys
from tkinter import messagebox
from datetime import datetime
import threading
import functools
from db_helper import get_conn


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
    "update_sop": "更新 SOP",
    "generate_sop": "生成 SOP",
    "apply_sop": "套用 SOP",
    "delete": "刪除紀錄",
    "login": "登入系統",
    "logout": "登出系統",
    "change_password": "變更密碼",
    "toggle_bypass": "啟用/停用 SOP"
}


def log_activity(user, action, filename, module=None):
    raw_action = (action or "").strip()
    with get_conn() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO activity_logs (
                activity_log_username,
                activity_log_action,
                activity_log_filename,
                activity_log_timestamp,
                activity_log_module
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                user,
                raw_action,
                filename,
                datetime.now().strftime("%Y%m%dT%H%M%S"),
                module,
            ),
        )
        conn.commit()


def safe_button_action(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        widget = kwargs.get("button")
        ui_root = kwargs.get("ui_root") or widget

        def _ui_call(cb):
            if ui_root and hasattr(ui_root, "after"):
                ui_root.after(0, cb)
                return True
            return False

        if widget:
            if not _ui_call(lambda: widget.config(state="disabled")):
                widget.config(state="disabled")

        def run():
            err_msg = None
            try:
                func(*args, **kwargs)
            except Exception as e:
                err_msg = str(e)
            finally:
                if err_msg:
                    if not _ui_call(lambda m=err_msg: messagebox.showerror("錯誤", m)):
                        messagebox.showerror("錯誤", err_msg)
                if widget:
                    if not _ui_call(lambda: widget.config(state="normal")):
                        widget.config(state="normal")

        threading.Thread(target=run, daemon=True).start()
    return wrapper