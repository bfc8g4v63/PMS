#$ PMS.py
#% 登入>主介面>整體初始化

import tkinter as tk
import os
import hashlib
import sys
import socket
import shutil
import time
import json
import tkinter.font as tkFont
from pathlib import Path
from tkinter import ttk, messagebox
from datetime import datetime

from fixture_helper import ensure_stock_consistency
from db_helper import get_conn, set_db_path
from fixture_logger import ensure_fixture_log_schema
from idle_logout import attach_idle_logout

from config import (
    apply_db_path,
    ENABLE_AUTO_LOGOUT,
    IDLE_TIMEOUT,
    VERBOSE_SCHEMA_CHECK,
    USE_LOCAL_DB,
    DB_NAME,
    Z_DRIVE_DB,
)

from utils import log_activity
from account_management_tab import build_user_management_tab
from sop_build_tab import build_sop_upload_tab, build_sop_apply_section
from changelog_tab import build_changelog_tab
from sop_logs_tab import build_sop_logs_tab
from sop_info_tab import build_sop_info_tab
from schema_helper import (
    get_required_columns,
    auto_add_missing_columns,
    ensure_changelog_schema,
    ensure_fixture_schema,
    print_tables_info,
)
from fixture_tabs import build_fixture_tab
from fixture_bom_tab import build_fixture_bom_tab
from fixture_logs_tab import build_fixture_logs_tab

DIP_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\DIP_SOP"
ASSEMBLY_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\組裝SOP"
TEST_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\測試SOP"
PACKAGING_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\包裝SOP"
OQC_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\檢查表OQC"

LOG_TABLE = "activity_logs"
_instance_lock = None

def get_login_config_path():
    appdata_dir = os.environ.get("APPDATA")
    if not appdata_dir:
        home_dir = str(Path.home())
        base_dir = os.path.join(home_dir, ".pms")
    else:
        base_dir = os.path.join(appdata_dir, "PMS")
    if not os.path.isdir(base_dir):
        os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, "login_config.json")

def load_saved_credentials():
    config_path = get_login_config_path()
    if not os.path.exists(config_path):
        return None
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {"username": data.get("username", ""), "password": data.get("password", "")}
    except Exception:
        return None

def save_credentials(username, password, remember_password):
    config_path = get_login_config_path()
    data = {"username": username}
    if remember_password:
        data["password"] = password
    else:
        data["password"] = ""
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass

def is_another_instance_running():
    global _instance_lock
    try:
        _instance_lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _instance_lock.bind(("localhost", 65432))
        return False
    except OSError:
        return True

def init_db(role):
    if not os.path.exists(DB_NAME):
        messagebox.showerror("錯誤", f"找不到資料庫檔案：\n{DB_NAME}\n\n請確認您是從 PMS Launcher.bat 啟動，或 Z: 磁碟有正確掛載。")
        sys.exit()

    if role == "leader":
        if not os.access(DB_NAME, os.R_OK):
            raise IOError(f"無法讀取資料庫檔案：{DB_NAME}")
    else:
        if not os.access(DB_NAME, os.R_OK | os.W_OK):
            raise IOError(f"無法讀寫資料庫檔案：{DB_NAME}")

    try:
        with get_conn() as conn:
            pass
    except Exception as e:
        messagebox.showerror("資料庫錯誤", f"無法開啟資料庫：\n{DB_NAME}\n\n錯誤訊息：{e}")
        sys.exit()

def sync_back_to_server():
    if not USE_LOCAL_DB:
        print("[雲端模式] 所有操作直接寫入網路 DB，無需回寫")
        return

    if not os.path.exists(Z_DRIVE_DB):
        print("偵測不到 Z: 掛載，跳過資料庫回寫")
        return

    if DB_NAME == Z_DRIVE_DB:
        print("無需回寫資料庫，因為 DB 實體與操作一致")
        return

    try:
        shutil.copy(DB_NAME, Z_DRIVE_DB)
        print("已同步本機資料庫回網路磁碟")
    except Exception as e:
        print(f"資料回寫失敗: {e}")

def logout_and_exit(root):
    global _instance_lock
    try:
        root.update_idletasks()
        root.update()
    except Exception:
        pass
    finally:
        try:
            if _instance_lock is not None:
                _instance_lock.close()
        except Exception:
            pass
        if USE_LOCAL_DB:
            sync_back_to_server()
        try:
            root.destroy()
        except Exception:
            pass
        os._exit(0)

def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def ensure_admin_full_permissions(username):
    try:
        u = (username or "").strip()
    except Exception:
        return
    if not u:
        return

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT role FROM users WHERE username=?", (u,))
        row = cur.fetchone()
        if not row:
            return
        role_norm = (row[0] or "").strip().lower()
        if role_norm != "admin":
            return

        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        targets = [c for c in cols if c.startswith("can_")]
        if "active" in cols:
            targets.append("active")
        if not targets:
            return

        set_sql = ", ".join([f"{c}=1" for c in targets])
        cur.execute(f"UPDATE users SET {set_sql} WHERE username=?", (u,))
        conn.commit()

def login():
    result = {
        "user": None,
        "role": None,
        "can_add": 0,
        "can_delete": 0,
        "specialty": "",
        "can_view_logs": 0,
        "can_delete_logs": 0,
        "can_upload_sop": 0,
        "can_view_sop_info": 0,
        "can_manage_users": 0,
        "can_view_fixture": 0,
        "can_edit_fixture": 0,
        "can_adjust_fixture": 0,
        "can_view_fixture_logs": 0,
        "can_delete_fixture_logs": 0,
    }

    def try_login():
        u = entry_user.get().strip()
        p = entry_pass.get().strip()
        if not u or not p:
            messagebox.showerror("錯誤", "請輸入帳號與密碼")
            return

        hashed_pw = hash_password(p)
        try:
            with get_conn() as conn:
                c = conn.cursor()
                c.execute(
                    (
                        "SELECT "
                        "role, "
                        "COALESCE(can_add, 0), "
                        "COALESCE(can_delete, 0), "
                        "COALESCE(specialty, ''), "
                        "COALESCE(can_view_logs, 0), "
                        "COALESCE(can_delete_logs, 0), "
                        "COALESCE(can_upload_sop, 0), "
                        "COALESCE(can_view_sop_info, 0), "
                        "COALESCE(can_manage_users, 0), "
                        "COALESCE(can_view_fixture, 0), "
                        "COALESCE(can_edit_fixture, 0), "
                        "COALESCE(can_adjust_fixture, 0), "
                        "COALESCE(can_view_fixture_logs, 0), "
                        "COALESCE(can_delete_fixture_logs, 0) "
                        "FROM users "
                        "WHERE username=? AND password=? AND active=1"
                    ),
                    (u, hashed_pw),
                )

                r = c.fetchone()
                if r:
                    role_norm = (r[0] or "").strip().lower()
                    result.update(
                        {
                            "user": u,
                            "role": role_norm,
                            "can_add": int(r[1] or 0),
                            "can_delete": int(r[2] or 0),
                            "specialty": r[3] or "",
                            "can_view_logs": int(r[4] or 0),
                            "can_delete_logs": int(r[5] or 0),
                            "can_upload_sop": int(r[6] or 0),
                            "can_view_sop_info": int(r[7] or 0),
                            "can_manage_users": int(r[8] or 0),
                            "can_view_fixture": int(r[9] or 0),
                            "can_edit_fixture": int(r[10] or 0),
                            "can_adjust_fixture": int(r[11] or 0),
                            "can_view_fixture_logs": int(r[12] or 0),
                            "can_delete_fixture_logs": int(r[13] or 0),
                        }
                    )

                    if role_norm == "admin":
                        result.update(
                            {
                                "can_add": 1,
                                "can_delete": 1,
                                "can_view_logs": 1,
                                "can_delete_logs": 1,
                                "can_upload_sop": 1,
                                "can_view_sop_info": 1,
                                "can_manage_users": 1,
                                "can_view_fixture": 1,
                                "can_edit_fixture": 1,
                                "can_adjust_fixture": 1,
                                "can_view_fixture_logs": 1,
                                "can_delete_fixture_logs": 1,
                            }
                        )
                        ensure_admin_full_permissions(u)

                    print("目前使用的 DB 檔案路徑：", DB_NAME)
                    save_credentials(u, p, remember_var.get())
                    login_window.destroy()
                else:
                    messagebox.showerror("錯誤", "帳號或密碼錯誤或帳號已停用")
        except Exception as e:
            messagebox.showerror("資料庫錯誤", f"無法連線資料庫，請稍後再試。\n\n錯誤訊息：{e}")

    login_window = tk.Tk()
    login_window.title("登入系統")
    login_window.geometry("300x260")
    try:
        login_window.iconbitmap("PMS.ico")
    except Exception:
        pass

    tk.Label(login_window, text="使用者名稱：").pack(pady=(15, 5))
    entry_user = tk.Entry(login_window)
    entry_user.pack()

    tk.Label(login_window, text="密碼：").pack(pady=(10, 5))
    entry_pass = tk.Entry(login_window, show="*")
    entry_pass.pack()

    remember_var = tk.BooleanVar(value=False)
    tk.Checkbutton(login_window, text="記住密碼", variable=remember_var).pack(pady=(5, 0))

    saved = load_saved_credentials()
    if saved:
        if saved.get("username"):
            entry_user.insert(0, saved.get("username"))
        if saved.get("password"):
            entry_pass.insert(0, saved.get("password"))
            remember_var.set(True)

    tk.Button(login_window, text="登入", command=try_login).pack(pady=15)

    entry_user.focus_set()
    entry_user.bind("<Return>", lambda e: entry_pass.focus_set())
    entry_pass.bind("<Return>", lambda e: try_login() or "break")

    def on_close():
        login_window.destroy()

    login_window.protocol("WM_DELETE_WINDOW", on_close)
    login_window.mainloop()
    return result

def initialize_database():
    with get_conn() as conn:
        cursor = conn.cursor()
        cursor.execute(
            (
                f"CREATE TABLE IF NOT EXISTS {LOG_TABLE} ("
                "activity_log_id INTEGER PRIMARY KEY AUTOINCREMENT,"
                "activity_log_username TEXT,"
                "activity_log_action TEXT,"
                "activity_log_filename TEXT,"
                "activity_log_timestamp TEXT,"
                "activity_log_module TEXT"
                ")"
            )
        )
        cursor.execute(
            (
                "CREATE TABLE IF NOT EXISTS sop_information ("
                "product_code TEXT PRIMARY KEY,"
                "product_name TEXT,"
                "dip_sop TEXT,"
                "assembly_sop TEXT,"
                "test_sop TEXT,"
                "packaging_sop TEXT,"
                "oqc_checklist TEXT,"
                "created_by TEXT,"
                "created_at TEXT,"
                "dip_sop_bypass INTEGER DEFAULT 0,"
                "assembly_sop_bypass INTEGER DEFAULT 0,"
                "test_sop_bypass INTEGER DEFAULT 0,"
                "packaging_sop_bypass INTEGER DEFAULT 0,"
                "oqc_checklist_bypass INTEGER DEFAULT 0"
                ")"
            )
        )
        cursor.execute(
            (
                "CREATE TABLE IF NOT EXISTS users ("
                "username TEXT PRIMARY KEY,"
                "password TEXT,"
                "role TEXT DEFAULT 'user',"
                "specialty TEXT DEFAULT '',"
                "can_view_logs INTEGER DEFAULT 0,"
                "can_delete_logs INTEGER DEFAULT 0,"
                "can_upload_sop INTEGER DEFAULT 0,"
                "can_view_sop_info INTEGER DEFAULT 0,"
                "can_manage_users INTEGER DEFAULT 0,"
                "can_add INTEGER DEFAULT 1,"
                "can_delete INTEGER DEFAULT 0,"
                "can_view_fixture INTEGER DEFAULT 1,"
                "can_edit_fixture INTEGER DEFAULT 0,"
                "can_adjust_fixture INTEGER DEFAULT 0,"
                "can_view_fixture_logs INTEGER DEFAULT 1,"
                "can_delete_fixture_logs INTEGER DEFAULT 0,"
                "active INTEGER DEFAULT 1"
                ")"
            )
        )
        conn.commit()

    print("資料庫初始化完成，實際位置：", DB_NAME)
    if hasattr(os, "sync"):
        try:
            os.sync()
        except Exception:
            pass
    auto_add_missing_columns(DB_NAME, get_required_columns())

def create_main_interface(root, db_name, login_info):
    current_role = login_info["role"]
    can_view_fixture = int(login_info.get("can_view_fixture", 0) or 0)
    can_view_fixture_logs = int(login_info.get("can_view_fixture_logs", 0) or 0)
    can_upload_sop = int(login_info.get("can_upload_sop", 0) or 0)

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tabs = {
        "SOP資訊": tk.Frame(notebook),
        "SOP生成": tk.Frame(notebook) if (current_role == "admin" or (current_role == "engineer" and can_upload_sop == 1)) else None,
        "SOP紀錄": tk.Frame(notebook) if current_role in ("admin", "engineer", "leader") else None,
        "治具管理": tk.Frame(notebook) if can_view_fixture else None,
        "治具紀錄": tk.Frame(notebook) if can_view_fixture_logs else None,
        "測試BOM": tk.Frame(notebook) if can_view_fixture else None,
        "治具申請": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "治具損耗": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "異常平台": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "帳號管理": tk.Frame(notebook) if current_role == "admin" else None,
        "改版歷程": tk.Frame(notebook) if current_role in ("admin", "engineer", "leader") else None,
    }

    for name, frame in tabs.items():
        if frame:
            notebook.add(frame, text=name)

    if tabs.get("SOP生成"):
        sop_tab = tabs["SOP生成"]
        left_frame = tk.Frame(sop_tab)
        left_frame.pack(side="left", fill="both", expand=True)
        right_frame = tk.Frame(sop_tab)
        right_frame.pack(side="left", fill="both", padx=10, pady=10)
        build_sop_upload_tab(left_frame, login_info, db_name)
        build_sop_apply_section(right_frame, login_info, db_name)

    def refresh_fixture_logs():
        if tabs.get("治具紀錄"):
            build_fixture_logs_tab(tabs["治具紀錄"], refresh_only=True, current_user=login_info)

    if tabs.get("治具紀錄"):
        build_fixture_logs_tab(tabs["治具紀錄"], current_user=login_info)

    if tabs.get("治具管理"):
        build_fixture_tab(tabs["治具管理"], login_info, on_change=refresh_fixture_logs)

    if tabs.get("測試BOM"):
        build_fixture_bom_tab(tabs["測試BOM"], login_info)

    if current_role == "admin" and tabs.get("帳號管理"):
        build_user_management_tab(tabs["帳號管理"], db_name, login_info)

    if current_role in ("admin", "engineer", "leader"):
        if tabs.get("SOP紀錄"):
            build_sop_logs_tab(
                tabs["SOP紀錄"],
                db_name,
                current_role,
                base_paths=[DIP_SOP_PATH, ASSEMBLY_SOP_PATH, TEST_SOP_PATH, PACKAGING_SOP_PATH, OQC_PATH],
            )
        if tabs.get("改版歷程"):
            build_changelog_tab(tabs["改版歷程"], current_role, db_name)

    if tabs.get("SOP資訊"):
        build_sop_info_tab(tabs["SOP資訊"], db_name, login_info)

def open_password_change_window(parent, db_name, username):
    win = tk.Toplevel(parent)
    win.title("變更密碼")
    win.geometry("300x220")
    win.resizable(False, False)
    try:
        win.iconbitmap("PMS.ico")
    except Exception:
        pass

    tk.Label(win, text="舊密碼：").pack(pady=(10, 0))
    entry_old = tk.Entry(win, show="*")
    entry_old.pack()

    tk.Label(win, text="新密碼：").pack(pady=(10, 0))
    entry_new = tk.Entry(win, show="*")
    entry_new.pack()

    tk.Label(win, text="確認新密碼：").pack(pady=(10, 0))
    entry_confirm = tk.Entry(win, show="*")
    entry_confirm.pack()

    def confirm_change():
        old_pw = entry_old.get().strip()
        new_pw = entry_new.get().strip()
        confirm_pw = entry_confirm.get().strip()

        if not old_pw or not new_pw or not confirm_pw:
            messagebox.showwarning("警告", "請填寫所有欄位")
            return
        if new_pw != confirm_pw:
            messagebox.showerror("錯誤", "新密碼與確認密碼不一致")
            return

        old_hash = hashlib.sha256(old_pw.encode("utf-8")).hexdigest()
        new_hash = hashlib.sha256(new_pw.encode("utf-8")).hexdigest()

        with get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT password FROM users WHERE username=? AND password=?", (username, old_hash))
            if not cursor.fetchone():
                messagebox.showerror("錯誤", "舊密碼不正確")
                return
            cursor.execute("UPDATE users SET password=? WHERE username=?", (new_hash, username))
            conn.commit()

        messagebox.showinfo("成功", "密碼已變更")
        win.destroy()

    tk.Button(win, text="變更密碼", bg="lightgreen", command=confirm_change).pack(pady=15)

if __name__ == "__main__":
    apply_db_path()
    print(f"使用資料庫：{DB_NAME}")
    set_db_path(DB_NAME)

    if is_another_instance_running():
        messagebox.showerror("錯誤", "本程式已在執行中，請勿重複開啟。")
        sys.exit()

    try:
        initialize_database()
        auto_add_missing_columns(DB_NAME, get_required_columns(), verbose=VERBOSE_SCHEMA_CHECK)
    except Exception as e:
        messagebox.showerror("資料庫錯誤", f"資料庫結構初始化失敗，請用可寫入權限帳號執行一次更新。\n\n錯誤訊息：{e}")
        sys.exit()

    login_info = login()

    if login_info and login_info.get("user"):
        current_role = login_info["role"]
        init_db(current_role)

        if current_role in ("admin", "engineer"):
            ensure_changelog_schema(DB_NAME, verbose=VERBOSE_SCHEMA_CHECK)
            ensure_fixture_schema(DB_NAME, verbose=VERBOSE_SCHEMA_CHECK)
            ensure_fixture_log_schema()
            auto_add_missing_columns(DB_NAME, get_required_columns(), verbose=VERBOSE_SCHEMA_CHECK)
            ensure_stock_consistency()

            if VERBOSE_SCHEMA_CHECK:
                print_tables_info(DB_NAME)

        root = tk.Tk()
        root.title("生產管理平台")
        root.geometry("1600x900")

        def idle_logout_callback():
            logout_and_exit(root)

        attach_idle_logout(root=root, enabled=ENABLE_AUTO_LOGOUT, idle_timeout_seconds=IDLE_TIMEOUT, logout_callback=idle_logout_callback)

        try:
            root.iconbitmap("PMS.ico")
        except Exception:
            pass

        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=10, family="Microsoft Calibri")

        top_bar = tk.Frame(root)
        top_bar.pack(fill="x", side="top")

        logout_btn = tk.Button(top_bar, text="登出並關閉", command=lambda: logout_and_exit(root), bg="orange")
        logout_btn.pack(side="right", padx=10, pady=5)

        if current_role in ("admin", "engineer"):
            change_pw_btn = tk.Button(
                top_bar,
                text="變更密碼",
                bg="lightgreen",
                command=lambda: open_password_change_window(root, DB_NAME, login_info["user"]),
            )
            change_pw_btn.pack(side="right", padx=10, pady=(0, 0))

        user_info = f"使用者：{login_info['user']}（{login_info['role']}）"
        tk.Label(top_bar, text=user_info).pack(side="right", padx=10)

        main_frame = tk.Frame(root)
        main_frame.pack(fill="both", expand=True)

        create_main_interface(main_frame, DB_NAME, login_info)

        def on_close():
            logout_and_exit(root)

        root.protocol("WM_DELETE_WINDOW", on_close)
        root.mainloop()
    else:
        print("使用者未登入或登入失敗，系統結束。")