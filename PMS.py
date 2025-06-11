import tkinter as tk
import sqlite3
import os
import hashlib
import subprocess
import sys
import socket
import re
import shutil
import atexit
import time

from utils import log_activity
from schema_helper import auto_add_missing_columns, get_required_columns
from account_management_tab import build_user_management_tab
from sop_build_tab import build_sop_upload_tab, build_sop_apply_section
from tkinter import ttk, filedialog, messagebox
from datetime import datetime


lock_path = os.path.join(os.environ.get("TEMP"), "PMS.lock")
with open(lock_path, "w") as f:
    f.write(str(time.time()))


@atexit.register
def remove_lock():
    try:
        os.remove(lock_path)
    except:
        pass
    
SOP_FIELDS = {
    "dip": ("DIP SOP", "dip_sop", "dip_sop_bypass", r"DIP_SOP"),
    "assembly": ("組裝SOP", "assembly_sop", "assembly_sop_bypass", r"組裝SOP"),
    "test": ("測試SOP", "test_sop", "test_sop_bypass", r"測試SOP"),
    "packaging": ("包裝SOP", "packaging_sop", "packaging_sop_bypass", r"包裝SOP"),
    "oqc": ("檢查表OQC", "oqc_checklist", "oqc_checklist_bypass", r"檢查表OQC")
}

BASE_SHARE = r"\\192.120.100.177\工程部\生產管理\生產資訊平台"
ORIGINAL_DB = os.path.join(BASE_SHARE, "PMS.db")
DB_NAME     = ORIGINAL_DB

DIP_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\DIP_SOP"
ASSEMBLY_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\組裝SOP"
TEST_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\測試SOP"
PACKAGING_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\包裝SOP"
OQC_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\檢查表OQC"

LOG_TABLE = "activity_logs"

_instance_lock = None

def is_valid_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            header = f.read(5)
            return header == b'%PDF-'
    except:
        return False

def save_file_if_exist(file_path, target_folder, username, product_code, product_name):
    if not file_path:
        return "", ""

    if not is_valid_pdf(file_path):
        messagebox.showerror("錯誤", f"這個檔案不是合法的 PDF 檔案：{file_path}")
        return "", ""

    if not product_name:
        messagebox.showerror("錯誤", f"品名為空，無法正確生成檔名")
        return "", ""

    timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    safe_name = product_name.replace("/", "-").replace("\\", "-").strip()
    if not safe_name:
        messagebox.showerror("錯誤", f"品名不合法，無法正確生成檔名")
        return "", ""

    filename = f"{product_code}_{safe_name}_{timestamp}.pdf"
    target_path = os.path.join(target_folder, filename)

    try:
        shutil.copy(file_path, target_path)
        return filename, timestamp
    except Exception as e:
        print(f"檔案儲存失敗: {e}")
        return "", ""

def is_another_instance_running():
    global _instance_lock
    try:
        _instance_lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _instance_lock.bind(('localhost', 65432))
        return False
    except OSError:
        return True

def init_db():
    if os.path.exists(DB_NAME) and not os.access(DB_NAME, os.R_OK | os.W_OK):
        raise IOError(f"無法讀寫資料庫檔案：{DB_NAME}")
    with sqlite3.connect(DB_NAME, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL")

def sync_back_to_server():
    if DB_NAME == ORIGINAL_DB:
        print("無需回寫資料庫，因為 DB 實體與操作一致")
        return
    try:
        shutil.copy(DB_NAME, ORIGINAL_DB)
        print("已同步本機資料庫回網路磁碟")
    except Exception as e:
        print(f"資料回寫失敗: {e}")

def logout_and_exit(root):
    global _instance_lock
    try:
        root.update_idletasks()
        root.update()
    except:
        pass
    finally:
        try:
            _instance_lock.close()
        except:
            pass

        try:
            with sqlite3.connect(DB_NAME, timeout=10) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        except Exception as e:
            print(f"[登出清理WAL失敗] {e}")

        sync_back_to_server()
        root.destroy()
        sys.exit()

def hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

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

def build_log_view_tab(tab, db_name, role):
    tk.Label(tab, text="操作紀錄查詢").pack(anchor="w", padx=10, pady=(10, 0))

    search_frame = tk.Frame(tab)
    search_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(search_frame, text="查詢關鍵字:").pack(side="left")
    entry_query = tk.Entry(search_frame)
    entry_query.pack(side="left")
    sort_desc = tk.BooleanVar(value=True)

    def toggle_sort():
        sort_desc.set(not sort_desc.get())
        refresh_logs()

    tk.Button(search_frame, text="↕排序", command=toggle_sort).pack(side="left", padx=5)
    tk.Button(search_frame, text="查詢", command=lambda: refresh_logs()).pack(side="left")

    columns = ("使用者", "動作", "檔案名稱", "時間")
    tree = ttk.Treeview(tab, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column("使用者", width=80)
        tree.column("動作", width=60)
        tree.column("檔案名稱", width=300)
        tree.column("時間", width=80)

    tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def refresh_logs():
        keyword = entry_query.get().strip()
        for row in tree.get_children():
            tree.delete(row)


        ACTION_MAP = {
            "add_user": "新增使用者",
            "update_user": "修改使用者",
            "delete_user": "刪除使用者",
            "upload": "新增SOP",
            "generate_sop": "生成 SOP",
            "apply_sop": "套用 SOP",
            "delete": "刪除紀錄",
            "login": "登入系統",
            "logout": "登出系統",
            "change_password": "變更密碼",
        }

        restricted_roles = ("engineer", "leader")
        restricted_keywords = ("add_user", "update_user", "delete_user")

        with sqlite3.connect(db_name, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            
            cursor = conn.cursor()
            base_sql = "SELECT id, username, action, filename, timestamp FROM activity_logs"
            params = []

            if role in restricted_roles:
                if keyword:
                    base_sql += " WHERE (username LIKE ? OR action LIKE ? OR filename LIKE ?) AND action NOT IN ({})".format(
                        ",".join("?" for _ in restricted_keywords)
                    )
                    params = [f"%{keyword}%"] * 3 + list(restricted_keywords)
                else:
                    base_sql += " WHERE action NOT IN ({})".format(
                        ",".join("?" for _ in restricted_keywords)
                    )
                    params = list(restricted_keywords)
            else:
                if keyword:
                    base_sql += " WHERE username LIKE ? OR action LIKE ? OR filename LIKE ?"
                    params = [f"%{keyword}%"] * 3

            base_sql += f" ORDER BY timestamp {'DESC' if sort_desc.get() else 'ASC'}"
            cursor.execute(base_sql, params)

            for row in cursor.fetchall():
                action_display = ACTION_MAP.get(row[2], row[2])
                tree.insert("", "end", iid=row[0], values=(row[1], action_display, row[3], row[4]))

    def on_double_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index == 2:
            filename = tree.item(item)["values"][2]
            base_paths = [DIP_SOP_PATH, ASSEMBLY_SOP_PATH, TEST_SOP_PATH, PACKAGING_SOP_PATH, OQC_PATH]
            for base in base_paths:
                path = os.path.join(base, filename)
                if os.path.exists(path):
                    open_file(path)
                    break

    tree.bind("<Double-1>", on_double_click)

    button_frame = tk.Frame(tab)
    button_frame.pack(anchor="e", padx=10, pady=(0, 10))
    tk.Button(button_frame, text="重新整理", command=refresh_logs).pack(side="left", padx=5)

    if role == "admin":
        def delete_selected_log():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提醒", "請先選取一筆操作紀錄")
                return
            if messagebox.askyesno("確認", "確定要刪除所選操作紀錄？"):
                with sqlite3.connect(db_name, timeout=10) as conn:
                    conn.execute("PRAGMA journal_mode=WAL;")
                    cursor = conn.cursor()
                    for iid in selected:
                        cursor.execute("DELETE FROM activity_logs WHERE id=?", (iid,))
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                refresh_logs()

        def delete_all_logs():
            if messagebox.askyesno("確認", "確定要刪除所有操作紀錄？此操作無法復原。"):
                with sqlite3.connect(db_name, timeout=10) as conn:
                    conn.execute("PRAGMA journal_mode=WAL;")
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM activity_logs")
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                refresh_logs()

        tk.Button(button_frame, text="刪除所選", command=delete_selected_log).pack(side="left", padx=5)
        tk.Button(button_frame, text="刪除全部", command=delete_all_logs).pack(side="left", padx=5)

    refresh_logs()

def initialize_database():
    with sqlite3.connect(DB_NAME,timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")

        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS issues (
                product_code TEXT PRIMARY KEY,
                product_name TEXT,
                dip_sop TEXT,
                assembly_sop TEXT,
                test_sop TEXT,
                packaging_sop TEXT,
                oqc_checklist TEXT,
                created_by TEXT,
                created_at TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT,
                role TEXT DEFAULT 'user',
                can_add INTEGER DEFAULT 1,
                can_delete INTEGER DEFAULT 0,
                active INTEGER DEFAULT 1
            )
        """)

        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {LOG_TABLE} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                action TEXT,
                filename TEXT,
                timestamp TEXT
            )
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS dev_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                version TEXT,
                content TEXT,
                updated_at TEXT,
                created_by TEXT
            )
        """)

        conn.commit()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
    print("資料庫初始化完成，實際位置：", DB_NAME)

    if hasattr(os, "sync"):
        os.sync()

    auto_add_missing_columns(DB_NAME, get_required_columns())

def save_file(file_path, target_folder, username, product_code=None, product_name=None, log=True):
    if not os.path.exists(file_path):
        return ""

    timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    safe_name = product_name.replace("/", "-").replace("\\", "-")
    filename = f"{product_code}_{safe_name}_{timestamp}.pdf"
    target_path = os.path.join(target_folder, filename)

    try:
        shutil.copy(file_path, target_path)
        if log:
            log_activity(DB_NAME, username, "新增SOP", filename, module="生產資訊")
        return filename
    except Exception as e:
        messagebox.showerror("錯誤", f"檔案儲存失敗: {e}")
        return ""

def update_sop_field(cursor, product_code, field_name, display_name):
    cursor.execute(f"UPDATE issues SET {field_name}=?, created_at=? WHERE product_code=?",
                   (display_name, datetime.now().strftime("%Y%m%dT%H%M%S"), product_code))

def handle_sop_update(product_code, product_name, sop_path, field_name, entry_widget, current_user):
    path = entry_widget.get().strip()
    if not path:
        return None

    with sqlite3.connect(DB_NAME, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        cursor = conn.cursor()
        cursor.execute("SELECT product_name FROM issues WHERE product_code=?", (product_code,))
        result = cursor.fetchone()
        if result:
            product_name = result[0]
        else:
            messagebox.showerror("錯誤", f"找不到料號 {product_code}，無法補上品名。")
            return None

    display_name = save_file(path, sop_path, current_user, product_code, product_name, log=False)

    if not display_name:
        return None

    with sqlite3.connect(DB_NAME, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        cursor = conn.cursor()
        update_sop_field(cursor, product_code, field_name, display_name)
        if cursor.rowcount == 0:
            messagebox.showerror("錯誤", f"找不到料號 {product_code}，無法更新！")
            return
        conn.commit()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")

    log_activity(DB_NAME, current_user, "更新SOP", display_name, module="生產資訊")

    return display_name


def create_sop_update_button(frame, row, label, sop_path, field_name, product_code_entry, product_name_entry, entry_widget, current_user, user_specialty, role, allowed_specialty):
    def update_action():
        if role != "admin" and user_specialty != allowed_specialty:
            messagebox.showerror("權限限制", f"您無法上傳 {label}，僅限 {allowed_specialty} 工程師")
            return
        product_code = product_code_entry.get().strip()
        product_name = product_name_entry.get().strip()
        if not product_code:
            messagebox.showwarning("警告", "請先輸入料號")
            return
        updated_filename = handle_sop_update(product_code, product_name, sop_path, field_name, entry_widget, current_user)

        if updated_filename:
            messagebox.showinfo("成功", f"已更新 {label} 檔案")

    btn = tk.Button(frame, text="更新", command=update_action)
    btn.grid(row=row, column=3, padx=5)
    return btn

def create_upload_field_with_update(row, label, folder, field_name, form, product_code_entry, product_name_entry, current_user, user_specialty, role, allowed_specialty):
    tk.Label(form, text=label).grid(row=row, column=0, sticky="e")
    entry = tk.Entry(form, width=50)
    entry.grid(row=row, column=1)

    def browse():
        path = filedialog.askopenfilename()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    tk.Button(form, text="選擇檔案", command=browse).grid(row=row, column=2)

    create_sop_update_button(
        form, row, label, folder, field_name, product_code_entry,
        product_name_entry, entry, current_user, user_specialty, role, allowed_specialty
    )

    return entry

def build_password_change_tab(tab, db_name, current_user):
    tk.Label(tab, text="變更密碼",).pack(pady=(10, 5))

    form = tk.Frame(tab)
    form.pack(pady=10)

    tk.Label(form, text="舊密碼：").grid(row=0, column=0, sticky="e", pady=5)
    old_pass_entry = tk.Entry(form, show="*", width=30)
    old_pass_entry.grid(row=0, column=1, padx=10)

    tk.Label(form, text="新密碼：").grid(row=1, column=0, sticky="e", pady=5)
    new_pass_entry = tk.Entry(form, show="*", width=30)
    new_pass_entry.grid(row=1, column=1, padx=10)

    tk.Label(form, text="確認新密碼：").grid(row=2, column=0, sticky="e", pady=5)
    confirm_pass_entry = tk.Entry(form, show="*", width=30)
    confirm_pass_entry.grid(row=2, column=1, padx=10)

    def change_password():
        old_pw = old_pass_entry.get().strip()
        new_pw = new_pass_entry.get().strip()
        confirm_pw = confirm_pass_entry.get().strip()

        if not old_pw or not new_pw or not confirm_pw:
            messagebox.showwarning("警告", "請填寫所有欄位")
            return

        if new_pw != confirm_pw:
            messagebox.showerror("錯誤", "新密碼與確認密碼不一致")
            return

        old_hash = hashlib.sha256(old_pw.encode()).hexdigest()
        new_hash = hashlib.sha256(new_pw.encode()).hexdigest()

        with sqlite3.connect(db_name, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            cursor = conn.cursor()
            cursor.execute("SELECT password FROM users WHERE username=? AND password=?", (current_user, old_hash))
            if not cursor.fetchone():
                messagebox.showerror("錯誤", "舊密碼不正確")
                return

            cursor.execute("UPDATE users SET password=? WHERE username=?", (new_hash, current_user))
            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")

        messagebox.showinfo("成功", "密碼已變更")
        old_pass_entry.delete(0, tk.END)
        new_pass_entry.delete(0, tk.END)
        confirm_pass_entry.delete(0, tk.END)

    tk.Button(tab, text="變更密碼", command=change_password, bg="lightgreen").pack(pady=10)

def create_main_interface(root, db_name, login_info):
    current_user = login_info['user']
    current_role = login_info['role']
    can_add = login_info['can_add']
    can_delete = login_info['can_delete']

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tabs = {
        "生產資訊": tk.Frame(notebook),
        "SOP生成": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "治具管理": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "測試BOM": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "異常平台": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "其他工具": tk.Frame(notebook) if current_role in ("admin", "engineer") else None,
        "帳號管理": tk.Frame(notebook) if current_role == "admin" else None,
        "操作紀錄": tk.Frame(notebook) if current_role in ("admin", "engineer", "leader") else None

    }

    for name, frame in tabs.items():
        if frame:
            notebook.add(frame, text=name)

    if current_role in ("admin", "engineer"):
        sop_tab = tabs["SOP生成"]

        left_frame = tk.Frame(sop_tab)
        left_frame.pack(side="left", fill="both", expand=True)

        right_frame = tk.Frame(sop_tab)
        right_frame.pack(side="left", fill="both", padx=10, pady=10)

        build_sop_upload_tab(left_frame, login_info, db_name)
        build_sop_apply_section(right_frame, login_info,db_name)

    if current_role in ("admin", "engineer", "leader"):
        build_log_view_tab(tabs["操作紀錄"], db_name, current_role)

    if current_role == "admin":
        build_user_management_tab(tabs["帳號管理"], db_name, current_user)

    frame = tabs["生產資訊"]
    if current_role != "leader":
        form = tk.LabelFrame(frame, text="新增紀錄")
        form.pack(fill="x", padx=10, pady=5)

        tk.Label(form, text="料號:").grid(row=0, column=0, sticky="e")
        entry_code = tk.Entry(form, width=50)
        entry_code.grid(row=0, column=1)

        tk.Label(form, text="品名:").grid(row=1, column=0, sticky="e")
        entry_name = tk.Entry(form, width=50)
        entry_name.grid(row=1, column=1)

        entry_dip = create_upload_field_with_update(2, "DIP SOP", DIP_SOP_PATH, "dip_sop", form, entry_code, entry_name, current_user, login_info['specialty'], current_role, "dip")
        entry_assembly = create_upload_field_with_update(3, "組裝SOP", ASSEMBLY_SOP_PATH, "assembly_sop", form, entry_code, entry_name, current_user, login_info['specialty'], current_role, "assembly")
        entry_test = create_upload_field_with_update(4, "測試SOP", TEST_SOP_PATH, "test_sop", form, entry_code, entry_name, current_user, login_info['specialty'], current_role, "test")
        entry_packaging = create_upload_field_with_update(5, "包裝SOP", PACKAGING_SOP_PATH, "packaging_sop", form, entry_code, entry_name, current_user, login_info['specialty'], current_role, "packaging")
        entry_oqc = create_upload_field_with_update(6, "檢查表OQC", OQC_PATH, "oqc_checklist", form, entry_code, entry_name, current_user, login_info['specialty'], current_role, "oqc")

        def save_data():
            code = entry_code.get().strip()
            name = entry_name.get().strip()

            if not name:
                messagebox.showerror("錯誤", "品名不能為空")
                return

            if len(code) not in (8, 12) or not code.isdigit():
                messagebox.showerror("錯誤", "必須為 8/12 碼數字")
                return

            with sqlite3.connect(db_name, timeout=10) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                
                cursor = conn.cursor()
                cursor.execute("SELECT product_code FROM issues WHERE product_code=?", (code,))
                if cursor.fetchone():
                    messagebox.showerror("錯誤", "料號已存在，請重新確認過。")
                    return

                d_file, d_time = save_file_if_exist(entry_dip.get().strip(), DIP_SOP_PATH, current_user, code, name)
                a_file, a_time = save_file_if_exist(entry_assembly.get().strip(), ASSEMBLY_SOP_PATH, current_user, code, name)
                t_file, t_time = save_file_if_exist(entry_test.get().strip(), TEST_SOP_PATH, current_user, code, name)
                p_file, p_time = save_file_if_exist(entry_packaging.get().strip(), PACKAGING_SOP_PATH, current_user, code, name)
                o_file, o_time = save_file_if_exist(entry_oqc.get().strip(), OQC_PATH, current_user, code, name)

                if d_file:
                    log_activity(db_name, current_user, "新增SOP", d_file, module="生產資訊")
                if a_file:
                    log_activity(db_name, current_user, "新增SOP", a_file, module="生產資訊")
                if t_file:
                    log_activity(db_name, current_user, "新增SOP", t_file, module="生產資訊")
                if p_file:
                    log_activity(db_name, current_user, "新增SOP", p_file, module="生產資訊")
                if o_file:
                    log_activity(db_name, current_user, "新增SOP", o_file, module="生產資訊")

                cursor.execute("""
                    INSERT INTO issues (product_code, product_name, dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist, created_by, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    code, name,
                    d_file, a_file, t_file, p_file, o_file,
                    current_user,
                    datetime.now().strftime("%Y%m%dT%H%M%S")
                ))
                conn.commit()
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            messagebox.showinfo("成功", "已新增紀錄")
            for e in [entry_code, entry_name, entry_dip, entry_assembly, entry_test, entry_packaging, entry_oqc]:
                e.delete(0, tk.END)
            query_data()

    if current_role != "leader":
        tk.Button(form, text="新增紀錄", command=save_data, bg="lightblue", state="normal" if can_add else "disabled").grid(row=7, column=1, pady=10)

    query_frame = tk.Frame(frame)
    query_frame.pack(fill="x", padx=10, pady=5)
    tk.Label(query_frame, text="查詢關鍵字: ").pack(side="left")
    entry_query = tk.Entry(query_frame)
    entry_query.pack(side="left")
    sort_desc = tk.BooleanVar(value=True)

    def toggle_sort():
        sort_desc.set(not sort_desc.get())
        query_data()

    tk.Button(query_frame, text="↕排序", command=toggle_sort).pack(side="left", padx=5)
    tk.Button(query_frame, text="查詢", command=lambda: query_data()).pack(side="left")

    columns = ("料號", "品名", "DIP SOP", "組裝SOP", "測試SOP", "包裝SOP", "檢查表OQC", "使用者", "建立時間")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    tree.pack(fill="both", expand=True, padx=10, pady=5)
    
    def on_right_click(event):
        if current_role not in ("admin", "engineer"):
            return

        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index not in range(2, 7):
            return

        product_code = tree.item(item)['values'][0]
        sop_key = list(SOP_FIELDS.keys())[col_index - 2]
        field_name = SOP_FIELDS[sop_key][1]
        bypass_field = SOP_FIELDS[sop_key][2]

        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(
            label="啟用/停用",
            command=lambda: toggle_bypass(product_code, field_name, bypass_field)
        )
        menu.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", on_right_click)

    tree.tag_configure("bypass", foreground="red")

    if current_role == "admin":

        def delete_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("提醒", "請先選取要刪除的資料")
                return
            if messagebox.askyesno("確認", "確定要刪除選取的資料？此操作無法復原。"):
                deleted_items = [] 
                with sqlite3.connect(db_name, timeout=10) as conn:
                    conn.execute("PRAGMA journal_mode=WAL;")
                    cursor = conn.cursor()
                    for item in selected_items:
                        product_code = str(tree.item(item)['values'][0]).zfill(8)
                        cursor.execute("DELETE FROM issues WHERE product_code=?", (product_code,))
                        deleted_items.append(product_code)
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                for code in deleted_items:
                    log_activity(DB_NAME, current_user, "delete", code, module="生產資訊")

                query_data()

        delete_frame = tk.Frame(frame)
        delete_frame.pack(fill="x", padx=10, pady=(0, 5), anchor="e")

        tk.Button(delete_frame, text="刪除選取資料", command=delete_selected,
                bg="lightcoral", fg="white").pack(side="right")

    def query_data():
        root.focus_force() 
        raw_input = entry_query.get().strip()
        for row in tree.get_children():
            tree.delete(row)

        with sqlite3.connect(db_name, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            #conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            cursor = conn.cursor()

            base_query = """
                SELECT product_code, product_name, dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist, created_by, created_at
                FROM issues
            """

            conditions = []
            params = []

            if raw_input:
                if "&" in raw_input:
                    terms = [t.strip() for t in raw_input.split("&")]
                    for term in terms:
                        conditions.append("((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    condition_sql = " AND ".join(conditions)
                elif "/" in raw_input:
                    terms = [t.strip() for t in raw_input.split("/")]
                    sub_conditions = []
                    for term in terms:
                        sub_conditions.append("((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    condition_sql = " OR ".join(sub_conditions)
                else:
                    condition_sql = "((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)"
                    params = [f"%{raw_input}%", f"%{raw_input}%"]

                final_query = f"{base_query} WHERE {condition_sql} ORDER BY created_at {'DESC' if sort_desc.get() else 'ASC'}"
            else:
                final_query = f"{base_query} ORDER BY created_at {'DESC' if sort_desc.get() else 'ASC'}"

            cursor.execute(final_query, params)
            for row in cursor.fetchall():
                row_display = list(row)
                product_code = str(row_display[0]).zfill(8)
                row_display[0] = product_code

                bypass_fields = [
                    "dip_sop_bypass",
                    "assembly_sop_bypass",
                    "test_sop_bypass",
                    "packaging_sop_bypass",
                    "oqc_checklist_bypass"
                ]

                product_name = row_display[1]
                timestamp = row_display[8]

                for i in range(2, 7):
                    sop_file = row_display[i]
                    if sop_file:

                        display_name = f"{product_code}_{product_name}_{timestamp}"

                        cursor.execute(f"SELECT {bypass_fields[i - 2]} FROM issues WHERE product_code=?", (product_code,))
                        bypass = cursor.fetchone()
                        if bypass and bypass[0]:
                            row_display[i] = f"（已停用） {display_name}"
                        else:
                            row_display[i] = display_name
                    else:
                        row_display[i] = ""

                tree.insert('', tk.END, values=row_display, tags=("bypass",) if "（已停用）" in str(row_display) else "")

    def on_double_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index in range(2, 7):  
            values = tree.item(item)['values']
            product_code = values[0]
            if "（已停用）" in str(values[col_index]):
                return

            field_map = {
                2: "dip_sop",
                3: "assembly_sop",
                4: "test_sop",
                5: "packaging_sop",
                6: "oqc_checklist"
            }
            target_field = field_map.get(col_index)
            if not target_field:
                return
            
            base_paths = [DIP_SOP_PATH, ASSEMBLY_SOP_PATH, TEST_SOP_PATH, PACKAGING_SOP_PATH, OQC_PATH]
            with sqlite3.connect(DB_NAME,timeout=10) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                cursor = conn.cursor()
                cursor.execute(f"SELECT {target_field} FROM issues WHERE product_code=?", (product_code,))
                result = cursor.fetchone()
                if result and result[0]:
                    filename = result[0]
                    sop_folder = base_paths[col_index - 2]
                    full_path = os.path.join(sop_folder, filename)
                    if os.path.exists(full_path):
                        open_file(full_path)
                    else:
                        messagebox.showerror("錯誤", f"找不到檔案：{full_path}")

    def on_copy(event):
        focus = tree.focus()
        if not focus:
            return
        col = tree.identify_column(event.x)
        col_index = int(col[1:]) - 1
        value = tree.item(focus)['values'][col_index]
        root.clipboard_clear()
        root.clipboard_append(str(value))
        root.update()

    tree.bind("<Double-1>", on_double_click)
    tree.bind("<Control-c>", on_copy)
    
    def toggle_bypass(product_code, field_name, bypass_field):
        with sqlite3.connect(DB_NAME, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            cursor = conn.cursor()

            cursor.execute(
                f"SELECT {bypass_field} FROM issues WHERE product_code=?",
                (product_code,)
            )
            current = cursor.fetchone()

            new_value = 0 if current and current[0] else 1
            cursor.execute(
                f"UPDATE issues SET {bypass_field}=? WHERE product_code=?",
                (new_value, product_code)
            )
            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        query_data()

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
        "can_view_issues": 0,
        "can_manage_users": 0
    }

    def try_login():
        u = entry_user.get().strip()
        p = entry_pass.get().strip()
        if not u or not p:
            messagebox.showerror("錯誤", "請輸入帳號與密碼")
            return

        hashed_pw = hash_password(p)

        try:
            with sqlite3.connect(DB_NAME, timeout=10) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                c = conn.cursor()
                c.execute("""
                    SELECT role, can_add, can_delete, specialty,
                           can_view_logs, can_delete_logs, can_upload_sop,
                           can_view_issues, can_manage_users
                    FROM users
                    WHERE username=? AND password=? AND active=1
                """, (u, hashed_pw))
                r = c.fetchone()
                if r:
                    result.update({
                        "user": u,
                        "role": r[0],
                        "can_add": r[1],
                        "can_delete": r[2],
                        "specialty": r[3],
                        "can_view_logs": r[4],
                        "can_delete_logs": r[5],
                        "can_upload_sop": r[6],
                        "can_view_issues": r[7],
                        "can_manage_users": r[8]
                    })
                    print("目前使用的 DB 檔案路徑：", DB_NAME)
                    login_window.destroy()
                else:
                    messagebox.showerror("錯誤", "帳號或密碼錯誤或帳號已停用")
        except sqlite3.OperationalError as e:
            messagebox.showerror("資料庫錯誤", f"無法連線資料庫，請稍後再試。\n\n錯誤訊息：{e}")

    login_window = tk.Tk()
    login_window.title("登入系統")
    login_window.geometry("300x180")
    try:
        login_window.iconbitmap("PMS.ico")
    except:
        pass

    tk.Label(login_window, text="使用者名稱：").pack(pady=(15, 5))
    entry_user = tk.Entry(login_window)
    entry_user.pack()

    tk.Label(login_window, text="密碼：").pack(pady=(10, 5))
    entry_pass = tk.Entry(login_window, show="*")
    entry_pass.pack()

    tk.Button(login_window, text="登入", command=try_login).pack(pady=15)

    def on_close():
        login_window.destroy()

    login_window.protocol("WM_DELETE_WINDOW", on_close)
    login_window.mainloop()

    return result

def open_password_change_window(parent, db_name, username):
    win = tk.Toplevel(parent)
    win.title("變更密碼")
    win.geometry("300x220")
    win.iconbitmap("PMS.ico")
    win.resizable(False, False)

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

        old_hash = hashlib.sha256(old_pw.encode()).hexdigest()
        new_hash = hashlib.sha256(new_pw.encode()).hexdigest()

        with sqlite3.connect(db_name,timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            cursor = conn.cursor()
            cursor.execute("SELECT password FROM users WHERE username=? AND password=?", (username, old_hash))
            if not cursor.fetchone():
                messagebox.showerror("錯誤", "舊密碼不正確")
                return

            cursor.execute("UPDATE users SET password=? WHERE username=?", (new_hash, username))
            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        messagebox.showinfo("成功", "密碼已變更")
        win.destroy()

    tk.Button(win, text="變更密碼", bg="lightgreen", command=confirm_change).pack(pady=15)

if __name__ == "__main__":
    if is_another_instance_running():
        messagebox.showerror("錯誤", "本程式已在執行中，請勿重複開啟。")
        sys.exit()
    init_db()
    initialize_database()
    login_info = login()

    if login_info and login_info.get("user"):
        root = tk.Tk()
        root.title("生產管理平台")
        root.geometry("1200x750")

        IDLE_TIMEOUT_MS = 3 * 60 * 1000

        def reset_idle_timer(event=None):
            if hasattr(root, "_idle_after_id"):
                root.after_cancel(root._idle_after_id)
            if hasattr(root, "_warning_after_id"):
                root.after_cancel(root._warning_after_id)

        def on_idle_timeout():
            messagebox.showinfo("自動登出", "您已閒置超過 3 分鐘，系統將自動登出")
            logout_and_exit(root)

        for event_type in ["<Motion>", "<Key>", "<Button>"]:
            root.bind_all(event_type, reset_idle_timer)
        reset_idle_timer()

        try:
            root.iconbitmap("PMS.ico")
        except:
            pass
        import tkinter.font as tkFont
        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=10, family="Microsoft Calibri")

        top_bar = tk.Frame(root)
        top_bar.pack(fill="x", side="top")
        logout_btn = tk.Button(top_bar, text="登出並關閉", command=lambda: logout_and_exit(root), bg="orange")
        logout_btn.pack(side="right", padx=10, pady=5)
        change_pw_btn = tk.Button(top_bar, text="變更密碼", bg="lightgreen",
            command=lambda: open_password_change_window(root, DB_NAME, login_info["user"]))
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