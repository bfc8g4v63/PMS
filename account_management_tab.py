import tkinter as tk
import sqlite3
import hashlib
import re
from tkinter import ttk, messagebox
from utils import log_activity

def build_user_management_tab(tab, db_name, current_user):
    PERMISSION_FLAGS = {
        "can_add": {"label": "新增", "default": 1},
        "can_delete": {"label": "刪除", "default": 0},
        "active": {"label": "啟用", "default": 1},
        "can_view_logs": {"label": "可見操作紀錄", "default": 1},
        "can_delete_logs": {"label": "刪除操作紀錄", "default": 0},
        "can_upload_sop": {"label": "上傳SOP", "default": 1},
        "can_view_issues": {"label": "可見生產資訊", "default": 1},
        "can_manage_users": {"label": "帳號管理", "default": 0}
    }

    frame = tk.Frame(tab)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    tk.Label(frame, text="帳號管理").pack(anchor="w")

    control_frame = tk.Frame(frame)
    control_frame.pack(fill="x", pady=5, padx=10)

    tk.Label(control_frame, text="顯示帳號：").pack(side="left")
    filter_var = tk.StringVar(value="全部")
    filter_combo = ttk.Combobox(control_frame, textvariable=filter_var, values=["全部", "僅啟用", "僅停用"], width=10, state="readonly")
    filter_combo.pack(side="left", padx=(0, 10))

    sort_asc = tk.BooleanVar(value=True)
    def toggle_sort():
        sort_asc.set(not sort_asc.get())
        refresh_users()

    tk.Button(control_frame, text="↕排序帳號", command=toggle_sort).pack(side="left")

    columns = ("帳號", "角色", "新增", "刪除", "啟用", "上傳SOP", "可見紀錄", "刪紀錄", "可見生產", "帳號管理")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    tree.pack(fill="both", expand=True, pady=5)

    def refresh_users():
        for row in tree.get_children():
            tree.delete(row)
        with sqlite3.connect(db_name, timeout=5) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            cursor = conn.cursor()
            sql = """SELECT username, role, can_add, can_delete, active,
                can_upload_sop, can_view_logs, can_delete_logs,
                can_view_issues, can_manage_users FROM users"""
            condition = filter_var.get()
            if condition == "僅啟用":
                sql += " WHERE active=1"
            elif condition == "僅停用":
                sql += " WHERE active=0"
            sql += f" ORDER BY username {'ASC' if sort_asc.get() else 'DESC'}"
            cursor.execute(sql)
            for row in cursor.fetchall():
                tags = ("disabled",) if row[4] == 0 else ()
                display_row = [
                    row[0],
                    row[1],
                    *["✓" if v else "✕" for v in row[2:]]
                ]
                tree.insert("", "end", values=display_row, tags=tags)

        tree.tag_configure("disabled", foreground="gray")

    filter_combo.bind("<<ComboboxSelected>>", lambda e: refresh_users())

    form_container = tk.Frame(frame)
    form_container.pack(fill="x", pady=10, padx=10)

    form = tk.LabelFrame(form_container, text="新增使用者")
    form.pack(side="left", fill="both", expand=True, padx=(0, 5))

    tk.Label(form, text="帳號：").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    entry_user = tk.Entry(form, width=30)
    entry_user.grid(row=0, column=1, columnspan=3, sticky="w", padx=5)

    tk.Label(form, text="密碼：").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_pass = tk.Entry(form, width=30, show="*")
    entry_pass.grid(row=1, column=1, columnspan=3, sticky="w", padx=5)

    tk.Label(form, text="角色：").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    role_var = tk.StringVar()
    role_menu = ttk.Combobox(form, textvariable=role_var, values=["admin", "engineer", "leader"], state="readonly", width=15)
    role_menu.grid(row=2, column=1, sticky="w", padx=5)

    tk.Label(form, text="專長：").grid(row=2, column=2, sticky="e", padx=5)
    specialty_var = tk.StringVar(value="")
    specialty_combo = ttk.Combobox(form, textvariable=specialty_var,
        values=["", "dip", "assembly", "test", "packaging", "oqc"], width=15, state="readonly")
    specialty_combo.grid(row=2, column=3, sticky="w", padx=5)

    permission_vars = {}
    row = 3
    col = 0
    for key, perm in PERMISSION_FLAGS.items():
        var = tk.IntVar(value=perm["default"])
        permission_vars[key] = var
        tk.Checkbutton(form, text=perm["label"], variable=var).grid(row=row, column=col, sticky="w", padx=5, pady=5)
        col += 1
        if col > 3:
            col = 0
            row += 1

    def collect_permission_values():
        return {k: v.get() for k, v in permission_vars.items()}

    def hash_password(password):
        return hashlib.sha256(password.encode("utf-8")).hexdigest()

    def add_user():
        new_user = entry_user.get().strip()
        new_pw = entry_pass.get().strip()
        role = role_var.get()

        permissions = collect_permission_values()
        can_add = permissions["can_add"]
        can_delete = permissions["can_delete"]
        active = permissions["active"]

        if not new_user or not new_pw:
            messagebox.showwarning("警告", "請填寫帳號與密碼")
            return
        if not role:
            messagebox.showwarning("警告", "請選擇角色")
            return
        if not re.match(r"^[A-Za-z0-9]{6,12}$", new_pw):
            messagebox.showerror("錯誤", "密碼須為6～12碼英文或數字組成")
            return
        hashed_pw = hash_password(new_pw)

        with sqlite3.connect(db_name,timeout=5) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            cursor = conn.cursor()
            cursor.execute("SELECT username FROM users WHERE username=?", (new_user,))
            if cursor.fetchone():
                messagebox.showerror("錯誤", "該使用者已存在")
                return

            fields = ["username", "password", "role", "specialty"] + list(permission_vars.keys())
            placeholders = ", ".join(["?"] * len(fields))
            sql = f"INSERT INTO users ({', '.join(fields)}) VALUES ({placeholders})"
            values = [new_user, hashed_pw, role, specialty_var.get()] + [permission_vars[k].get() for k in permission_vars]
            cursor.execute(sql, values)

            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")

        messagebox.showinfo("成功", "使用者已新增")
        log_activity(db_name, current_user, "add_user", new_user, module="帳號管理")

        entry_user.delete(0, tk.END)
        entry_pass.delete(0, tk.END)
        specialty_var.set("")
        for v in permission_vars.values():
            v.set(0)
        refresh_users()

    tk.Button(form, text="新增使用者", command=add_user, bg="lightblue").grid(row=5, column=1, pady=10)

    edit_frame = tk.LabelFrame(form_container, text="修改權限")
    edit_frame.pack(side="left", fill="both", expand=True)

    tk.Label(edit_frame, text="新帳號:").grid(row=0, column=0)
    entry_edit_user = tk.Entry(edit_frame)
    entry_edit_user.grid(row=0, column=1)

    tk.Label(edit_frame, text="新密碼:").grid(row=1, column=0)
    entry_edit_pass = tk.Entry(edit_frame, show="*")
    entry_edit_pass.grid(row=1, column=1)

    tk.Label(edit_frame, text="角色:").grid(row=2, column=0)
    role_edit = tk.StringVar()
    combo_role = ttk.Combobox(edit_frame, textvariable=role_edit, values=["admin", "engineer", "leader"], state="readonly")
    combo_role.grid(row=2, column=1)

    edit_specialty = tk.StringVar()
    tk.Label(edit_frame, text="專長:").grid(row=3, column=0)
    combo_specialty = ttk.Combobox(edit_frame, textvariable=edit_specialty, values=["", "dip", "assembly", "test", "packaging", "oqc"], state="readonly")
    combo_specialty.grid(row=3, column=1)

    def on_select_user(event):
        selected = tree.selection()
        if not selected:
            return
        item = tree.item(selected[0])["values"]
        username = item[0]

        entry_edit_user.delete(0, tk.END)
        entry_edit_user.insert(0, username)
        role_edit.set(item[1])

        with sqlite3.connect(db_name,timeout=5) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            cursor = conn.cursor()

            cursor.execute("SELECT specialty FROM users WHERE username=?", (username,))
            specialty = cursor.fetchone()
            edit_specialty.set(specialty[0] if specialty else "")

            cursor.execute(f"""
                SELECT {', '.join(permission_vars.keys())}
                FROM users WHERE username=?
            """, (username,))
            result = cursor.fetchone()
            if result:
                for i, key in enumerate(permission_vars):
                    permission_vars[key].set(result[i])

    tree.bind("<<TreeviewSelect>>", on_select_user)

    def update_user():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("未選擇", "請選擇帳號")
            return

        original_username = tree.item(selected[0])["values"][0]
        if original_username == current_user:
            messagebox.showerror("錯誤", "無法修改當前登入帳號")
            return

        new_username = entry_edit_user.get().strip()
        new_pass = entry_edit_pass.get().strip()
        role = role_edit.get()
        specialty = edit_specialty.get()

        permissions = {k: v.get() for k, v in permission_vars.items()}

        with sqlite3.connect(db_name, timeout=5) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            cursor = conn.cursor()

            if new_username and new_username != original_username:
                cursor.execute("SELECT username FROM users WHERE username=?", (new_username,))
                if cursor.fetchone():
                    messagebox.showerror("錯誤", "新帳號名稱已存在")
                    return
                cursor.execute("UPDATE users SET username=? WHERE username=?", (new_username, original_username))
                original_username = new_username

            if new_pass:
                if not re.match(r"^[A-Za-z0-9]{6,12}$", new_pass):
                    messagebox.showerror("錯誤", "新密碼須為6～12碼英文或數字組成")
                    return
                hashed_pw = hash_password(new_pass)
                cursor.execute("""
                    UPDATE users SET password=?, role=?, specialty=?,
                        can_add=?, can_delete=?, active=?,
                        can_view_logs=?, can_delete_logs=?, can_upload_sop=?,
                        can_view_issues=?, can_manage_users=?
                    WHERE username=?
                """, (
                    hashed_pw, role, specialty,
                    permissions["can_add"], permissions["can_delete"], permissions["active"],
                    permissions["can_view_logs"], permissions["can_delete_logs"], permissions["can_upload_sop"],
                    permissions["can_view_issues"], permissions["can_manage_users"],
                    original_username
                ))
            else:
                cursor.execute("""
                    UPDATE users SET role=?, specialty=?,
                        can_add=?, can_delete=?, active=?,
                        can_view_logs=?, can_delete_logs=?, can_upload_sop=?,
                        can_view_issues=?, can_manage_users=?
                    WHERE username=?
                """, (
                    role, specialty,
                    permissions["can_add"], permissions["can_delete"], permissions["active"],
                    permissions["can_view_logs"], permissions["can_delete_logs"], permissions["can_upload_sop"],
                    permissions["can_view_issues"], permissions["can_manage_users"],
                    original_username
                ))

            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        messagebox.showinfo("成功", "已更新")
        log_activity(db_name, current_user, "update_user", original_username)
        entry_edit_user.delete(0, tk.END)
        entry_edit_pass.delete(0, tk.END)
        refresh_users()

    def delete_user():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("未選擇", "請選擇帳號")
            return
        username = tree.item(selected[0])["values"][0]
        if username == current_user:
            messagebox.showerror("錯誤", "無法刪除自己")
            return
        if messagebox.askyesno("確認", f"是否確定要刪除帳號「{username}」？"):
            with sqlite3.connect(db_name,timeout=5) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                cursor = conn.cursor()
                cursor.execute("DELETE FROM users WHERE username=?", (username,))
                conn.commit()
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            messagebox.showinfo("成功", "使用者已刪除")
            log_activity(db_name, current_user, "delete_user", username)
            refresh_users()

    tk.Button(edit_frame, text="更新權限", command=update_user).grid(row=5, column=1, pady=5)
    tk.Button(edit_frame, text="刪除帳號", command=delete_user, bg="lightcoral", fg="white")\
        .grid(row=5, column=2, padx=10, pady=5)

    refresh_users()
    return tree, refresh_users