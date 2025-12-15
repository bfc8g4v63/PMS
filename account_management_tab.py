#$ account_management_tab.py
#% 帳號管理分頁，包含新增/刪除/修改使用者、權限設定。

import tkinter as tk
import hashlib
import re
from tkinter import ttk, messagebox
from utils import log_activity
from schema_helper import auto_add_missing_columns, get_required_columns
from db_helper import get_conn


def build_user_management_tab(tab, db_name, current_user):

    auto_add_missing_columns(db_name, get_required_columns())

    PERMISSION_FLAGS = {
        "can_add": {"label": "新增SOP列", "default": 1},
        "can_delete": {"label": "刪除SOP列", "default": 0},
        "active": {"label": "帳號啟用", "default": 1},
        "can_view_logs": {"label": "可見SOP紀錄", "default": 1},
        "can_delete_logs": {"label": "刪除SOP紀錄", "default": 0},
        "can_upload_sop": {"label": "上傳SOP", "default": 1},
        "can_view_sop_info": {"label": "可見SOP資訊", "default": 1},
        "can_manage_users": {"label": "帳號管理", "default": 0},
        "can_view_fixture": {"label": "可見治具", "default": 1},
        "can_edit_fixture": {"label": "可編輯治具", "default": 0},
        "can_adjust_fixture": {"label": "可調帳治具", "default": 0},
        "can_view_fixture_logs": {"label": "可見治具紀錄", "default": 1},
        "can_delete_fixture_logs": {"label": "刪除治具紀錄", "default": 0}
    }

    def _normalize_current_user(u):
        if u is None:
            return ""
        if isinstance(u, str):
            return u.strip()
        if isinstance(u, dict):
            v = u.get("user") or u.get("username") or ""
            return str(v).strip()
        return str(u).strip()

    current_username = _normalize_current_user(current_user)

    def _open_conn():
        try:
            return get_conn(db_name)
        except TypeError:
            return get_conn()

    frame = tk.Frame(tab)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    tk.Label(frame, text="帳號管理").pack(anchor="w")

    control_frame = tk.Frame(frame)
    control_frame.pack(fill="x", pady=5, padx=10)

    tk.Label(control_frame, text="顯示帳號：").pack(side="left")
    filter_var = tk.StringVar(value="全部")
    filter_combo = ttk.Combobox(
        control_frame,
        textvariable=filter_var,
        values=["全部", "僅啟用", "僅停用"],
        width=10,
        state="readonly"
    )
    filter_combo.pack(side="left", padx=(0, 10))

    sort_asc = tk.BooleanVar(value=True)

    def toggle_sort():
        sort_asc.set(not sort_asc.get())
        refresh_users()

    tk.Button(control_frame, text="排序帳號", command=toggle_sort).pack(side="left")

    columns = (
        "帳號",
        "角色",
        "新增SOP列",
        "刪除SOP列",
        "帳號啟用",
        "上傳SOP",
        "SOP紀錄",
        "刪除紀錄",
        "SOP資訊",
        "帳號管理",
        "可見治具",
        "可編輯治具",
        "可調帳治具",
        "可見治具紀錄",
        "刪除治具紀錄"
    )

    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")
    tree.pack(fill="both", expand=True, pady=5)

    def _flag_text(v):
        return "Y" if int(v or 0) == 1 else "N"

    def refresh_users():
        for row in tree.get_children():
            tree.delete(row)
        with _open_conn() as conn:
            cursor = conn.cursor()
            sql = (
                "SELECT username, role, "
                "can_add, can_delete, active, "
                "can_upload_sop, can_view_logs, can_delete_logs, "
                "can_view_sop_info, can_manage_users, "
                "can_view_fixture, can_edit_fixture, can_adjust_fixture, "
                "can_view_fixture_logs, can_delete_fixture_logs "
                "FROM users"
            )
            condition = filter_var.get()
            if condition == "僅啟用":
                sql += " WHERE active=1"
            elif condition == "僅停用":
                sql += " WHERE active=0"
            sql += f" ORDER BY username {'ASC' if sort_asc.get() else 'DESC'}"
            cursor.execute(sql)
            for r in cursor.fetchall():
                tags = ("disabled",) if int(r[4] or 0) == 0 else ()
                flags = [_flag_text(v) for v in r[2:]]
                display_row = [r[0], r[1], *flags]
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
    specialty_combo = ttk.Combobox(
        form,
        textvariable=specialty_var,
        values=["", "dip", "assembly", "test", "packaging", "oqc"],
        width=15,
        state="readonly"
    )
    specialty_combo.grid(row=2, column=3, sticky="w", padx=5)

    permission_vars = {}
    row_idx = 3
    col_idx = 0
    for key, perm in PERMISSION_FLAGS.items():
        var = tk.IntVar(value=perm["default"])
        permission_vars[key] = var
        tk.Checkbutton(form, text=perm["label"], variable=var).grid(row=row_idx, column=col_idx, sticky="w", padx=5, pady=5)
        col_idx += 1
        if col_idx > 3:
            col_idx = 0
            row_idx += 1

    def collect_permission_values():
        return {k: v.get() for k, v in permission_vars.items()}

    def hash_password(password):
        return hashlib.sha256(password.encode("utf-8")).hexdigest()

    def add_user():
        new_user = entry_user.get().strip()
        new_pw = entry_pass.get().strip()
        role = role_var.get()
        permissions = collect_permission_values()

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

        with _open_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT username FROM users WHERE username=?", (new_user,))
            if cursor.fetchone():
                messagebox.showerror("錯誤", "該使用者已存在")
                return

            fields = ["username", "password", "role", "specialty"] + list(permission_vars.keys())
            placeholders = ", ".join(["?"] * len(fields))
            sql = f"INSERT INTO users ({', '.join(fields)}) VALUES ({placeholders})"
            values = [new_user, hashed_pw, role, specialty_var.get()] + [permissions[k] for k in permission_vars.keys()]
            cursor.execute(sql, values)
            conn.commit()

        messagebox.showinfo("成功", "使用者已新增")
        if current_username:
            log_activity(user=current_username, action="add_user", filename=new_user, module="帳號管理")

        entry_user.delete(0, tk.END)
        entry_pass.delete(0, tk.END)

        for k, v in permission_vars.items():
            v.set(PERMISSION_FLAGS[k]["default"])

        refresh_users()

    tk.Button(form, text="新增使用者", command=add_user, bg="lightblue").grid(row=row_idx + 2, column=1, pady=10)

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

        with _open_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT specialty FROM users WHERE username=?", (username,))
            specialty = cursor.fetchone()
            edit_specialty.set(specialty[0] if specialty else "")

            cursor.execute(
                "SELECT " + ", ".join(permission_vars.keys()) + " FROM users WHERE username=?",
                (username,)
            )
            result = cursor.fetchone()
            if result:
                for i, key in enumerate(permission_vars):
                    permission_vars[key].set(int(result[i] or 0))

    tree.bind("<<TreeviewSelect>>", on_select_user)

    def update_user():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("未選擇", "請選擇帳號")
            return

        original_username = tree.item(selected[0])["values"][0]
        if original_username == current_username and current_username:
            messagebox.showerror("錯誤", "無法修改當前登入帳號")
            return

        new_username = entry_edit_user.get().strip()
        new_pass = entry_edit_pass.get().strip()
        role = role_edit.get()
        specialty = edit_specialty.get()
        permissions = {k: v.get() for k, v in permission_vars.items()}

        if not role:
            messagebox.showerror("錯誤", "請選擇角色")
            return

        with _open_conn() as conn:
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
                sql = (
                    "UPDATE users SET password=?, role=?, specialty=?, "
                    "can_add=?, can_delete=?, active=?, "
                    "can_view_logs=?, can_delete_logs=?, can_upload_sop=?, "
                    "can_view_sop_info=?, can_manage_users=?, "
                    "can_view_fixture=?, can_edit_fixture=?, can_adjust_fixture=?, "
                    "can_view_fixture_logs=?, can_delete_fixture_logs=? "
                    "WHERE username=?"
                )
                params = (
                    hashed_pw, role, specialty,
                    permissions["can_add"], permissions["can_delete"], permissions["active"],
                    permissions["can_view_logs"], permissions["can_delete_logs"], permissions["can_upload_sop"],
                    permissions["can_view_sop_info"], permissions["can_manage_users"],
                    permissions["can_view_fixture"], permissions["can_edit_fixture"], permissions["can_adjust_fixture"],
                    permissions["can_view_fixture_logs"], permissions["can_delete_fixture_logs"],
                    original_username
                )
                cursor.execute(sql, params)
            else:
                sql = (
                    "UPDATE users SET role=?, specialty=?, "
                    "can_add=?, can_delete=?, active=?, "
                    "can_view_logs=?, can_delete_logs=?, can_upload_sop=?, "
                    "can_view_sop_info=?, can_manage_users=?, "
                    "can_view_fixture=?, can_edit_fixture=?, can_adjust_fixture=?, "
                    "can_view_fixture_logs=?, can_delete_fixture_logs=? "
                    "WHERE username=?"
                )
                params = (
                    role, specialty,
                    permissions["can_add"], permissions["can_delete"], permissions["active"],
                    permissions["can_view_logs"], permissions["can_delete_logs"], permissions["can_upload_sop"],
                    permissions["can_view_sop_info"], permissions["can_manage_users"],
                    permissions["can_view_fixture"], permissions["can_edit_fixture"], permissions["can_adjust_fixture"],
                    permissions["can_view_fixture_logs"], permissions["can_delete_fixture_logs"],
                    original_username
                )
                cursor.execute(sql, params)

            conn.commit()

        messagebox.showinfo("成功", "已更新")
        if current_username:
            log_activity(user=current_username, action="update_user", filename=original_username, module="帳號管理")
        entry_edit_user.delete(0, tk.END)
        entry_edit_pass.delete(0, tk.END)
        refresh_users()

    def delete_user():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("未選擇", "請選擇帳號")
            return
        username = tree.item(selected[0])["values"][0]
        if username == current_username and current_username:
            messagebox.showerror("錯誤", "無法刪除自己")
            return
        if messagebox.askyesno("確認", f"是否確定要停用帳號「{username}」？"):
            with _open_conn() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE users SET active=0 WHERE username=?", (username,))
                conn.commit()
            messagebox.showinfo("成功", "使用者已停用")
            if current_username:
                log_activity(user=current_username, action="disable_user", filename=username, module="帳號管理")
            refresh_users()

    tk.Button(edit_frame, text="更新權限", command=update_user).grid(row=5, column=1, pady=5)
    tk.Button(edit_frame, text="停用帳號", command=delete_user, bg="lightcoral", fg="white").grid(row=5, column=2, padx=10, pady=5)

    refresh_users()
    return tree, refresh_users