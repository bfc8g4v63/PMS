#$ sop_logs_tab.py
#% SOP紀錄分頁（查詢、排序、開檔、管理刪除）

import os
import tkinter as tk
from tkinter import ttk, messagebox

from db_helper import get_conn
from utils import open_file


def build_sop_logs_tab(tab, db_name, role, base_paths):
    tk.Label(tab, text="SOP紀錄查詢").pack(anchor="w", padx=10, pady=(10, 0))

    search_frame = tk.Frame(tab)
    search_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(search_frame, text="查詢關鍵字:").pack(side="left")
    entry_query = tk.Entry(search_frame)
    entry_query.pack(side="left")

    sort_desc = tk.BooleanVar(value=True)

    def toggle_sort():
        sort_desc.set(not sort_desc.get())
        refresh_logs(full=last_full_fetch.get())

    tk.Button(search_frame, text="↕排序", command=toggle_sort).pack(side="left", padx=5)
    tk.Button(search_frame, text="查詢", command=lambda: refresh_logs(full=True)).pack(side="left")

    columns = ("SOP建立人", "動作", "檔案名稱", "時間")

    tree_container = tk.Frame(tab)
    tree_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    tree = ttk.Treeview(tree_container, columns=columns, show="headings")
    tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll_y.set)

    for col in columns:
        tree.heading(col, text=col)
    tree.column("SOP建立人", width=80, anchor="center")
    tree.column("動作", width=120, anchor="center")
    tree.column("檔案名稱", width=520, anchor="w")
    tree.column("時間", width=160, anchor="center")

    tree.pack(side="left", fill="both", expand=True)
    tree_scroll_y.pack(side="right", fill="y")

    last_full_fetch = tk.BooleanVar(value=False)

    def _detect_log_schema(conn):
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(activity_logs)")
        cols = [r[1] for r in cur.fetchall()]

        prefixed = {
            "id": "activity_log_id",
            "user": "activity_log_username",
            "action": "activity_log_action",
            "filename": "activity_log_filename",
            "ts": "activity_log_timestamp",
            "module": "activity_log_module",
        }

        legacy = {
            "id": "id",
            "user": "username",
            "action": "action",
            "filename": "filename",
            "ts": "timestamp",
            "module": "module",
        }

        if prefixed["user"] in cols and prefixed["action"] in cols and prefixed["filename"] in cols and prefixed["ts"] in cols:
            return prefixed

        if legacy["user"] in cols and legacy["action"] in cols and legacy["filename"] in cols and legacy["ts"] in cols:
            return legacy

        return None

    def refresh_logs(full=False):
        keyword = entry_query.get().strip()
        last_full_fetch.set(bool(full) or bool(keyword))

        for row in tree.get_children():
            tree.delete(row)

        action_map_local = {
            "add_user": "新增使用者",
            "update_user": "修改使用者",
            "delete_user": "刪除使用者",
            "update_permissions": "修改權限",
            "set_permissions": "修改權限",
            "change_password": "變更密碼",
            "upload": "新增 SOP",
            "update_sop": "更新 SOP",
            "generate_sop": "生成 SOP",
            "apply_sop": "套用 SOP",
            "delete": "刪除紀錄",
            "login": "登入系統",
            "logout": "登出系統",
            "toggle_bypass": "啟用/停用 SOP",
        }

        restricted_roles = ("engineer", "leader")
        restricted_actions = (
            "add_user",
            "update_user",
            "delete_user",
            "disable_user",
            "enable_user",
            "update_permissions",
            "set_permissions",
            "change_password",
        )

        with get_conn() as conn:
            schema = _detect_log_schema(conn)
            if not schema:
                messagebox.showerror("錯誤", "activity_logs 欄位結構無法辨識，請確認資料表欄位是否正確。")
                return

            c = conn.cursor()

            sql = (
                "SELECT "
                f"{schema['id']} AS log_id, "
                f"{schema['user']} AS log_user, "
                f"{schema['action']} AS log_action, "
                f"{schema['filename']} AS log_filename, "
                f"{schema['ts']} AS log_ts "
                "FROM activity_logs"
            )

            params = []

            if role in restricted_roles:
                if keyword:
                    sql += " WHERE (log_user LIKE ? OR log_action LIKE ? OR log_filename LIKE ?)"
                    sql += " AND log_action NOT IN ({})".format(",".join("?" for _ in restricted_actions))
                    params = [f"%{keyword}%"] * 3 + list(restricted_actions)
                else:
                    sql += " WHERE log_action NOT IN ({})".format(",".join("?" for _ in restricted_actions))
                    params = list(restricted_actions)
            else:
                if keyword:
                    sql += " WHERE log_user LIKE ? OR log_action LIKE ? OR log_filename LIKE ?"
                    params = [f"%{keyword}%"] * 3

            sql += f" ORDER BY log_ts {'DESC' if sort_desc.get() else 'ASC'}"

            if not keyword and not full:
                sql += " LIMIT 200"

            c.execute(sql, params)

            for r in c.fetchall():
                log_id = r[0]
                log_user = (r[1] or "").strip()
                log_action_raw = (r[2] or "").strip()
                log_filename = (r[3] or "").strip()
                log_ts = (r[4] or "").strip()

                if log_id is None:
                    iid = f"temp_{log_ts}_{log_user}_{log_action_raw}"
                else:
                    iid = str(log_id)

                action_display = action_map_local.get(log_action_raw, log_action_raw or "未知")
                tree.insert("", "end", iid=iid, values=(log_user, action_display, log_filename, log_ts))

    def on_double_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index != 2:
            return

        filename = tree.item(item)["values"][2]
        if not filename:
            return

        for base in base_paths:
            path = os.path.join(base, filename)
            if os.path.exists(path):
                open_file(path)
                break

    tree.bind("<Double-1>", on_double_click)

    button_frame = tk.Frame(tab)
    button_frame.pack(anchor="e", padx=10, pady=(0, 10))

    tk.Button(button_frame, text="重新整理", command=lambda: refresh_logs(full=False)).pack(side="left", padx=5)

    if role == "admin":

        def delete_selected_log():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提醒", "請先選取一筆SOP紀錄")
                return
            if messagebox.askyesno("確認", "確定要刪除所選SOP紀錄？"):
                with get_conn() as conn:
                    schema = _detect_log_schema(conn)
                    if not schema:
                        messagebox.showerror("錯誤", "activity_logs 欄位結構無法辨識，請確認資料表欄位是否正確。")
                        return
                    c = conn.cursor()
                    for iid in selected:
                        if str(iid).isdigit():
                            c.execute(f"DELETE FROM activity_logs WHERE {schema['id']}=?", (int(iid),))
                    conn.commit()
                refresh_logs(full=True)

        def delete_all_logs():
            if messagebox.askyesno("確認", "確定要刪除所有SOP紀錄？此操作無法復原。"):
                with get_conn() as conn:
                    c = conn.cursor()
                    c.execute("DELETE FROM activity_logs")
                    conn.commit()
                refresh_logs(full=True)

        tk.Button(button_frame, text="刪除所選", command=delete_selected_log).pack(side="left", padx=5)
        tk.Button(button_frame, text="刪除全部", command=delete_all_logs).pack(side="left", padx=5)

    refresh_logs(full=False)