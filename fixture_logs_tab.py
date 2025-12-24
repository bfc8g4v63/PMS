#$ fixture_logs_tab.py
#% 治具紀錄GUI

import tkinter as tk
from tkinter import ttk, messagebox
from fixture_logger import get_fixture_logs, delete_fixture_logs
from date_picker import open_date_picker
from datetime import datetime
from db_helper import tx

def build_fixture_logs_tab(frame, refresh_only=False, current_user=None):
    if refresh_only:
        for widget in frame.winfo_children():
            widget.destroy()

    def _normalize_username(u):
        if u is None:
            return ""
        if isinstance(u, str):
            return u.strip()
        if isinstance(u, dict):
            v = u.get("user") or u.get("username") or ""
            return str(v).strip()
        return str(u).strip()

    username = _normalize_username(current_user)

    def _get_user_permissions(user_name: str, ctx):
        flags = {
            "can_view_fixture_logs": 0,
            "can_delete_fixture_logs": 0
        }

        if isinstance(ctx, dict):
            flags["can_view_fixture_logs"] = int(ctx.get("can_view_fixture_logs", 0) or 0)
            flags["can_delete_fixture_logs"] = int(ctx.get("can_delete_fixture_logs", 0) or 0)
            return flags

        if not user_name:
            return flags

        with tx() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT "
                "COALESCE(can_view_fixture_logs, 0), "
                "COALESCE(can_delete_fixture_logs, 0) "
                "FROM users WHERE username=?",
                (user_name,),
            )
            row = cur.fetchone()
            if not row:
                return flags
            flags["can_view_fixture_logs"] = int(row[0] or 0)
            flags["can_delete_fixture_logs"] = int(row[1] or 0)
            return flags

    perms = _get_user_permissions(username, current_user)

    if int(perms.get("can_view_fixture_logs") or 0) != 1:
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        ttk.Label(holder, text="無權限查看治具紀錄").pack(padx=10, pady=10)
        return

    query_frame = ttk.LabelFrame(frame, text="查詢條件")
    query_frame.pack(fill="x", padx=5, pady=5)

    tk.Label(query_frame, text="治具料號:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
    part_no_var = tk.StringVar()
    tk.Entry(query_frame, textvariable=part_no_var, width=15).grid(row=0, column=1, padx=5, pady=2)

    tk.Label(query_frame, text="治具操作人:").grid(row=0, column=2, padx=5, pady=2, sticky="e")
    user_var = tk.StringVar()
    tk.Entry(query_frame, textvariable=user_var, width=15).grid(row=0, column=3, padx=5, pady=2)

    tk.Label(query_frame, text="動作:").grid(row=0, column=4, padx=5, pady=2, sticky="e")
    action_var = tk.StringVar()
    action_combo = ttk.Combobox(query_frame, textvariable=action_var, width=12, state="readonly")
    action_combo["values"] = ("", "建立治具", "刪除治具", "修改資料", "入庫", "調撥", "調帳")
    action_combo.grid(row=0, column=5, padx=5, pady=2)

    tk.Label(query_frame, text="起始日期:").grid(row=1, column=0, padx=5, pady=2, sticky="e")
    start_date_var = tk.StringVar()
    start_entry = tk.Entry(query_frame, textvariable=start_date_var, width=15)
    start_entry.grid(row=1, column=1, padx=5, pady=2)

    def pick_start_date():
        open_date_picker(frame, "選擇起始日期", start_date_var.get().strip(), lambda v: start_date_var.set(v))

    ttk.Button(query_frame, text="選擇", command=pick_start_date).grid(row=1, column=2, padx=5, pady=2)

    tk.Label(query_frame, text="結束日期:").grid(row=1, column=3, padx=5, pady=2, sticky="e")
    end_date_var = tk.StringVar()
    end_entry = tk.Entry(query_frame, textvariable=end_date_var, width=15)
    end_entry.grid(row=1, column=4, padx=5, pady=2)

    def pick_end_date():
        open_date_picker(frame, "選擇結束日期", end_date_var.get().strip(), lambda v: end_date_var.set(v))

    ttk.Button(query_frame, text="選擇", command=pick_end_date).grid(row=1, column=5, padx=5, pady=2)

    columns = (
        "單號",
        "治具料號",
        "治具品名",
        "治具規格",
        "動作",
        "異動數量",
        "來源倉",
        "目標倉",
        "治具操作人",
        "調帳原因",
        "時間"
    )

    tree_container = ttk.Frame(frame)
    tree_container.pack(fill="both", expand=True, padx=5, pady=5)

    tree = ttk.Treeview(tree_container, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        if col == "治具料號":
            tree.column(col, width=80, anchor="center")
        elif col in ("治具品名", "治具規格"):
            tree.column(col, width=180, anchor="center")
        elif col == "動作":
            tree.column(col, width=80, anchor="center")
        elif col == "異動數量":
            tree.column(col, width=80, anchor="center")
        elif col in ("來源倉", "目標倉"):
            tree.column(col, width=100, anchor="center")
        elif col == "治具操作人":
            tree.column(col, width=100, anchor="center")
        elif col == "調帳原因":
            tree.column(col, width=120, anchor="center")
        elif col == "時間":
            tree.column(col, width=150, anchor="center")
        elif col == "單號":
            tree.column(col, width=120, anchor="center")
        else:
            tree.column(col, width=100, anchor="center")

    yscroll = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

    tree_container.grid_rowconfigure(0, weight=1)
    tree_container.grid_columnconfigure(0, weight=1)

    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew")

    def _has_query_filters():
        if part_no_var.get().strip():
            return True
        if user_var.get().strip():
            return True
        if action_var.get().strip():
            return True
        if start_date_var.get().strip():
            return True
        if end_date_var.get().strip():
            return True
        return False

    def _parse_ymd(text: str):
        s = (text or "").strip()
        if not s:
            return None

        try:
            return datetime.strptime(s, "%Y-%m-%d")
        except ValueError:
            raise ValueError("日期格式需為 YYYY-MM-DD")

    def _align_row_values(row, col_count: int):
        r = list(row) if row is not None else []
        if len(r) == 10 and col_count == 11:
            r = r[:9] + [""] + r[9:]
        if len(r) < col_count:
            r = r + ([""] * (col_count - len(r)))
        if len(r) > col_count:
            r = r[:col_count]
        return [("" if v is None else str(v)) for v in r]

    def on_query():
        try:
            start_text = start_date_var.get().strip()
            end_text = end_date_var.get().strip()

            start_ts = None
            end_ts = None

            if start_text:
                dt = _parse_ymd(start_text)
                start_ts = dt.strftime("%Y%m%dT000000")
            if end_text:
                dt = _parse_ymd(end_text)
                end_ts = dt.strftime("%Y%m%dT235959")

            action_filter = action_var.get().strip() or None

            limit_value = None if _has_query_filters() else 200
            logs = get_fixture_logs(
                part_no=part_no_var.get().strip() or None,
                user=user_var.get().strip() or None,
                action=action_filter,
                start_ts=start_ts,
                end_ts=end_ts,
                limit=limit_value
            )

            for row in tree.get_children():
                tree.delete(row)

            for r in logs:
                safe_values = _align_row_values(r, len(columns))
                tree.insert("", "end", text="", values=safe_values)

            if logs:
                first = tree.get_children()[0]
                tree.focus(first)
                tree.selection_set(first)
                tree.yview_moveto(0)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_delete_selected():
        if int(perms.get("can_delete_fixture_logs") or 0) != 1:
            messagebox.showerror("錯誤", "無刪除治具紀錄權限")
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提醒", "請先選擇要刪除的紀錄")
            return
        if not messagebox.askyesno("確認", "確定要刪除所選紀錄？"):
            return
        try:
            ids = []
            for i in sel:
                row_vals = tree.item(i, "values")
                if not row_vals:
                    continue
                log_id = row_vals[0]
                if log_id is None:
                    continue
                log_id = str(log_id).strip()
                if not log_id:
                    continue
                ids.append(log_id)

            if not ids:
                messagebox.showwarning("提醒", "所選資料沒有有效單號，無法刪除")
                return

            delete_fixture_logs(ids, deleted_by=username, note="GUI delete selected")
            on_query()
            messagebox.showinfo("完成", "所選紀錄已刪除")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            
    def on_delete_all():
        if int(perms.get("can_delete_fixture_logs") or 0) != 1:
            messagebox.showerror("錯誤", "無刪除治具紀錄權限")
            return
        if not messagebox.askyesno("確認", "確定要刪除全部紀錄？"):
            return
        try:
            delete_fixture_logs(None, deleted_by=username, note="GUI delete all")
            on_query()
            messagebox.showinfo("完成", "全部紀錄已刪除")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_reset():
        part_no_var.set("")
        user_var.set("")
        action_var.set("")
        start_date_var.set("")
        end_date_var.set("")
        on_query()

    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", padx=5, pady=5)

    ttk.Button(btn_frame, text="查詢", command=on_query).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="重置", command=on_reset).pack(side="left", padx=5)

    btn_delete_selected = ttk.Button(btn_frame, text="刪除所選", command=on_delete_selected)
    btn_delete_selected.pack(side="left", padx=5)

    btn_delete_all = ttk.Button(btn_frame, text="刪除全部", command=on_delete_all)
    btn_delete_all.pack(side="left", padx=5)

    if int(perms.get("can_delete_fixture_logs") or 0) != 1:
        btn_delete_selected.config(state="disabled")
        btn_delete_all.config(state="disabled")

    on_query()