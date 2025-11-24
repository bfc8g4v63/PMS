#$ fixture_logs_tab.py
#% 治具紀錄GUI

import tkinter as tk
from tkinter import ttk, messagebox
from fixture_logger import get_fixture_logs, delete_fixture_logs
from date_picker import open_date_picker
from datetime import datetime

def build_fixture_logs_tab(frame, refresh_only=False):
    if refresh_only:
        for widget in frame.winfo_children():
            widget.destroy()
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
    action_combo["values"] = ("", "建立治具", "刪除治具", "修改資料", "入庫", "調撥")
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
        "時間"
    )

    tree = ttk.Treeview(frame, columns=columns, show="headings")
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
        elif col == "時間":
            tree.column(col, width=150, anchor="center")
        elif col == "單號":
            tree.column(col, width=120, anchor="center")
        else:
            tree.column(col, width=100, anchor="center")

    tree.pack(fill="both", expand=True, padx=5, pady=5)

    def on_query():
        try:
            start_text = start_date_var.get().strip()
            end_text = end_date_var.get().strip()

            start_ts = None
            end_ts = None

            if start_text:
                dt = datetime.strptime(start_text, "%Y-%m-%d")
                start_ts = dt.strftime("%Y%m%dT000000")
            if end_text:
                dt = datetime.strptime(end_text, "%Y-%m-%d")
                end_ts = dt.strftime("%Y%m%dT235959")

            logs = get_fixture_logs(
                part_no=part_no_var.get().strip() or None,
                user=user_var.get().strip() or None,
                action=action_var.get().strip() or None,
                start_ts=start_ts,
                end_ts=end_ts,
                limit=200
            )
            for row in tree.get_children():
                tree.delete(row)
            for r in logs:
                safe_values = [("" if v is None else str(v).strip("'")) for v in r]
                tree.insert("", "end", text="", values=safe_values)
            if logs:
                tree.focus(tree.get_children()[0])
                tree.selection_set(tree.get_children()[0])
                tree.yview_moveto(0)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_delete_selected():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提醒", "請先選擇要刪除的紀錄")
            return
        if not messagebox.askyesno("確認", "確定要刪除所選紀錄？"):
            return
        try:
            ids = [tree.item(i, "values")[0] for i in sel]
            delete_fixture_logs(ids)
            on_query()
            messagebox.showinfo("完成", "所選紀錄已刪除")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            
    def on_delete_all():
        if not messagebox.askyesno("確認", "確定要刪除全部紀錄？"):
            return
        try:
            delete_fixture_logs()
            on_query()
            messagebox.showinfo("完成", "全部紀錄已刪除")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", padx=5, pady=5)
    ttk.Button(btn_frame, text="查詢", command=on_query).pack(side="left", padx=5)
    ttk.Button(
        btn_frame,
        text="重置",
        command=lambda: [
            part_no_var.set(""),
            user_var.set(""),
            action_var.set(""),
            start_date_var.set(""),
            end_date_var.set(""),
            on_query()
        ],
    ).pack(side="left", padx=5)

    ttk.Button(btn_frame, text="刪除所選", command=on_delete_selected).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="刪除全部", command=on_delete_all).pack(side="left", padx=5)

    on_query()