import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
import sqlite3
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from schema_helper import ensure_changelog_schema

def build_changelog_tab(tab, current_role, db_name):
    ensure_changelog_schema(db_name)

    tree = ttk.Treeview(tab, columns=("version", "date", "content"), show="headings", height=20)
    tree.heading("version", text="版本", anchor="w")
    tree.heading("date", text="時間", anchor="center")
    tree.heading("content", text="內容", anchor="w")
    tree.column("version", width=100)
    tree.column("date", width=160)
    tree.column("content", width=600)
    tree.pack(fill="both", expand=True)

    def load_changelog():
        tree.delete(*tree.get_children())
        with sqlite3.connect(db_name) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, version, date, content FROM changelog ORDER BY date DESC")
            for row in cursor.fetchall():
                tree.insert("", "end", iid=row[0], values=row[1:])
        tree.yview_moveto(0)

    def add_changelog():
        version = version_entry.get().strip()
        content = content_entry.get().strip()
        if not version or not content:
            messagebox.showwarning("欄位未填", "請輸入版本與內容")
            return
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(db_name, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("INSERT INTO changelog (version, date, content) VALUES (?, ?, ?)",
                         (version, date, content))
            conn.commit()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        version_entry.delete(0, tk.END)
        content_entry.delete(0, tk.END)
        load_changelog()

    def on_double_click(event):
        if current_role != "admin":
            return
        item_id = tree.focus()
        if not item_id:
            return
        version, date, content = tree.item(item_id, "values")
        new_version = simpledialog.askstring("修改版本", "請輸入新版本：", initialvalue=version)
        if new_version is None:
            return
        new_content = simpledialog.askstring("修改內容", "請輸入新內容：", initialvalue=content)
        if new_content is None:
            return
        with sqlite3.connect(db_name) as conn:
            conn.execute("UPDATE changelog SET version=?, content=? WHERE id=?",
                         (new_version, new_content, item_id))
            conn.commit()
        load_changelog()

    def on_right_click(event):
        if current_role != "admin":
            return
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_set(iid)
            menu = tk.Menu(tab, tearoff=0)
            menu.add_command(label="刪除紀錄", command=lambda: delete_changelog(iid))
            menu.add_command(label="編輯紀錄", command=lambda: on_double_click(None))
            menu.post(event.x_root, event.y_root)

    def delete_changelog(item_id):
        confirm = messagebox.askyesno("刪除確認", "確定要刪除此筆改版紀錄？")
        if not confirm:
            return
        with sqlite3.connect(db_name) as conn:
            conn.execute("DELETE FROM changelog WHERE id=?", (item_id,))
            conn.commit()
        load_changelog()

    tree.bind("<Double-1>", on_double_click)
    tree.bind("<Button-3>", on_right_click)

    if current_role == "admin":
        entry_frame = tk.Frame(tab)
        entry_frame.pack(fill="x", pady=5)

        tk.Label(entry_frame, text="版本：").grid(row=0, column=0)
        version_entry = tk.Entry(entry_frame)
        version_entry.grid(row=0, column=1)

        tk.Label(entry_frame, text="內容：").grid(row=1, column=0)
        content_entry = tk.Entry(entry_frame, width=80)
        content_entry.grid(row=1, column=1, columnspan=3)

        tk.Button(entry_frame, text="新增紀錄", command=add_changelog).grid(row=0, column=3, rowspan=2, padx=10)

    load_changelog()