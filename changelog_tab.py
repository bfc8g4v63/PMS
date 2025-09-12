#$ changelog_tab.py
#% 改版歷程分頁，顯示/管理 changelog 表內容
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import re

from schema_helper import get_next_changelog_version
import fixture_helper as FH

def build_changelog_tab(tab, current_role, db_name):
    frame = ttk.Frame(tab)
    frame.pack(fill="both", expand=True)

    tree = ttk.Treeview(frame, columns=("version", "date", "content"), show="headings")
    tree.heading("version", text="版本")
    tree.heading("date", text="時間")
    tree.heading("content", text="內容")
    tree.column("version", width=100)
    tree.column("date", width=150)
    tree.column("content", width=500)
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    entry_frame = ttk.Frame(frame)
    entry_frame.pack(fill="x", padx=10, pady=5)

    version_entry = ttk.Entry(entry_frame)
    content_entry = ttk.Entry(entry_frame, width=60)
    if current_role == "admin":
        ttk.Label(entry_frame, text="版本:").grid(row=0, column=0, sticky="e")
        version_entry.grid(row=0, column=1, sticky="w", padx=(5, 10))
        ttk.Label(entry_frame, text="內容:").grid(row=0, column=2, sticky="e")
        content_entry.grid(row=0, column=3, sticky="w", padx=(5, 10))

    status_var = tk.StringVar()
    status_label = ttk.Label(entry_frame, textvariable=status_var, foreground="blue")
    if current_role == "admin":
        status_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(5, 0))

    original_content = {"value": ""}

    def expected_next_version():
        return get_next_changelog_version(db_name)

    if current_role == "admin":
        def validate_version(*args):
            version = version_entry.get().strip()
            pattern = r"^v\d+\.\d+\.\d+$"
            if not re.match(pattern, version):
                add_button.config(state=tk.DISABLED)
                return
            major, minor, patch = map(int, version.lstrip("v").split("."))
            if (major, minor, patch) < (1, 0, 0):
                add_button.config(state=tk.DISABLED)
                return
            with FH.get_conn(db_name) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM changelog WHERE version = ?", (version,))
                exists = cursor.fetchone()[0] > 0
                if exists:
                    add_button.config(state=tk.DISABLED)
                    return
            expected = expected_next_version()
            if version != expected:
                add_button.config(state=tk.DISABLED)
                return
            add_button.config(state=tk.NORMAL)

        def on_content_change(*args):
            current = content_entry.get().strip()
            if current == original_content["value"] or not version_entry.get().strip():
                update_button.config(state=tk.DISABLED)
            else:
                update_button.config(state=tk.NORMAL)

        version_entry.bind("<KeyRelease>", validate_version)
        content_entry.bind("<KeyRelease>", on_content_change)

        def insert_changelog():
            version = version_entry.get().strip()
            content = content_entry.get().strip()
            if not version or not content:
                messagebox.showwarning("欄位缺漏", "請輸入版本與內容")
                return
            pattern = r"^v\d+\.\d+\.\d+$"
            if not re.match(pattern, version):
                messagebox.showerror("版本格式錯誤", "版本格式必須為 vX.Y.Z")
                return
            major, minor, patch = map(int, version.lstrip("v").split("."))
            if (major, minor, patch) < (1, 0, 0):
                messagebox.showerror("版本過低", "版本不得低於 v1.0.0")
                return
            expected = expected_next_version()
            if version != expected:
                messagebox.showerror("版本跳號", f"不可以跳版。下一個合法版本應該是 {expected}")
                return
            try:
                with FH.get_conn(db_name) as conn:
                    cursor = conn.cursor()
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    cursor.execute(
                        "INSERT INTO changelog (version, date, content) VALUES (?, ?, ?)",
                        (version, now, content),
                    )
                    conn.commit()
            except Exception:
                messagebox.showerror("版本重複", f"版本 {version} 已存在，請重新輸入")
                return
            refresh_changelog()
            version_entry.delete(0, tk.END)
            content_entry.delete(0, tk.END)
            add_button.config(state=tk.DISABLED)
            status_var.set("")

        def update_changelog():
            version = version_entry.get().strip()
            content = content_entry.get().strip()
            if not version or not content:
                messagebox.showwarning("欄位缺漏", "請輸入版本與內容")
                return
            with FH.get_conn(db_name) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE changelog SET content = ? WHERE version = ?",
                    (content, version),
                )
                conn.commit()
            refresh_changelog()
            messagebox.showinfo("更新成功", f"版本 {version} 的內容已更新")
            version_entry.delete(0, tk.END)
            content_entry.delete(0, tk.END)
            update_button.config(state=tk.DISABLED)
            status_var.set("")
            original_content["value"] = ""

        def delete_selected():
            selected = tree.selection()
            if not selected:
                return
            item = tree.item(selected)
            version = item["values"][0]
            confirm = messagebox.askyesno("確認刪除", f"是否刪除版本 {version}？")
            if confirm:
                with FH.get_conn(db_name) as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM changelog WHERE version = ?", (version,))
                    conn.commit()
                refresh_changelog()

        def fill_next_version():
            next_version = expected_next_version()
            version_entry.delete(0, tk.END)
            version_entry.insert(0, next_version)
            validate_version()

        auto_button = ttk.Button(entry_frame, text="自動產生版本", command=fill_next_version)
        auto_button.grid(row=0, column=4, padx=(10, 0), sticky="w")

        add_button = ttk.Button(entry_frame, text="新增紀錄", command=insert_changelog, state=tk.DISABLED)
        add_button.grid(row=0, column=5, padx=(10, 0), sticky="w")

        update_button = ttk.Button(entry_frame, text="儲存修改", command=update_changelog, state=tk.DISABLED)
        update_button.grid(row=0, column=6, padx=(10, 0), sticky="w")

        delete_button = ttk.Button(entry_frame, text="刪除所選", command=delete_selected)
        delete_button.grid(row=0, column=7, padx=(10, 0), sticky="w")

    def refresh_changelog():
        tree.delete(*tree.get_children())
        with FH.get_conn(db_name) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT version, date, content FROM changelog ORDER BY rowid DESC")
            for row in cursor.fetchall():
                tree.insert("", "end", values=row)

    def on_tree_double_click(event):
        if current_role != "admin":
            return
        item = tree.selection()
        if not item:
            return
        values = tree.item(item, "values")
        version_entry.delete(0, tk.END)
        content_entry.delete(0, tk.END)
        version_entry.insert(0, values[0])
        content_entry.insert(0, values[2])
        status_var.set(f"目前編輯中版本：{values[0]}")
        original_content["value"] = values[2]
        validate_version()
        if current_role == "admin":
            content_entry.bind("<KeyRelease>", lambda e: update_button.config(state=tk.NORMAL))

    tree.bind("<Double-1>", on_tree_double_click)

    refresh_changelog()