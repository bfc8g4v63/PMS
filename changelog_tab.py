#$ changelog_tab.py
#% 改版歷程分頁，顯示/管理 changelog 表內容

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import re

from db_helper import get_conn


def build_changelog_tab(tab, current_role, db_name):
    frame = ttk.Frame(tab)
    frame.pack(fill="both", expand=True)

    tree_container = ttk.Frame(frame)
    tree_container.pack(fill="both", expand=True, padx=10, pady=10)

    tree = ttk.Treeview(
        tree_container,
        columns=("version", "date", "content"),
        show="headings",
        selectmode="extended"
    )
    tree.heading("version", text="版本")
    tree.heading("date", text="時間")
    tree.heading("content", text="內容")
    tree.column("version", width=100, anchor="center")
    tree.column("date", width=150, anchor="center")
    tree.column("content", width=500)

    tree_yscroll = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_yscroll.set)

    tree_container.grid_rowconfigure(0, weight=1)
    tree_container.grid_columnconfigure(0, weight=1)

    tree.grid(row=0, column=0, sticky="nsew")
    tree_yscroll.grid(row=0, column=1, sticky="ns")

    entry_frame = ttk.Frame(frame)
    entry_frame.pack(fill="x", padx=10, pady=5)

    search_frame = ttk.Frame(frame)
    search_frame.pack(fill="x", padx=10, pady=(0, 5))

    tk.Label(search_frame, text="查詢(版本/內容關鍵字):").pack(side="left")
    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
    search_entry.pack(side="left", padx=(5, 10))

    version_entry = ttk.Entry(entry_frame, width=15)
    date_entry = ttk.Entry(entry_frame, width=20)
    content_entry = ttk.Entry(entry_frame, width=80)

    bump_var = tk.StringVar(value="patch")
    current_version_var = tk.StringVar(value="")

    if current_role == "admin":
        ttk.Label(entry_frame, text="版本:").grid(row=0, column=0, sticky="e")
        version_entry.grid(row=0, column=1, sticky="w", padx=(5, 10))

        ttk.Label(entry_frame, text="進版:").grid(row=0, column=2, sticky="e")
        bump_combo = ttk.Combobox(
            entry_frame,
            textvariable=bump_var,
            values=["patch", "minor", "major"],
            state="readonly",
            width=8
        )
        bump_combo.grid(row=0, column=3, sticky="w", padx=(5, 10))

        ttk.Label(entry_frame, textvariable=current_version_var).grid(row=0, column=4, sticky="w", padx=(0, 10))

        ttk.Label(entry_frame, text="時間:").grid(row=0, column=5, sticky="e")
        date_entry.grid(row=0, column=6, sticky="w", padx=(5, 10))

        ttk.Label(entry_frame, text="內容:").grid(row=1, column=0, sticky="e", pady=(5, 0))
        content_entry.grid(row=1, column=1, columnspan=8, sticky="we", padx=(5, 10), pady=(5, 0))

    status_var = tk.StringVar()
    status_label = ttk.Label(entry_frame, textvariable=status_var, foreground="blue")
    if current_role == "admin":
        status_label.grid(row=2, column=0, columnspan=6, sticky="w", pady=(5, 0))

    original_content = {"value": ""}

    def _parse_version(v):
        m = re.match(r"^v(\d+)\.(\d+)\.(\d+)$", (v or "").strip())
        if not m:
            return None
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))

    def _format_version(t):
        return f"v{t[0]}.{t[1]}.{t[2]}"

    def _get_latest_version_from_db():
        versions = []
        with get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT version FROM change_log")
            for (v,) in cursor.fetchall():
                t = _parse_version(v)
                if t is not None:
                    versions.append(t)
        if not versions:
            return (0, 0, 0)
        return max(versions)

    def _get_next_version_tuple(bump):
        major, minor, patch = _get_latest_version_from_db()
        if bump == "major":
            return (major + 1, 0, 0)
        if bump == "minor":
            return (major, minor + 1, 0)
        return (major, minor, patch + 1)

    def _refresh_current_version_label():
        current_version_var.set(f"目前版本: {_format_version(_get_latest_version_from_db())}")

    def expected_next_version():
        return _format_version(_get_next_version_tuple(bump_var.get()))

    if current_role == "admin":
        def validate_version(*args):
            version = version_entry.get().strip()
            pattern = r"^v\d+\.\d+\.\d+$"
            if not re.match(pattern, version):
                add_button.config(state=tk.DISABLED)
                return
            major, minor, patch = map(int, version.lstrip("v").split("."))
            if (major, minor, patch) < (0, 0, 1):
                add_button.config(state=tk.DISABLED)
                return
            with get_conn() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM change_log WHERE version = ?", (version,))
                exists = cursor.fetchone()[0] > 0
                if exists:
                    add_button.config(state=tk.DISABLED)
                    return
            expected = expected_next_version()
            if version != expected:
                add_button.config(state=tk.DISABLED)
                return
            add_button.config(state=tk.NORMAL)

        def on_field_change(*args):
            current_content = content_entry.get().strip()
            current_date = date_entry.get().strip()
            if (
                (current_content == original_content["value"] and current_date == original_content.get("date", ""))
                or not version_entry.get().strip()
            ):
                update_button.config(state=tk.DISABLED)
            else:
                update_button.config(state=tk.NORMAL)

        content_entry.bind("<KeyRelease>", on_field_change)
        date_entry.bind("<KeyRelease>", on_field_change)

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
            if (major, minor, patch) < (0, 0, 1):
                messagebox.showerror("版本過低", "版本不得低於 v0.0.1")
                return
            expected = expected_next_version()
            if version != expected:
                messagebox.showerror("版本跳號", f"不可以跳版。下一個合法版本應該是 {expected}")
                return
            try:
                date_val = date_entry.get().strip()
                if not date_val:
                    date_val = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                else:
                    try:
                        datetime.strptime(date_val, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        messagebox.showerror("時間格式錯誤", "時間必須是 YYYY-MM-DD HH:MM:SS 格式")
                        return
                with get_conn() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "INSERT INTO change_log (version, date, content) VALUES (?, ?, ?)",
                        (version, date_val, content),
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
            original_content["value"] = ""
            _refresh_current_version_label()

        def update_changelog():
            version = version_entry.get().strip()
            content = content_entry.get().strip()
            if not version or not content:
                messagebox.showwarning("欄位缺漏", "請輸入版本與內容")
                return
            with get_conn() as conn:
                cursor = conn.cursor()
                date_val = date_entry.get().strip()
                if not date_val:
                    messagebox.showwarning("欄位缺漏", "請輸入時間")
                    return
                try:
                    datetime.strptime(date_val, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    messagebox.showerror("時間格式錯誤", "時間必須是 YYYY-MM-DD HH:MM:SS 格式")
                    return
                cursor.execute(
                    "UPDATE change_log SET content = ?, date = ? WHERE version = ?",
                    (content, date_val, version),
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
            selected_ids = tree.selection()
            if not selected_ids:
                return
            versions = []
            for iid in selected_ids:
                values = tree.item(iid, "values")
                if values and len(values) >= 1 and values[0]:
                    versions.append(values[0])
            versions = [v for v in versions if v]
            if not versions:
                return
            if len(versions) == 1:
                confirm_msg = f"是否刪除版本 {versions[0]}？"
            else:
                confirm_msg = f"是否刪除所選 {len(versions)} 筆版本？"
            confirm = messagebox.askyesno("確認刪除", confirm_msg)
            if not confirm:
                return
            with get_conn() as conn:
                cursor = conn.cursor()
                cursor.executemany("DELETE FROM change_log WHERE version = ?", [(v,) for v in versions])
                conn.commit()
            refresh_changelog()
            _refresh_current_version_label()

        def fill_next_version():
            next_version = expected_next_version()
            version_entry.delete(0, tk.END)
            version_entry.insert(0, next_version)
            validate_version()

        def on_bump_change(event=None):
            fill_next_version()

        bump_combo.bind("<<ComboboxSelected>>", on_bump_change)

        auto_button = ttk.Button(entry_frame, text="產生版本", command=fill_next_version)
        auto_button.grid(row=0, column=7, padx=10, sticky="w")

        add_button = ttk.Button(entry_frame, text="新增紀錄", command=insert_changelog, state=tk.DISABLED)
        add_button.grid(row=0, column=8, padx=10, sticky="w")

        update_button = ttk.Button(entry_frame, text="儲存修改", command=update_changelog, state=tk.DISABLED)
        update_button.grid(row=0, column=9, padx=10, sticky="w")

        delete_button = ttk.Button(entry_frame, text="刪除所選", command=delete_selected)
        delete_button.grid(row=0, column=10, padx=10, sticky="w")

    def refresh_changelog(keyword=None, limit=200):
        tree.delete(*tree.get_children())
        kw = (keyword or "").strip()
        with get_conn() as conn:
            cursor = conn.cursor()
            if kw:
                cursor.execute(
                    "SELECT version, date, content FROM change_log "
                    "WHERE version LIKE ? OR content LIKE ? "
                    "ORDER BY rowid DESC LIMIT ?",
                    (f"%{kw}%", f"%{kw}%", int(limit or 200)),
                )
            else:
                if limit is None:
                    cursor.execute(
                        "SELECT version, date, content FROM change_log "
                        "ORDER BY rowid DESC"
                    )
                else:
                    cursor.execute(
                        "SELECT version, date, content FROM change_log "
                        "ORDER BY rowid DESC LIMIT ?",
                        (int(limit),),
                    )
            for row in cursor.fetchall():
                tree.insert("", "end", values=row)
        children = tree.get_children()
        if children:
            tree.selection_set(children[0])
            tree.focus(children[0])
            tree.see(children[0])

    def on_search():
        kw = search_var.get().strip()
        if kw:
            refresh_changelog(keyword=kw, limit=200)
        else:
            refresh_changelog(keyword=None, limit=None)

    def on_search_reset():
        search_var.set("")
        refresh_changelog()

    ttk.Button(search_frame, text="查詢", command=on_search).pack(side="left")
    ttk.Button(search_frame, text="重置", command=on_search_reset).pack(side="left", padx=(5, 0))

    def on_tree_double_click(event):
        if current_role != "admin":
            return
        items = tree.selection()
        if not items:
            return
        values = tree.item(items[0], "values")
        if not values:
            return
        version_entry.delete(0, tk.END)
        date_entry.delete(0, tk.END)
        content_entry.delete(0, tk.END)
        version_entry.insert(0, values[0])
        date_entry.insert(0, values[1])
        content_entry.insert(0, values[2])
        status_var.set(f"目前編輯中版本：{values[0]}")
        original_content["value"] = values[2]
        original_content["date"] = values[1]
        validate_version()
        content_entry.bind("<KeyRelease>", lambda e: update_button.config(state=tk.NORMAL))

    tree.bind("<Double-1>", on_tree_double_click)

    refresh_changelog()
    if current_role == "admin":
        _refresh_current_version_label()
        fill_next_version()