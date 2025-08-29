# fixture_tabs.py
import tkinter as tk
from tkinter import ttk, messagebox
from fixture_helper import (
    insert_fixture,
    add_stock,
    transfer_stock,
    consume_stock,
    get_overview_by_warehouse,
    CORE_WAREHOUSES
)

WAREHOUSES = [
    "治具室", "上齊", "睿均", "捷曄", "立榮",
    "華勳", "上貿", "麥博", "GC", "不良品"
]

WAREHOUSES_ALL = CORE_WAREHOUSES + ["消耗"]

def _validate_part_no_local(part_no: str) -> None:
    p = (part_no or "").strip()
    if len(p) not in (8, 12) or (not p.isdigit()):
        raise ValueError("料號必須為 8 或 12 碼數字。")

def build_fixture_tab(parent: tk.Widget, current_user: str, db_name: str = None) -> tk.Frame:
    root = tk.Frame(parent)
    root.pack(fill="both", expand=True)
    control_frame = tk.Frame(root)
    control_frame.pack(fill="x", padx=8, pady=6)
    search_var = tk.StringVar()
    tk.Label(control_frame, text="依料號查詢").grid(row=0, column=0, sticky="w")
    tk.Entry(control_frame, textvariable=search_var, width=20).grid(row=0, column=1, padx=(6, 12))
    btn_search = tk.Button(control_frame, text="查詢", width=8)
    btn_search.grid(row=0, column=2, padx=4)
    btn_reset = tk.Button(control_frame, text="重置", width=8, command=lambda: [search_var.set(""), refresh_tables()])
    btn_reset.grid(row=0, column=3, padx=4)

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tree_views = {}

    def refresh_tables():
        try:
            part_filter = search_var.get().strip() or None
            data, warehouses = get_overview_by_warehouse(part_filter)
            for wh in CORE_WAREHOUSES:
                tree = tree_views.get(wh)
                if not tree: continue
                tree.delete(*tree.get_children())
                for row in data:
                    qty = row.get(wh, 0)
                    if qty > 0:
                        tree.insert("", "end", values=(
                            row["part_no"], row["name"], row["type"], row.get("spec", ""), qty
                        ))
        except Exception as e:
            messagebox.showerror("錯誤", f"更新表格失敗：{e}")

    def open_stockin_window():
        win = tk.Toplevel()
        win.title("⭫ 入庫操作")
        win.geometry("320x250")

        tk.Label(win, text="選擇料號", font=("Arial", 10)).pack(pady=10)

        frm = ttk.Frame(win)
        frm.pack(pady=5)

        tk.Label(frm, text="料號:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        part_combo = ttk.Combobox(frm, values=[], state="readonly", width=20)
        part_combo.grid(row=0, column=1)

        tk.Label(frm, text="數量:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        qty_entry = ttk.Entry(frm)
        qty_entry.grid(row=1, column=1)

        tk.Label(frm, text="備註:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        remark_entry = ttk.Entry(frm)
        remark_entry.grid(row=2, column=1)

        try:
            data, _ = get_overview_by_warehouse()
            part_combo["values"] = [r["part_no"] for r in data]
        except Exception as e:
            messagebox.showerror("錯誤", f"載入料號失敗：{e}")

        def do_stockin():
            part_no = part_combo.get()
            try:
                qty = int(qty_entry.get())
                remark = remark_entry.get()
                result = add_stock(part_no, qty, current_user, remark)
                messagebox.showinfo("成功", f"已入庫 {qty} 件\n料號：{part_no}")
                win.destroy()
                refresh_tables()
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        ttk.Button(win, text="執行入庫", command=do_stockin).pack(pady=10)

    def open_transfer_window():
        win = tk.Toplevel()
        win.title("🔁 轉倉操作")
        win.geometry("340x280")

        frm = ttk.Frame(win)
        frm.pack(pady=10)

        tk.Label(frm, text="料號:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        part_combo = ttk.Combobox(frm, values=[], state="readonly", width=22)
        part_combo.grid(row=0, column=1)

        tk.Label(frm, text="來源倉:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        from_wh = ttk.Combobox(frm, values=CORE_WAREHOUSES, state="readonly")
        from_wh.grid(row=1, column=1)

        tk.Label(frm, text="目的倉:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        to_wh = ttk.Combobox(frm, values=CORE_WAREHOUSES, state="readonly")
        to_wh.grid(row=2, column=1)

        tk.Label(frm, text="數量:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        qty_entry = ttk.Entry(frm)
        qty_entry.grid(row=3, column=1)

        tk.Label(frm, text="備註:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        remark_entry = ttk.Entry(frm)
        remark_entry.grid(row=4, column=1)

        try:
            data, _ = get_overview_by_warehouse()
            part_combo["values"] = [r["part_no"] for r in data]
        except Exception as e:
            messagebox.showerror("錯誤", f"載入料號失敗：{e}")

        def do_transfer():
            try:
                part = part_combo.get()
                qty = int(qty_entry.get())
                result = transfer_stock(part, from_wh.get(), to_wh.get(), qty, current_user, remark_entry.get())
                messagebox.showinfo("成功", f"已完成轉倉\n料號: {part}\n數量: {qty}")
                win.destroy()
                refresh_tables()
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        ttk.Button(win, text="執行轉倉", command=do_transfer).pack(pady=10)

    def open_consume_window():
        win = tk.Toplevel()
        win.title("⚙️ 消耗操作")
        win.geometry("340x300")

        frm = ttk.Frame(win)
        frm.pack(pady=10)

        tk.Label(frm, text="料號:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        part_combo = ttk.Combobox(frm, values=[], state="readonly", width=22)
        part_combo.grid(row=0, column=1)

        tk.Label(frm, text="來源倉:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        from_wh = ttk.Combobox(frm, values=CORE_WAREHOUSES, state="readonly")
        from_wh.grid(row=1, column=1)

        tk.Label(frm, text="數量:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        qty_entry = ttk.Entry(frm)
        qty_entry.grid(row=2, column=1)

        tk.Label(frm, text="產線:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        line_entry = ttk.Entry(frm)
        line_entry.grid(row=3, column=1)

        tk.Label(frm, text="用途:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        purpose_entry = ttk.Entry(frm)
        purpose_entry.grid(row=4, column=1)

        try:
            data, _ = get_overview_by_warehouse()
            part_combo["values"] = [r["part_no"] for r in data]
        except Exception as e:
            messagebox.showerror("錯誤", f"載入料號失敗：{e}")

        def do_consume():
            try:
                result = consume_stock(
                    part_no=part_combo.get(),
                    qty=int(qty_entry.get()),
                    user=current_user,
                    line=line_entry.get(),
                    purpose=purpose_entry.get(),
                    from_wh=from_wh.get(),
                    move_to_consumed=True
                )
                messagebox.showinfo("成功", f"已完成消耗\n料號: {result['fixture_id']}\n數量: {result['consumed']}")
                win.destroy()
                refresh_tables()
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        ttk.Button(win, text="執行消耗", command=do_consume).pack(pady=10)

    for wh in CORE_WAREHOUSES:
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=wh)

        tree = ttk.Treeview(frame, columns=("料號", "名稱", "分類", "規格", "數量"), show="headings")
        for col in ("料號", "名稱", "分類", "規格", "數量"):
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True)
        tree_views[wh] = tree

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=4)
        ttk.Button(btn_frame, text="🔁 轉倉操作", command=open_transfer_window).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="⏫ 入庫操作", command=open_stockin_window).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="⚙️ 消耗操作", command=open_consume_window).pack(side="left", padx=4)

    refresh_tables()
    btn_search.config(command=refresh_tables)
    return root

def build_fixture_bom_tab(parent, current_user, db_name):
    frame = ttk.Frame(parent)
    ttk.Label(frame, text="測試BOM分頁（開發中）", font=("Arial", 12)).pack(pady=10)
    return frame