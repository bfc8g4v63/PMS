# Fixture Bom Tab.py
import tkinter as tk
from tkinter import ttk, messagebox
from fixture_helper import (
    insert_fixture,
    add_stock,
    transfer_stock,
    consume_stock,
    get_overview_by_warehouse,
    get_bom_by_part,
    add_bom_item,
    delete_bom_item,
    CORE_WAREHOUSES
)

WAREHOUSES_ALL = CORE_WAREHOUSES + ["消耗"]

def build_fixture_bom_tab(parent: tk.Widget, current_user: str, db_name: str = None) -> tk.Frame:
    root = tk.Frame(parent)
    root.pack(fill="both", expand=True)

    top = tk.Frame(root)
    top.pack(fill="x", padx=8, pady=6)

    part_var = tk.StringVar()
    tk.Label(top, text="料號(8/12碼)").grid(row=0, column=0, sticky="w")
    tk.Entry(top, textvariable=part_var, width=24).grid(row=0, column=1, padx=(6, 12))
    btn_query = tk.Button(top, text="查詢", width=8)
    btn_query.grid(row=0, column=2, padx=4)
    btn_reset = tk.Button(top, text="重置", width=8)
    btn_reset.grid(row=0, column=3, padx=4)

    mid = ttk.LabelFrame(root, text="BOM 清單")
    mid.pack(fill="both", expand=True, padx=8, pady=6)

    columns = ("bom_id", "fixture_name", "quantity", "created_at")
    tree = ttk.Treeview(mid, columns=columns, show="headings")
    for col, text in zip(columns, ["ID", "治具名稱", "數量", "建立時間"]):
        tree.heading(col, text=text)
    tree.column("bom_id", width=60, anchor="center")
    tree.column("fixture_name", width=180, anchor="w")
    tree.column("quantity", width=80, anchor="e")
    tree.column("created_at", width=140, anchor="center")
    vsb = ttk.Scrollbar(mid, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    bottom = ttk.LabelFrame(root, text="新增 BOM 項目")
    bottom.pack(fill="x", padx=8, pady=6)

    bom_name_var = tk.StringVar()
    bom_qty_var = tk.StringVar()
    tk.Label(bottom, text="治具名稱").grid(row=0, column=0, sticky="e", padx=4, pady=3)
    tk.Entry(bottom, textvariable=bom_name_var, width=28).grid(row=0, column=1, padx=4, pady=3)
    tk.Label(bottom, text="數量").grid(row=0, column=2, sticky="e", padx=4, pady=3)
    tk.Entry(bottom, textvariable=bom_qty_var, width=8).grid(row=0, column=3, padx=4, pady=3)

    btn_add = tk.Button(bottom, text="新增", width=10)
    btn_add.grid(row=0, column=4, padx=6)
    btn_del = tk.Button(bottom, text="刪除選取", width=10)
    btn_del.grid(row=0, column=5, padx=6)

    def refresh_tree():
        try:
            part = part_var.get().strip()
            if not part:
                raise ValueError("請輸入料號後再查詢")
            for i in tree.get_children():
                tree.delete(i)
            rows = get_bom_by_part(part)
            for row in rows:
                tree.insert("", "end", values=(row["bom_id"], row["fixture_name"], row["quantity"], row["created_at"]))
        except Exception as e:
            messagebox.showerror("查詢失敗", str(e))

    def on_add():
        try:
            part = part_var.get().strip()
            name = bom_name_var.get().strip()
            qty = int(bom_qty_var.get().strip())
            if not part or not name or qty <= 0:
                raise ValueError("請輸入完整資訊並確認數量為正整數")
            add_bom_item(part, name, qty)
            bom_name_var.set("")
            bom_qty_var.set("")
            refresh_tree()
        except Exception as e:
            messagebox.showerror("新增失敗", str(e))

    def on_delete():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選取要刪除的項目")
            return
        for item in sel:
            bom_id = tree.item(item, "values")[0]
            delete_bom_item(bom_id)
        refresh_tree()

    def on_reset():
        part_var.set("")
        for i in tree.get_children():
            tree.delete(i)

    btn_query.configure(command=refresh_tree)
    btn_reset.configure(command=on_reset)
    btn_add.configure(command=on_add)
    btn_del.configure(command=on_delete)

    return root