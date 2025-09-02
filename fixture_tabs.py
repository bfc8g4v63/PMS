import tkinter as tk
from tkinter import ttk, messagebox
from fixture_helper import (
    insert_fixture,
    delete_fixture,
    add_stock,
    transfer_stock,
    get_overview_by_warehouse,
    CORE_WAREHOUSES
)

CATEGORIES = [
    "電腦/設備類",
    "載具類",
    "治具類",
    "裸板類",
    "主機類",
    "測具/板子類",
    "線材類",
    "卡片類",
    "電供類",
    "消耗類"
]

def build_fixture_tab(parent, current_user: str = None, db_name: str = None):
    frames = {}
    trees = {}
    total_labels = {}

    notebook = ttk.Notebook(parent)
    notebook.pack(fill="both", expand=True)

    for wh in CORE_WAREHOUSES:
        frame = ttk.Frame(notebook)
        frames[wh] = frame
        notebook.add(frame, text=wh)

        topbar = ttk.Frame(frame)
        topbar.pack(fill="x")
        total_label = ttk.Label(topbar, text="合計: 0")
        total_label.pack(side="right", padx=8, pady=4)
        total_labels[wh] = total_label

        columns = ("part_no", "name", "spec", "category", "qty")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        trees[wh] = tree

        tree.heading("part_no", text="治具料號")
        tree.heading("name", text="治具品名")
        tree.heading("spec", text="治具規格")
        tree.heading("category", text="治具類群")
        tree.heading("qty", text="數量")

        tree.column("part_no", width=20, anchor="center")
        tree.column("name", width=350, anchor="center")
        tree.column("spec", width=350, anchor="center")
        tree.column("category", width=20, anchor="center")
        tree.column("qty", width=20, anchor="center")

        tree.pack(fill="both", expand=True)

        refresh_fixture_tree(tree, wh, total_label)

        def on_tree_double_click(event, tree=tree):
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values")
            entry_part.delete(0, tk.END)
            entry_part.insert(0, vals[0])
            entry_name.delete(0, tk.END)
            entry_name.insert(0, vals[1])
            entry_spec.delete(0, tk.END)
            entry_spec.insert(0, vals[2])
            combo_cat.set(vals[3])

        tree.bind("<Double-1>", on_tree_double_click)

    control_frame = ttk.Frame(parent)
    control_frame.pack(fill="x", pady=10)

    ttk.Label(control_frame, text="治具料號:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_part = ttk.Entry(control_frame)
    entry_part.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="治具品名:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    entry_name = ttk.Entry(control_frame)
    entry_name.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="治具規格:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
    entry_spec = ttk.Entry(control_frame)
    entry_spec.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="治具類群:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
    combo_cat = ttk.Combobox(control_frame, values=CATEGORIES, state="readonly")
    combo_cat.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    def on_add_fixture():
        part = entry_part.get().strip()
        name = entry_name.get().strip()
        spec = entry_spec.get().strip()
        cat = combo_cat.get().strip()
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        if not name or not spec or not cat:
            messagebox.showerror("錯誤", "治具品名、治具規格、治具類群不可為空")
            return
        try:
            insert_fixture(part, name, spec, cat, 0)
            messagebox.showinfo("完成", f"已新增治具料號 {part}")
            for wh in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh], wh, total_labels[wh])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_delete_fixture():
        part = entry_part.get().strip()
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        confirm = messagebox.askyesno("確認刪除", f"確定要刪除治具料號 {part}？\n此動作會移除所有倉別的存量與 BOM。")
        if not confirm:
            return
        try:
            delete_fixture(part)
            messagebox.showinfo("完成", f"已刪除治具料號 {part}")
            for wh in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh], wh, total_labels[wh])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    ttk.Button(control_frame, text="新增治具", command=on_add_fixture).grid(row=0, column=2, padx=5, pady=5)
    ttk.Button(control_frame, text="刪除治具", command=on_delete_fixture).grid(row=0, column=3, padx=5, pady=5)

    ttk.Label(control_frame, text="入庫倉別:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    combo_wh = ttk.Combobox(control_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_wh.grid(row=4, column=1, padx=5, pady=5, sticky="w")
    combo_wh.current(0)

    ttk.Label(control_frame, text="入庫數量:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
    entry_qty = ttk.Entry(control_frame)
    entry_qty.grid(row=5, column=1, padx=5, pady=5, sticky="w")

    def on_add_stock():
        part = entry_part.get().strip()
        name = entry_name.get().strip()
        spec = entry_spec.get().strip()
        cat = combo_cat.get().strip()
        wh = combo_wh.get()
        try:
            qty = int(entry_qty.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "數量必須是整數")
            return
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        if not name or not spec or not cat:
            messagebox.showerror("錯誤", "治具品名、治具規格、治具類群不可為空")
            return
        try:
            add_stock(part, qty, wh)
            messagebox.showinfo("完成", f"{part} 已入庫 {qty} 至 {wh}")
            for wh2 in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh2], wh2, total_labels[wh2])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    ttk.Button(control_frame, text="⏫ 入庫", command=on_add_stock).grid(row=5, column=2, padx=5, pady=5)

    transfer_frame = ttk.LabelFrame(control_frame, text="調撥")
    transfer_frame.grid(row=6, column=0, columnspan=4, pady=10, sticky="ew")

    ttk.Label(transfer_frame, text="來源倉別:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    combo_from = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_from.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(transfer_frame, text="目標倉別:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
    combo_to = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_to.grid(row=0, column=3, padx=5, pady=5, sticky="w")

    ttk.Label(transfer_frame, text="數量:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
    entry_transfer_qty = ttk.Entry(transfer_frame, width=8)
    entry_transfer_qty.grid(row=0, column=5, padx=5, pady=5, sticky="w")

    def on_transfer():
        part = entry_part.get().strip()
        name = entry_name.get().strip()
        spec = entry_spec.get().strip()
        cat = combo_cat.get().strip()
        from_wh = combo_from.get()
        to_wh = combo_to.get()
        try:
            qty = int(entry_transfer_qty.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "數量必須是整數")
            return
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        if not name or not spec or not cat:
            messagebox.showerror("錯誤", "治具品名、治具規格、治具類群不可為空")
            return
        if from_wh == to_wh:
            messagebox.showerror("錯誤", "來源與目標倉別不可相同")
            return
        try:
            transfer_stock(part, qty, from_wh, to_wh)
            messagebox.showinfo("完成", f"{part} 已調撥 {qty} 從 {from_wh} 到 {to_wh}")
            for wh2 in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh2], wh2, total_labels[wh2])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    ttk.Button(transfer_frame, text="🔁 調撥", command=on_transfer).grid(row=0, column=6, padx=5, pady=5)

def refresh_fixture_tree(tree, warehouse, total_label=None):
    tree.delete(*tree.get_children())
    rows = get_overview_by_warehouse(warehouse)

    total = 0
    for part_no, name, spec, category, safety_stock, qty in rows:
        tree.insert("", "end", values=(part_no, name, spec, category, qty))
        try:
            total += int(qty) if qty is not None else 0
        except:
            pass

    if total_label is not None:
        total_label.config(text=f"合計: {total}")