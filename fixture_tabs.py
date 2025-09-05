import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from fixture_helper import (
    insert_fixture,
    delete_fixture,
    add_stock,
    transfer_stock,
    get_overview_by_warehouse,
    update_safety_stock,
    CORE_WAREHOUSES
)

EXCHANGE_RATE = 30.375

CATEGORIES = [
    "電腦設備類",
    "載具類",
    "治具類",
    "板子類",
    "主機類",
    "其它類",
    "線材類",
    "卡片類",
    "電供類",
    "消耗類"
]

def build_fixture_tab(parent, current_user: str = None, db_name: str = None):
    frames = {}
    trees = {}
    total_labels = {}
    total_price_labels = {}

    notebook = ttk.Notebook(parent)
    notebook.pack(fill="both", expand=True)

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

    ttk.Label(control_frame, text="治具單價 (USD):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    entry_price = ttk.Entry(control_frame)
    entry_price.grid(row=4, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="安庫量:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
    entry_safety_stock = ttk.Entry(control_frame)
    entry_safety_stock.grid(row=5, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="入庫量:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
    entry_qty = ttk.Entry(control_frame)
    entry_qty.grid(row=6, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(control_frame, text="入庫倉別:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
    combo_wh = ttk.Combobox(control_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_wh.grid(row=7, column=1, padx=5, pady=5, sticky="w")
    combo_wh.current(0)
        
    def on_add_fixture():
        part = entry_part.get().strip()
        name = entry_name.get().strip()
        spec = entry_spec.get().strip()
        cat = combo_cat.get().strip()
        price = entry_price.get().strip()
        safety = entry_safety_stock.get().strip()

        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        if not name or not spec or not cat or not price or not safety:
            messagebox.showerror("錯誤", "治具品名、規格、類群、單價、安庫量不可為空")
            return
        try:
            price_val = round(float(price), 3)
        except:
            messagebox.showerror("錯誤", "治具單價必須是數字 (可小數三位)")
            return
        try:
            safety_val = int(safety)
        except:
            messagebox.showerror("錯誤", "安庫量必須是整數")
            return

        try:
            insert_fixture(part, name, spec, cat, price_val, safety_val)
            messagebox.showinfo("完成", f"已新增治具料號 {part}")
            for wh in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh], wh, total_labels[wh], total_price_labels[wh])
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
                refresh_fixture_tree(trees[wh], wh, total_labels[wh], total_price_labels[wh])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_update_safety():
        part = entry_part.get().strip()
        safety = entry_safety_stock.get().strip()
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤", "治具料號必須為 8 或 12 碼數字")
            return
        try:
            safety_val = int(safety)
        except:
            messagebox.showerror("錯誤", "安庫量必須是整數")
            return
        try:
            update_safety_stock(part, safety_val)
            messagebox.showinfo("完成", f"{part} 的安庫量已更新為 {safety_val}")
            for wh in CORE_WAREHOUSES:
                refresh_fixture_tree(trees[wh], wh, total_labels[wh], total_price_labels[wh])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_add_stock():
        part = entry_part.get().strip()
        name = entry_name.get().strip()
        spec = entry_spec.get().strip()
        cat = combo_cat.get().strip()
        wh = combo_wh.get()
        try:
            qty = int(entry_qty.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "在庫量必須是整數")
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
                refresh_fixture_tree(trees[wh2], wh2, total_labels[wh2], total_price_labels[wh2])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

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
            messagebox.showerror("錯誤", "調撥量必須是整數")
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
                refresh_fixture_tree(trees[wh2], wh2, total_labels[wh2], total_price_labels[wh2])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
    def on_export_excel():
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="匯出治具存量檢查"
        )
        if not file_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "治具存量"

        headers = ["治具料號", "治具品名", "治具規格", "治具類群", "單價 USD", "單價 NTD", "安庫量"]
        headers += CORE_WAREHOUSES
        headers += ["總在庫量", "缺少數量", "總價 NTD"]
        ws.append(headers)

        all_data = {}
        for wh in CORE_WAREHOUSES:
            rows = get_overview_by_warehouse(wh)
            for part_no, name, spec, category, unit_price, safety_stock, qty in rows:
                if part_no not in all_data:
                    all_data[part_no] = {
                        "name": name,
                        "spec": spec,
                        "category": category,
                        "unit_price": unit_price or 0,
                        "safety_stock": safety_stock or 0,
                        "warehouses": {w: 0 for w in CORE_WAREHOUSES}
                    }
                all_data[part_no]["warehouses"][wh] = qty or 0

        total_purchase_cost = 0.0
        for part_no, data in all_data.items():
            total_qty = sum(data["warehouses"].values())
            unit_price_ntd = round(data["unit_price"] * EXCHANGE_RATE, 3)
            shortage = max(data["safety_stock"] - total_qty, 0)
            shortage_cost = round(shortage * unit_price_ntd, 3)
            total_purchase_cost += shortage_cost
            row = [
                part_no,
                data["name"],
                data["spec"],
                data["category"],
                data["unit_price"],
                unit_price_ntd,
                data["safety_stock"],
            ]
            row += [data["warehouses"][wh] for wh in CORE_WAREHOUSES]
            row += [total_qty, shortage, shortage_cost]
            ws.append(row)

        ws.append([])
        ws.append(["預估採購總價 (NTD)", "", "", "", "", "", "", "", "", "", "", "", "", "", "", round(total_purchase_cost, 3)])

        wb.save(file_path)
        messagebox.showinfo("完成", f"已匯出 Excel：\n{file_path}")

    ttk.Button(control_frame, text="新增治具", command=on_add_fixture).grid(row=0, column=2, padx=5, pady=5)
    ttk.Button(control_frame, text="刪除治具", command=on_delete_fixture).grid(row=0, column=3, padx=5, pady=5)
    ttk.Button(control_frame, text="⚙ 修改安庫", command=on_update_safety).grid(row=5, column=2, padx=5, pady=5)
    ttk.Button(control_frame, text="⏫ 入庫", command=on_add_stock).grid(row=6, column=2, padx=5, pady=5)
    ttk.Button(control_frame, text="📤 匯出 Excel", command=on_export_excel).grid(row=9, column=0, padx=5, pady=5, sticky="w")

    transfer_frame = ttk.LabelFrame(control_frame, text="調撥")
    transfer_frame.grid(row=8, column=0, columnspan=4, pady=10, sticky="ew")

    ttk.Label(transfer_frame, text="來源倉別:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    combo_from = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_from.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(transfer_frame, text="目標倉別:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
    combo_to = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly")
    combo_to.grid(row=0, column=3, padx=5, pady=5, sticky="w")

    ttk.Label(transfer_frame, text="調撥量:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
    entry_transfer_qty = ttk.Entry(transfer_frame, width=8)
    entry_transfer_qty.grid(row=0, column=5, padx=5, pady=5, sticky="w")

    ttk.Button(transfer_frame, text="🔁 調撥", command=on_transfer).grid(row=0, column=6, padx=5, pady=5)

    for wh in CORE_WAREHOUSES:
        frame = ttk.Frame(notebook)
        frames[wh] = frame
        notebook.add(frame, text=wh)

        topbar = ttk.Frame(frame)
        topbar.pack(fill="x")
        total_label = ttk.Label(topbar, text="合計在庫量: 0")
        total_label.pack(side="right", padx=8, pady=4)
        total_labels[wh] = total_label

        total_price_label = ttk.Label(topbar, text="倉別總價: 0 NTD")
        total_price_label.pack(side="right", padx=8, pady=4)
        total_price_labels[wh] = total_price_label

        columns = ("part_no", "name", "spec", "category", "unit_price_usd", "unit_price_ntd", "total_price_ntd", "safety_stock", "qty")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        trees[wh] = tree

        tree.heading("part_no", text="治具料號")
        tree.heading("name", text="治具品名")
        tree.heading("spec", text="治具規格")
        tree.heading("category", text="治具類群")
        tree.heading("unit_price_usd", text="單價 USD")
        tree.heading("unit_price_ntd", text="單價 NTD")
        tree.heading("total_price_ntd", text="總價 NTD")
        tree.heading("safety_stock", text="安庫量")
        tree.heading("qty", text="在庫量")

        tree.column("part_no", width=70, anchor="center")
        tree.column("name", width=300, anchor="center")
        tree.column("spec", width=300, anchor="center")
        tree.column("category", width=50, anchor="center")
        tree.column("unit_price_usd", width=50, anchor="center")
        tree.column("unit_price_ntd", width=50, anchor="center")
        tree.column("total_price_ntd", width=50, anchor="center")
        tree.column("safety_stock", width=50, anchor="center")
        tree.column("qty", width=50, anchor="center")

        tree.pack(fill="both", expand=True)

        refresh_fixture_tree(tree, wh, total_labels[wh], total_price_labels[wh])

        def on_tree_double_click(event, wh=wh, tree=tree):
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
            entry_price.delete(0, tk.END)
            entry_price.insert(0, vals[4])
            entry_safety_stock.delete(0, tk.END)
            entry_safety_stock.insert(0, vals[7])
            combo_from.set(wh)
            combo_wh.set(wh)

        tree.bind("<Double-1>", on_tree_double_click)

def refresh_fixture_tree(tree, warehouse, total_label=None, total_price_label=None):
    tree.delete(*tree.get_children())
    rows = get_overview_by_warehouse(warehouse)

    total_qty = 0
    total_price = 0.0
    for part_no, name, spec, category, unit_price, safety_stock, qty in rows:
        unit_price = unit_price or 0
        qty = qty or 0
        unit_price_ntd = round(unit_price * EXCHANGE_RATE, 3)
        total_price_item = round(unit_price_ntd * qty, 3)
        tree.insert("", "end", values=(part_no, name, spec, category, unit_price, unit_price_ntd, total_price_item, safety_stock, qty))
        total_qty += qty
        total_price += total_price_item

    if total_label is not None:
        total_label.config(text=f"合計在庫量: {total_qty}")
    if total_price_label is not None:
        total_price_label.config(text=f"倉別總價: {round(total_price, 3)} NTD")