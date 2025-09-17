#$ fixture_tabs.py
#% GUI 分頁「治具管理」
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from fixture_helper import (
    insert_fixture,
    delete_fixture,
    add_stock,
    transfer_stock,
    consume_stock,
    update_fixture,
    get_overview_by_warehouse,
    get_fixture_by_part_no,
    validate_location,
    CORE_WAREHOUSES
)

EXCHANGE_RATE = 30.375
CATEGORIES = ["電腦設備類","載具類","治具類","板子類","主機類","其它類","線材類","卡片類","電供類","消耗類"]

WAREHOUSES = CORE_WAREHOUSES + ["消耗"]

def storage_location(raw: str) -> str:
    s = (raw or "").strip()
    parts = s.split("-")
    if len(parts) != 3:
        raise ValueError("儲位格式必須為 車-層-位置，例如 1-1-1")
    try:
        car = int(parts[0])
        layer = int(parts[1])
        pos = int(parts[2])
    except:
        raise ValueError("儲位格式必須為數字，例如 1-1-1")
    if not (1 <= car <= 9):
        raise ValueError("車號必須介於 1~9")
    if not (1 <= layer <= 4):
        raise ValueError("層號必須介於 1~4")
    if not (1 <= pos <= 50):
        raise ValueError("位置必須介於 1~50")
    return f"{car}-{layer}-{pos}"

def build_fixture_tab(parent, current_user: str = None):
    frames, trees, total_labels, total_price_labels = {}, {}, {}, {}

    notebook = ttk.Notebook(parent)
    notebook.pack(fill="both", expand=True)

    form = ttk.LabelFrame(parent, text="治具操作")
    form.pack(fill="x", padx=5, pady=5)

    labels = ["治具料號","治具品名","治具規格","治具類群","治具單價(USD)","安全庫存","儲位","入庫數量"]
    entries = {}
    for i, lab in enumerate(labels):
        ttk.Label(form, text=lab+":").grid(row=i, column=0, sticky="e", padx=3, pady=2)
        ent = ttk.Entry(form, width=20)
        ent.grid(row=i, column=1, sticky="w", padx=3, pady=2)
        entries[lab] = ent

    combo_cat = ttk.Combobox(form, values=CATEGORIES, state="readonly", width=18)
    combo_cat.grid(row=3, column=1, sticky="w", padx=3, pady=2)
    entries["治具類群"] = combo_cat

    def on_add_fixture():
        part = entries["治具料號"].get().strip()
        name = entries["治具品名"].get().strip()
        spec = entries["治具規格"].get().strip()
        part_group = combo_cat.get().strip()
        usd  = entries["治具單價(USD)"].get().strip()
        safety = entries["安全庫存"].get().strip()
        loc_raw = entries["儲位"].get().strip()

        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤","治具料號必須為 8 或 12 碼數字"); return
        if not name or not spec or not part_group:
            messagebox.showerror("錯誤","治具品名/規格/類群不可為空"); return
        try:
            usd_val   = round(float(usd or 0), 3)
            ntd_val   = round(usd_val * EXCHANGE_RATE, 3)
            safety_val = int(safety or 0)
        except:
            messagebox.showerror("錯誤","單價或安全庫存格式錯誤"); return
        try:
            loc_fmt = validate_location(part, loc_raw)
        except Exception as e:
            messagebox.showerror("錯誤", str(e)); return

        try:
            insert_fixture(
                part, name, spec, part_group,
                ntd_val, safety_val, loc_fmt,
                unit_price_usd=usd_val
            )
            messagebox.showinfo("完成", f"治具 {part} 已新增")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_delete_fixture():
        part = entries["治具料號"].get().strip()
        if not part:
            return
        if not messagebox.askyesno("確認", f"確定刪除 {part}？將移除所有倉別存量與 BOM 及相關紀錄"):
            return
        try:
            delete_fixture(part)
            messagebox.showinfo("完成", f"{part} 已刪除")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_add_stock():
        part = entries["治具料號"].get().strip()
        qty_str = entries["入庫數量"].get().strip()
        try:
            qty = int(qty_str)
        except:
            messagebox.showerror("錯誤","入庫量必須是整數"); return
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤","治具料號必須為 8 或 12 碼數字"); return
        current_tab = notebook.tab(notebook.select(), "text")
        wh = current_tab if current_tab != "消耗" else "虹堡"
        try:
            add_stock(part, qty, wh, user=current_user or "", remark="入庫作業")
            messagebox.showinfo("完成", f"{part} 已入庫 {qty} 至 {wh}")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_transfer():
        part = entries["治具料號"].get().strip()
        qty_str = entries["入庫數量"].get().strip()
        from_wh, to_wh = combo_from.get(), combo_to.get()
        try:
            qty = int(qty_str)
        except:
            messagebox.showerror("錯誤","調撥量必須是整數"); return
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤","治具料號必須為 8 或 12 碼數字"); return
        if from_wh == to_wh:
            messagebox.showerror("錯誤","來源與目標倉別相同"); return
        try:
            transfer_stock(part, qty, from_wh, to_wh, user=current_user or "", remark="調撥作業")
            messagebox.showinfo("完成", f"{part} 已調撥 {qty} 從 {from_wh} 到 {to_wh}")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_consume():
        part = entries["治具料號"].get().strip()
        qty_str = entries["入庫數量"].get().strip()
        current_tab = notebook.tab(notebook.select(), "text")
        wh = current_tab if current_tab != "消耗" else "虹堡"
        try:
            qty = int(qty_str)
        except:
            messagebox.showerror("錯誤","消耗量必須是整數"); return
        if not (part.isdigit() and len(part) in (8, 12)):
            messagebox.showerror("錯誤","治具料號必須為 8 或 12 碼數字"); return
        try:
            consume_stock(part, qty, wh, user=current_user or "", line="", purpose="生產消耗")
            messagebox.showinfo("完成", f"{part} 已自 {wh} 消耗 {qty}")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_update_fixture():
        part = entries["治具料號"].get().strip()
        name = entries["治具品名"].get().strip()
        spec = entries["治具規格"].get().strip()
        part_group = combo_cat.get().strip()
        usd  = entries["治具單價(USD)"].get().strip()
        loc_raw  = entries["儲位"].get().strip()
        safety = entries["安全庫存"].get().strip()
        try:
            usd_val   = round(float(usd or 0), 3)
            ntd_val   = round(usd_val * EXCHANGE_RATE, 3)
            safety_val = int(safety or 0)
        except:
            messagebox.showerror("錯誤","單價或安全庫存格式錯誤"); return
        try:
            loc_fmt = validate_location(part, loc_raw)
        except Exception as e:
            messagebox.showerror("錯誤", str(e)); return
        try:
            update_fixture(
                part,
                part_name=name,
                part_spec=spec,
                part_group=part_group,
                unit_price_ntd=ntd_val,
                unit_price_usd=usd_val,
                safety_stock=safety_val,
                storage_location=loc_fmt
            )
            messagebox.showinfo("完成", f"{part} 的資料已更新")
            refresh_all()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_export_excel():
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files","*.xlsx")])
        if not file: return
        wb = openpyxl.Workbook()
        for wh in WAREHOUSES:
            ws = wb.create_sheet(title=wh)
            headers = [
                "治具料號","治具品名","治具規格","治具類群",
                "單價USD","單價NTD","總價NTD",
                "儲位","安全庫存","可用數量",
                "預估請購量","預估請購金額"
            ]
            ws.append(headers)
            rows = get_overview_by_warehouse(wh if wh != "消耗" else "虹堡")

            total_estimate_cost = 0
            for (part_no, name, spec, part_group, unit_price_usd, unit_price_ntd, safety_stock, location, qty) in rows:
                unit_price_usd = round((unit_price_usd or 0), 3)
                total_price_item = round((unit_price_ntd or 0) * (qty or 0), 3)

                est_qty = max((safety_stock or 0) - (qty or 0), 0)
                est_cost = round(est_qty * (unit_price_ntd or 0), 3)
                total_estimate_cost += est_cost

                ws.append([
                    part_no,name,spec,part_group,
                    unit_price_usd,unit_price_ntd,total_price_item,
                    location,safety_stock,qty,
                    est_qty,est_cost
                ])

            ws.append([])
            ws.append(["", "", "", "", "", "", "", "", "", "預估請購總金額", total_estimate_cost])

        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        wb.save(file)
        messagebox.showinfo("完成", f"已匯出 Excel：{file}")

    ttk.Button(form, text="新增治具", command=on_add_fixture).grid(row=0, column=2, padx=5)
    ttk.Button(form, text="刪除治具", command=on_delete_fixture).grid(row=1, column=2, padx=5)
    ttk.Button(form, text="⚙修改資料", command=on_update_fixture).grid(row=6, column=2, padx=5)
    ttk.Button(form, text="⏫入庫", command=on_add_stock).grid(row=7, column=2, padx=5)
    ttk.Button(form, text="⚙消耗", command=on_consume).grid(row=9, column=2, padx=5)
    ttk.Button(form, text="⇩匯出Excel", command=on_export_excel).grid(row=10, column=2, padx=5)

    transfer_frame = ttk.LabelFrame(form, text="調撥")
    transfer_frame.grid(row=8, column=0, columnspan=2, sticky="ew", pady=5)
    ttk.Label(transfer_frame, text="來源:").grid(row=0, column=0, padx=3)
    ttk.Label(transfer_frame, text="目標:").grid(row=0, column=2, padx=3)
    global combo_from, combo_to
    combo_from = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly", width=12)
    combo_to   = ttk.Combobox(transfer_frame, values=CORE_WAREHOUSES, state="readonly", width=12)
    combo_from.grid(row=0, column=1, padx=3)
    combo_to.grid(row=0, column=3, padx=3)

    for wh in WAREHOUSES:
        frame = ttk.Frame(notebook); frames[wh] = frame; notebook.add(frame, text=wh)
        top = ttk.Frame(frame); top.pack(fill="x")
        total_labels[wh] = ttk.Label(top, text="合計可用數量: 0"); total_labels[wh].pack(side="right", padx=6)
        total_price_labels[wh] = ttk.Label(top, text="倉別總價: 0 NTD"); total_price_labels[wh].pack(side="right", padx=6)

        cols = ("part_no","name","spec","part_group","unit_price_usd","unit_price_ntd","total_price_ntd","safety_stock","location","qty")
        headers = ["治具料號","治具品名","治具規格","治具類群","單價USD","單價NTD","總價NTD","安全庫存","儲位","可用數量"]

        tree = ttk.Treeview(frame, columns=cols, show="headings"); trees[wh] = tree
        for c, t in zip(cols, headers):
            tree.heading(c, text=t)
            tree.column(c, width=90, anchor="center")
        tree.column("name", width=240); tree.column("spec", width=240)
        tree.pack(fill="both", expand=True)
        refresh_fixture_tree(tree, wh, total_labels[wh], total_price_labels[wh])

        def on_double(event, tree=tree, wh=wh):
            sel = tree.selection()
            if not sel: return
            vals = tree.item(sel[0], "values")
            entries["治具料號"].delete(0, tk.END); entries["治具料號"].insert(0, vals[0])
            entries["治具品名"].delete(0, tk.END); entries["治具品名"].insert(0, vals[1])
            entries["治具規格"].delete(0, tk.END); entries["治具規格"].insert(0, vals[2])
            combo_cat.set(vals[3])
            entries["治具單價(USD)"].delete(0, tk.END); entries["治具單價(USD)"].insert(0, vals[4])
            entries["安全庫存"].delete(0, tk.END); entries["安全庫存"].insert(0, vals[7])
            entries["儲位"].delete(0, tk.END); entries["儲位"].insert(0, vals[8])
            combo_from.set(wh if wh != "消耗" else "虹堡")

        tree.bind("<Double-1>", on_double)

    def refresh_all():
        for wh in WAREHOUSES:
            refresh_fixture_tree(trees[wh], wh, total_labels[wh], total_price_labels[wh])

def refresh_fixture_tree(tree, warehouse, total_label=None, total_price_label=None):
    tree.delete(*tree.get_children())
    wh_for_query = warehouse if warehouse != "消耗" else "虹堡"
    rows = get_overview_by_warehouse(wh_for_query)

    total_qty = 0
    total_price = 0.0
    for (part_no, name, spec, part_group, unit_price_usd, unit_price_ntd, safety_stock, location, qty) in rows:
        unit_price_ntd = unit_price_ntd or 0
        qty = qty or 0
        total_price_item = round(unit_price_ntd * qty, 3)
        show_location = location if warehouse == "虹堡" else ""
        tree.insert(
            "",
            "end",
            values=(
                part_no,
                name,
                spec,
                part_group,
                unit_price_usd,
                unit_price_ntd,
                total_price_item,
                safety_stock,
                show_location,
                qty,
            ),
        )
        total_qty += qty
        total_price += total_price_item

    if total_label is not None:
        total_label.config(text=f"合計可用數量: {total_qty}")
    if total_price_label is not None:
        total_price_label.config(text=f"倉別總價: {round(total_price, 3)} NTD")