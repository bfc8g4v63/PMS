# fixture_bom_tab.py
# GUI 分頁「治具 BOM 管理」
import tkinter as tk
from tkinter import ttk, messagebox
from fixture_helper import (
    get_bom_by_part,
    add_bom_item,
    delete_bom_item,
    get_fixture_by_part_no
)

def build_fixture_bom_tab(parent, current_user=None):
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)

    form = ttk.LabelFrame(frame, text="治具 BOM 操作")
    form.pack(fill="x", padx=5, pady=5)

    ttk.Label(form, text="主料號:").grid(row=0, column=0, padx=3, pady=2, sticky="e")
    entry_parent = ttk.Entry(form, width=20)
    entry_parent.grid(row=0, column=1, padx=3, pady=2)

    ttk.Label(form, text="子料號:").grid(row=1, column=0, padx=3, pady=2, sticky="e")
    entry_child = ttk.Entry(form, width=20)
    entry_child.grid(row=1, column=1, padx=3, pady=2)

    ttk.Label(form, text="數量:").grid(row=2, column=0, padx=3, pady=2, sticky="e")
    entry_qty = ttk.Entry(form, width=20)
    entry_qty.grid(row=2, column=1, padx=3, pady=2)

    tree = ttk.Treeview(frame, columns=("id","parent_part_no","child_part_no","bom_qty"), show="headings")
    tree.heading("id", text="ID")
    tree.heading("parent_part_no", text="主料號")
    tree.heading("child_part_no", text="子料號")
    tree.heading("bom_qty", text="數量")

    tree.column("id", width=60, anchor="center")
    tree.column("parent_part_no", width=120, anchor="center")
    tree.column("child_part_no", width=120, anchor="center")
    tree.column("bom_qty", width=80, anchor="center")
    tree.pack(fill="both", expand=True, padx=5, pady=5)

    def refresh_tree():
        tree.delete(*tree.get_children())
        parent_no = entry_parent.get().strip()
        if not parent_no:
            return
        try:
            rows = get_bom_by_part(parent_no)
            for r in rows:
                tree.insert("", "end", values=r[:4])
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_add_bom():
        parent_no = entry_parent.get().strip()
        child_no = entry_child.get().strip()
        qty_str = entry_qty.get().strip()
        if not (parent_no and child_no and qty_str):
            messagebox.showerror("錯誤", "主料號 / 子料號 / 數量 不可為空")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                messagebox.showerror("錯誤", "數量必須大於 0")
                return
        except:
            messagebox.showerror("錯誤", "數量必須是整數")
            return
        if not get_fixture_by_part_no(parent_no):
            messagebox.showerror("錯誤", f"主料號 {parent_no} 不存在於治具清單")
            return
        if not get_fixture_by_part_no(child_no):
            messagebox.showerror("錯誤", f"子料號 {child_no} 不存在於治具清單")
            return
        try:
            add_bom_item(parent_no, child_no, qty)
            messagebox.showinfo("完成", f"{parent_no} 新增 BOM 子料號 {child_no} x{qty}")
            refresh_tree()
            entry_child.delete(0, tk.END)
            entry_qty.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    def on_delete_bom():
        sel = tree.selection()
        if not sel:
            return
        bom_id = tree.item(sel[0], "values")[0]
        if not messagebox.askyesno("確認", f"確定刪除 BOM ID={bom_id}？"):
            return
        try:
            delete_bom_item(int(bom_id))
            messagebox.showinfo("完成", f"BOM ID {bom_id} 已刪除")
            refresh_tree()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    btn_frame = ttk.Frame(form)
    btn_frame.grid(row=0, column=2, rowspan=3, padx=5, pady=2, sticky="ns")
    ttk.Button(btn_frame, text="查詢 BOM", command=refresh_tree).pack(fill="x", pady=2)
    ttk.Button(btn_frame, text="新增子料", command=on_add_bom).pack(fill="x", pady=2)
    ttk.Button(btn_frame, text="刪除子料", command=on_delete_bom).pack(fill="x", pady=2)