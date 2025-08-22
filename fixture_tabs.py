import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from fixture_helper import insert_fixture

def build_fixture_tab(parent, current_user, db_name):
    """建立治具管理分頁（依倉別加總視圖）"""

    form_frame = tk.LabelFrame(parent, text="新增治具")
    form_frame.pack(fill="x", padx=10, pady=5)

    notebook = ttk.Notebook(parent)
    notebook.pack(fill="both", expand=True, padx=10, pady=5)

    tabs = {}
    treeviews = {}
    warehouse_names = ["治具室", "上齊", "睿均", "不良品"]

    for wh in warehouse_names:
        tab = tk.Frame(notebook)
        notebook.add(tab, text=wh)
        tabs[wh] = tab

        tree = ttk.Treeview(tab, columns=("part_no", "name", "spec", "safety_stock", "qty", "status"), show="headings")
        tree.pack(fill="both", expand=True, padx=5, pady=5)

        for col, label in [
            ("part_no", "料號"),
            ("name", "品名"),
            ("spec", "規格"),
            ("safety_stock", "安庫量"),
            ("qty", "數量"),
            ("status", "狀態")
        ]:
            tree.heading(col, text=label)

        treeviews[wh] = tree

    tk.Label(form_frame, text="料號:").grid(row=0, column=0, sticky="e")
    part_no_entry = tk.Entry(form_frame, width=30)
    part_no_entry.grid(row=0, column=1)

    tk.Label(form_frame, text="品名:").grid(row=1, column=0, sticky="e")
    name_entry = tk.Entry(form_frame, width=30)
    name_entry.grid(row=1, column=1)

    tk.Label(form_frame, text="規格:").grid(row=2, column=0, sticky="e")
    spec_entry = tk.Entry(form_frame, width=30)
    spec_entry.grid(row=2, column=1)

    tk.Label(form_frame, text="安庫量:").grid(row=3, column=0, sticky="e")
    safety_entry = tk.Entry(form_frame, width=30)
    safety_entry.grid(row=3, column=1)

    def refresh_treeviews():
        with sqlite3.connect(db_name) as conn:
            cursor = conn.cursor()
            for warehouse, tree in treeviews.items():
                for row in tree.get_children():
                    tree.delete(row)
                cursor.execute("""
                    SELECT part_no, item_name, spec, safety_stock,
                           SUM(qty) as total_qty,
                           CASE 
                             WHEN SUM(qty) < safety_stock THEN '低於安庫'
                             ELSE '充足'
                           END as status
                    FROM v_item_stock_summary
                    WHERE warehouse_name = ?
                    GROUP BY part_no, item_name, spec, safety_stock
                    ORDER BY part_no
                """, (warehouse,))
                for row in cursor.fetchall():
                    tree.insert("", "end", values=(row[0], row[1], row[2], row[3], row[4], row[5]))

    def handle_add_fixture():
        part_no = part_no_entry.get().strip()
        name = name_entry.get().strip()
        spec = spec_entry.get().strip()
        try:
            safety_stock = int(safety_entry.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "安庫量必須是整數")
            return

        if not part_no or not name:
            messagebox.showerror("錯誤", "料號與品名為必填")
            return

        success, msg = insert_fixture(db_name, part_no, name, spec, safety_stock, current_user)
        if success:
            refresh_treeviews()
            part_no_entry.delete(0, tk.END)
            name_entry.delete(0, tk.END)
            spec_entry.delete(0, tk.END)
            safety_entry.delete(0, tk.END)
            messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("資料庫錯誤", msg)

    tk.Button(form_frame, text="新增治具", command=handle_add_fixture).grid(row=4, column=0, columnspan=2, pady=5)

    refresh_treeviews()

def build_fixture_bom_tab(parent, current_user, db_name):
    """建立治具 BOM 分頁（需求對應）"""
    tree = ttk.Treeview(parent, columns=("product_code", "part_no", "name", "qty", "note"), show="headings")
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    tree.heading("product_code", text="產品料號")
    tree.heading("part_no", text="治具料號")
    tree.heading("name", text="治具品名")
    tree.heading("qty", text="需求數量")
    tree.heading("note", text="備註")

    with sqlite3.connect(db_name) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT fb.product_code, f.part_no, f.name, fb.qty, fb.note
            FROM fixture_boms fb
            LEFT JOIN fixtures f ON fb.fixture_id = f.fixture_id
        """)
        for row in cursor.fetchall():
            tree.insert("", "end", values=row)