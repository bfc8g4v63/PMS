#$ sop_info_tab.py
#% SOP資訊分頁（新增/查詢/更新/停用/開檔/複製）

import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from db_helper import get_conn
from utils import log_activity, open_file

SOP_FIELDS = {
    "dip": ("DIP SOP", "dip_sop", "dip_sop_bypass"),
    "assembly": ("組裝SOP", "assembly_sop", "assembly_sop_bypass"),
    "test": ("測試SOP", "test_sop", "test_sop_bypass"),
    "packaging": ("包裝SOP", "packaging_sop", "packaging_sop_bypass"),
    "oqc": ("檢查表OQC", "oqc_checklist", "oqc_checklist_bypass"),
}

DIP_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\DIP_SOP"
ASSEMBLY_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\組裝SOP"
TEST_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\測試SOP"
PACKAGING_SOP_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\包裝SOP"
OQC_PATH = r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\檢查表OQC"

SOP_BASE_PATHS = [DIP_SOP_PATH, ASSEMBLY_SOP_PATH, TEST_SOP_PATH, PACKAGING_SOP_PATH, OQC_PATH]

def _is_valid_file(file_path: str, field_name: str) -> bool:
    allowed_extensions = [".pdf"]
    if field_name == "oqc_checklist":
        allowed_extensions.append(".xlsx")
    ext = os.path.splitext(file_path)[1].lower()
    return ext in allowed_extensions

def _save_file_if_exist(file_path, target_folder, product_code, product_name, field_name):
    if not file_path:
        return "", ""
    if not os.path.exists(file_path):
        messagebox.showerror("錯誤", f"找不到檔案：{file_path}")
        return "", ""
    if not _is_valid_file(file_path, field_name):
        messagebox.showerror(
            "錯誤",
            f"這個檔案格式不合法：{file_path}\n只允許副檔名：.pdf{('、.xlsx' if field_name == 'oqc_checklist' else '')}",
        )
        return "", ""
    if not product_name:
        messagebox.showerror("錯誤", "品名為空，無法正確生成檔名")
        return "", ""

    timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    safe_name = product_name.replace("/", "-").replace("\\", "-").strip()
    if not safe_name:
        messagebox.showerror("錯誤", "品名不合法，無法正確生成檔名")
        return "", ""

    ext = os.path.splitext(file_path)[1].lower()
    filename = f"{product_code}_{safe_name}_{timestamp}{ext}"
    target_path = os.path.join(target_folder, filename)

    try:
        shutil.copy(file_path, target_path)
        return filename, timestamp
    except Exception as e:
        messagebox.showerror("錯誤", f"檔案儲存失敗: {e}")
        return "", ""

def _save_file(file_path, target_folder, username, product_code, product_name, field_name, log=True):
    if not os.path.exists(file_path):
        messagebox.showerror("錯誤", f"找不到檔案：{file_path}")
        return ""
    if not _is_valid_file(file_path, field_name):
        messagebox.showerror(
            "錯誤",
            f"這個檔案格式不合法：{file_path}\n只允許副檔名：.pdf{('、.xlsx' if field_name == 'oqc_checklist' else '')}",
        )
        return ""

    timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    safe_name = (product_name or "").replace("/", "-").replace("\\", "-").strip()
    if not safe_name:
        messagebox.showerror("錯誤", "品名不合法，無法正確生成檔名")
        return ""

    ext = os.path.splitext(file_path)[1].lower()
    filename = f"{product_code}_{safe_name}_{timestamp}{ext}"
    target_path = os.path.join(target_folder, filename)

    try:
        shutil.copy(file_path, target_path)
        if log:
            log_activity(user=username, action="upload", filename=filename, module="SOP資訊")
        return filename
    except Exception as e:
        messagebox.showerror("錯誤", f"檔案儲存失敗: {e}")
        return ""

def _update_sop_field(cursor, product_code, field_name, display_name):
    cursor.execute(
        f"UPDATE sop_information SET {field_name}=?, created_at=? WHERE product_code=?",
        (display_name, datetime.now().strftime("%Y%m%dT%H%M%S"), product_code),
    )

def _handle_sop_update(product_code, entry_widget, sop_path, field_name, current_user):
    path = entry_widget.get().strip()
    if not path:
        return None

    with get_conn() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT product_name FROM sop_information WHERE product_code=?", (product_code,))
        result = cursor.fetchone()
        if result:
            product_name = result[0]
        else:
            messagebox.showerror("錯誤", f"找不到料號 {product_code}，無法補上品名。")
            return None

    display_name = _save_file(path, sop_path, current_user, product_code, product_name, field_name, log=False)
    if not display_name:
        return None

    with get_conn() as conn:
        cursor = conn.cursor()
        _update_sop_field(cursor, product_code, field_name, display_name)
        if cursor.rowcount == 0:
            messagebox.showerror("錯誤", f"找不到料號 {product_code}，無法更新！")
            return None
        conn.commit()

    log_activity(user=current_user, action="update_sop", filename=display_name, module="SOP資訊")
    return display_name

def _create_sop_update_button(
    frame,
    row,
    label,
    sop_path,
    field_name,
    product_code_entry,
    entry_widget,
    current_user,
    user_specialty,
    role,
    allowed_specialty,
    on_refresh=None,
):
    def update_action():
        if role != "admin" and user_specialty != allowed_specialty:
            messagebox.showerror("權限限制", f"您無法上傳 {label}，僅限 {allowed_specialty} 工程師")
            return

        product_code = product_code_entry.get().strip()
        if not product_code:
            messagebox.showwarning("警告", "請先輸入料號")
            return

        updated_filename = _handle_sop_update(product_code, entry_widget, sop_path, field_name, current_user)
        if updated_filename:
            messagebox.showinfo("成功", f"已更新 {label} 檔案")
            try:
                if callable(on_refresh):
                    on_refresh()
            except Exception:
                pass

    btn = tk.Button(frame, text="更新", command=update_action)
    btn.grid(row=row, column=3, padx=5)
    return btn

def _create_upload_field_with_update(
    row,
    label,
    folder,
    field_name,
    form,
    product_code_entry,
    current_user,
    user_specialty,
    role,
    allowed_specialty,
    on_refresh=None,
):
    tk.Label(form, text=label).grid(row=row, column=0, sticky="e")
    entry = tk.Entry(form, width=50)
    entry.grid(row=row, column=1)

    def browse():
        path = filedialog.askopenfilename()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    tk.Button(form, text="選擇檔案", command=browse).grid(row=row, column=2)

    _create_sop_update_button(
        form,
        row,
        label,
        folder,
        field_name,
        product_code_entry,
        entry,
        current_user,
        user_specialty,
        role,
        allowed_specialty,
        on_refresh=on_refresh,
    )
    return entry

def build_sop_info_tab(tab_frame, db_name, login_info):
    current_user = login_info["user"]
    current_role = login_info["role"]
    can_add = int(login_info.get("can_add", 0) or 0)
    user_specialty = login_info.get("specialty", "") or ""

    root = tab_frame.winfo_toplevel()

    if current_role != "leader":
        form = tk.LabelFrame(tab_frame, text="新增紀錄")
        form.pack(fill="x", padx=10, pady=5)

        tk.Label(form, text="料號:").grid(row=0, column=0, sticky="e")
        entry_code = tk.Entry(form, width=50)
        entry_code.grid(row=0, column=1)

        tk.Label(form, text="品名:").grid(row=1, column=0, sticky="e")
        entry_name = tk.Entry(form, width=50)
        entry_name.grid(row=1, column=1)

        def _refresh():
            try:
                query_data()
            except Exception:
                pass

        entry_dip = _create_upload_field_with_update(
            2,
            "DIP SOP",
            DIP_SOP_PATH,
            "dip_sop",
            form,
            entry_code,
            current_user,
            user_specialty,
            current_role,
            "dip",
            on_refresh=_refresh,
        )

        entry_assembly = _create_upload_field_with_update(
            3,
            "組裝SOP",
            ASSEMBLY_SOP_PATH,
            "assembly_sop",
            form,
            entry_code,
            current_user,
            user_specialty,
            current_role,
            "assembly",
            on_refresh=_refresh,
        )

        entry_test = _create_upload_field_with_update(
            4,
            "測試SOP",
            TEST_SOP_PATH,
            "test_sop",
            form,
            entry_code,
            current_user,
            user_specialty,
            current_role,
            "test",
            on_refresh=_refresh,
        )

        entry_packaging = _create_upload_field_with_update(
            5,
            "包裝SOP",
            PACKAGING_SOP_PATH,
            "packaging_sop",
            form,
            entry_code,
            current_user,
            user_specialty,
            current_role,
            "packaging",
            on_refresh=_refresh,
        )

        entry_oqc = _create_upload_field_with_update(
            6,
            "檢查表OQC",
            OQC_PATH,
            "oqc_checklist",
            form,
            entry_code,
            current_user,
            user_specialty,
            current_role,
            "oqc",
            on_refresh=_refresh,
        )

        def save_data():
            code = entry_code.get().strip()
            name = entry_name.get().strip()

            if not name:
                messagebox.showerror("錯誤", "品名不能為空")
                return

            if len(code) not in (8, 12) or not code.isdigit():
                messagebox.showerror("錯誤", "必須為 8/12 碼數字")
                return

            with get_conn() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT product_code FROM sop_information WHERE product_code=?", (code,))
                if cursor.fetchone():
                    messagebox.showerror("錯誤", "料號已存在，請重新確認過。")
                    return

                d_file, _ = _save_file_if_exist(entry_dip.get().strip(), DIP_SOP_PATH, code, name, "dip_sop")
                a_file, _ = _save_file_if_exist(entry_assembly.get().strip(), ASSEMBLY_SOP_PATH, code, name, "assembly_sop")
                t_file, _ = _save_file_if_exist(entry_test.get().strip(), TEST_SOP_PATH, code, name, "test_sop")
                p_file, _ = _save_file_if_exist(entry_packaging.get().strip(), PACKAGING_SOP_PATH, code, name, "packaging_sop")
                o_file, _ = _save_file_if_exist(entry_oqc.get().strip(), OQC_PATH, code, name, "oqc_checklist")

                if d_file:
                    log_activity(user=current_user, action="upload", filename=d_file, module="SOP資訊")
                if a_file:
                    log_activity(user=current_user, action="upload", filename=a_file, module="SOP資訊")
                if t_file:
                    log_activity(user=current_user, action="upload", filename=t_file, module="SOP資訊")
                if p_file:
                    log_activity(user=current_user, action="upload", filename=p_file, module="SOP資訊")
                if o_file:
                    log_activity(user=current_user, action="upload", filename=o_file, module="SOP資訊")

                sql = (
                    "INSERT INTO sop_information "
                    "(product_code, product_name, dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist, created_by, created_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
                )
                cursor.execute(
                    sql,
                    (
                        code,
                        name,
                        d_file,
                        a_file,
                        t_file,
                        p_file,
                        o_file,
                        current_user,
                        datetime.now().strftime("%Y%m%dT%H%M%S"),
                    ),
                )
                conn.commit()

            messagebox.showinfo("成功", "已新增紀錄")
            for e in (entry_code, entry_name, entry_dip, entry_assembly, entry_test, entry_packaging, entry_oqc):
                e.delete(0, tk.END)
            query_data()

        tk.Button(
            form,
            text="新增紀錄",
            command=save_data,
            bg="lightblue",
            state="normal" if can_add else "disabled",
        ).grid(row=7, column=1, pady=10)

    query_frame = tk.Frame(tab_frame)
    query_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(query_frame, text="查詢關鍵字: ").pack(side="left")
    entry_query = tk.Entry(query_frame)
    entry_query.pack(side="left")

    sort_desc = tk.BooleanVar(value=True)

    def toggle_sort():
        sort_desc.set(not sort_desc.get())
        query_data()

    tk.Button(query_frame, text="↕排序", command=toggle_sort).pack(side="left", padx=5)
    tk.Button(query_frame, text="查詢", command=lambda: query_data()).pack(side="left")

    columns = ("料號", "品名", "DIP SOP", "組裝SOP", "測試SOP", "包裝SOP", "檢查表OQC", "SOP建立人", "建立時間")

    tree_container = tk.Frame(tab_frame)
    tree_container.pack(fill="both", expand=True, padx=10, pady=5)

    tree = ttk.Treeview(tree_container, columns=columns, show="headings")
    tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll_y.set)

    for col in columns:
        tree.heading(col, text=col)
        if col == "料號":
            tree.column(col, width=60, anchor="center")
        elif col == "品名":
            tree.column(col, width=180, anchor="center")
        elif col == "SOP建立人":
            tree.column(col, width=80, anchor="center")
        else:
            tree.column(col, width=120, anchor="center")

    tree.pack(side="left", fill="both", expand=True)
    tree_scroll_y.pack(side="right", fill="y")

    sop_stats_var = tk.StringVar()
    tk.Label(tab_frame, textvariable=sop_stats_var, anchor="w", fg="blue", font=("Arial", 10)).pack(fill="x", padx=10, pady=(0, 5))

    def update_sop_statistics():
        stats = {"DIP SOP": 0, "組裝 SOP": 0, "測試 SOP": 0, "包裝 SOP": 0, "檢查表 OQC": 0}
        for row_id in tree.get_children():
            values = tree.item(row_id)["values"]
            if values[2]:
                stats["DIP SOP"] += 1
            if values[3]:
                stats["組裝 SOP"] += 1
            if values[4]:
                stats["測試 SOP"] += 1
            if values[5]:
                stats["包裝 SOP"] += 1
            if values[6]:
                stats["檢查表 OQC"] += 1
        sop_stats_var.set("　｜　".join([f"{k}: {v}" for k, v in stats.items()]))

    def query_data():
        try:
            root.focus_force()
        except Exception:
            pass

        raw_input = entry_query.get().strip()

        for row in tree.get_children():
            tree.delete(row)

        with get_conn() as conn:
            cursor = conn.cursor()
            base_query = (
                "SELECT "
                "product_code, product_name, "
                "dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist, "
                "created_by, created_at, "
                "dip_sop_bypass, assembly_sop_bypass, test_sop_bypass, packaging_sop_bypass, oqc_checklist_bypass "
                "FROM sop_information"
            )

            conditions = []
            params = []

            if raw_input:
                if "&" in raw_input:
                    terms = [t.strip() for t in raw_input.split("&")]
                    for term in terms:
                        conditions.append("((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    condition_sql = " AND ".join(conditions)
                elif "/" in raw_input:
                    terms = [t.strip() for t in raw_input.split("/")]
                    sub_conditions = []
                    for term in terms:
                        sub_conditions.append("((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    condition_sql = " OR ".join(sub_conditions)
                else:
                    condition_sql = "((product_code COLLATE NOCASE) LIKE ? OR (product_name COLLATE NOCASE) LIKE ?)"
                    params = [f"%{raw_input}%", f"%{raw_input}%"]

            order_sql = f" ORDER BY created_at {'DESC' if sort_desc.get() else 'ASC'}"
            if raw_input:
                final_query = f"{base_query} WHERE {condition_sql}{order_sql}"
            else:
                final_query = f"{base_query}{order_sql} LIMIT 200"

            cursor.execute(final_query, params)
            rows = cursor.fetchall()

            for row in rows:
                product_code_raw = row[0]
                product_name = row[1]
                timestamp = row[8]
                product_code = str(product_code_raw).zfill(8)

                row_display = [product_code, product_name, "", "", "", "", "", row[7], timestamp]

                bypass_flags = {
                    2: int(row[9] or 0),
                    3: int(row[10] or 0),
                    4: int(row[11] or 0),
                    5: int(row[12] or 0),
                    6: int(row[13] or 0),
                }

                for i in range(2, 7):
                    sop_file = row[i]
                    if sop_file:
                        display_name = f"{product_code}_{product_name}_{timestamp}"
                        if bypass_flags.get(i, 0) == 1:
                            row_display[i] = f"（已停用） {display_name}"
                        else:
                            row_display[i] = display_name

                tree.insert("", tk.END, values=row_display)

        update_sop_statistics()

    def on_double_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index not in range(2, 7):
            return

        values = tree.item(item)["values"]
        product_code = values[0]
        if "（已停用）" in str(values[col_index]):
            return

        field_map = {2: "dip_sop", 3: "assembly_sop", 4: "test_sop", 5: "packaging_sop", 6: "oqc_checklist"}
        target_field = field_map.get(col_index)
        if not target_field:
            return

        with get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT {target_field} FROM sop_information WHERE product_code=?", (product_code,))
            result = cursor.fetchone()
            if result and result[0]:
                filename = result[0]
                sop_folder = SOP_BASE_PATHS[col_index - 2]
                full_path = os.path.join(sop_folder, filename)
                if os.path.exists(full_path):
                    open_file(full_path)
                else:
                    messagebox.showerror("錯誤", f"找不到檔案：{full_path}")

    tree.last_hovered_item = ""
    tree.last_hovered_col_index = 0

    def on_tree_motion(event):
        tree_widget = event.widget
        try:
            tree_widget.focus_set()
        except Exception:
            pass

        item = tree_widget.identify_row(event.y)
        col = tree_widget.identify_column(event.x)
        if item:
            tree_widget.last_hovered_item = item
        if col:
            try:
                index = int(col[1:]) - 1
            except ValueError:
                index = 0
            if index >= 0:
                tree_widget.last_hovered_col_index = index

    def on_copy(event):
        tree_widget = event.widget
        item_id = getattr(tree_widget, "last_hovered_item", "")
        col_index = getattr(tree_widget, "last_hovered_col_index", 0)
        if not item_id:
            return
        values = tree_widget.item(item_id, "values")
        if not values:
            return
        if col_index < 0 or col_index >= len(values):
            return
        if col_index not in (0, 1):
            return

        value_str = str(values[col_index])
        try:
            root.clipboard_clear()
            root.clipboard_append(value_str)
            root.update()
        except Exception:
            pass

    def toggle_bypass(product_code, field_name, bypass_field):
        with get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT {bypass_field} FROM sop_information WHERE product_code=?", (product_code,))
            current = cursor.fetchone()
            new_value = 0 if current and current[0] else 1
            cursor.execute(f"UPDATE sop_information SET {bypass_field}=? WHERE product_code=?", (new_value, product_code))
            conn.commit()

        log_activity(user=current_user, action="toggle_bypass", filename=f"{product_code}:{field_name}", module="SOP資訊")
        query_data()

    def on_right_click(event):
        if current_role not in ("admin", "engineer"):
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        if col_index not in range(2, 7):
            return

        product_code = tree.item(item)["values"][0]
        field_map = {
            2: ("dip_sop", "dip_sop_bypass"),
            3: ("assembly_sop", "assembly_sop_bypass"),
            4: ("test_sop", "test_sop_bypass"),
            5: ("packaging_sop", "packaging_sop_bypass"),
            6: ("oqc_checklist", "oqc_checklist_bypass"),
        }
        field_name, bypass_field = field_map[col_index]

        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="啟用/停用", command=lambda: toggle_bypass(product_code, field_name, bypass_field))
        menu.post(event.x_root, event.y_root)

    tree.bind("<Double-1>", on_double_click)
    tree.bind("<Motion>", on_tree_motion)
    tree.bind("<Control-c>", on_copy)
    tree.bind("<Button-3>", on_right_click)

    if current_role == "admin":
        def delete_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("提醒", "請先選取要刪除的資料")
                return
            if messagebox.askyesno("確認", "確定要刪除選取的資料？此操作無法復原。"):
                deleted_codes = []
                with get_conn() as conn:
                    cursor = conn.cursor()
                    for item in selected_items:
                        product_code = str(tree.item(item)["values"][0]).zfill(8)
                        cursor.execute("DELETE FROM sop_information WHERE product_code=?", (product_code,))
                        deleted_codes.append(product_code)
                    conn.commit()
                for code in deleted_codes:
                    log_activity(user=current_user, action="delete", filename=code, module="SOP資訊")
                query_data()

        delete_frame = tk.Frame(tab_frame)
        delete_frame.pack(fill="x", padx=10, pady=(0, 5), anchor="e")
        tk.Button(delete_frame, text="刪除選取資料", command=delete_selected, bg="lightcoral", fg="white").pack(side="right")

    query_data()