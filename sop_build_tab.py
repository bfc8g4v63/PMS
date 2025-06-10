import tkinter as tk
import threading
import shutil
import os
import sqlite3
from tkinter import filedialog, messagebox, ttk
import fitz
import re
from datetime import datetime
from utils import log_activity


UPLOAD_PATHS = {
    "dip": r"\\192.120.100.177\工程部\生產管理\SOP生成\DIP",
    "assembly": r"\\192.120.100.177\工程部\生產管理\SOP生成\組裝",
    "test": r"\\192.120.100.177\工程部\生產管理\SOP生成\測試",
    "packaging": r"\\192.120.100.177\工程部\生產管理\SOP生成\包裝",
    "oqc": r"\\192.120.100.177\工程部\生產管理\SOP生成\OQC"
}
SOP_SAVE_PATHS = {
    "dip": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\DIP_SOP",
    "assembly": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\組裝SOP",
    "test": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\測試SOP",
    "packaging": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\包裝SOP",
    "oqc": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\檢查表OQC"
}

def build_sop_upload_tab(tab_frame, current_user, db_name):
    role = current_user.get("role", "")
    specialty = current_user.get("specialty", "").lower()
    selected_uploads = []
    selected_files = []
    search_results = []

    if role == "engineer" and not specialty:
        tk.Label(tab_frame, text="您僅有查閱權限，無法執行 SOP 上傳。", fg="red").pack(padx=10, pady=20)
        return
    
    main_frame = tk.Frame(tab_frame)
    main_frame.pack(fill="both", expand=True)

    left = tk.Frame(main_frame, width=700)
    left.pack(side="left", fill="y", padx=10, pady=5)

    right = tk.Frame(main_frame)
    right.pack(side="left", fill="both", expand=True, padx=10, pady=5)

    main_frame.columnconfigure(0, weight=2)
    main_frame.columnconfigure(1, weight=1)

    if role == "admin":
        dest_path_var = tk.StringVar(value="dip")
        select_frame = tk.Frame(left)
        select_frame.pack(anchor="w")
        tk.Label(select_frame, text="專業分類：").pack(side="left")
        dropdown = ttk.Combobox(select_frame, textvariable=dest_path_var, values=list(UPLOAD_PATHS.keys()), state="readonly", width=12)
        dropdown.pack(side="left")
        def on_dropdown_change(event):
            search_results.clear()
            result_list.delete(0, tk.END)
            selected_files.clear()
        dropdown.bind("<<ComboboxSelected>>", on_dropdown_change)
    else:
        dest_path_var = tk.StringVar(value=specialty)
        search_path = UPLOAD_PATHS.get(specialty)
        if not search_path:
            tk.Label(left, text="您沒有 SOP 上傳權限", fg="red").pack(pady=20)
            return
    
    title_font = ("Arial", 10, "bold")
    tk.Label(left, text="📄 SOP 生成區", font=title_font, fg="navy").pack(anchor="w", pady=(10, 5))
    tk.Label(left, text="\nSOP 拼圖上傳").pack(anchor="w")
    upload_frame = tk.Frame(left)
    upload_frame.pack(anchor="w", pady=5)
    tk.Button(upload_frame, text="選擇PDF檔案", command=lambda: select_upload_files()).pack(side="left")
    tk.Button(upload_frame, text="上傳", command=lambda: upload_files()).pack(side="left", padx=10)

    upload_listbox = tk.Listbox(left, width=60, height=4)
    upload_listbox.pack(anchor="w", pady=5)

    def select_upload_files():
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file in files:
            if file not in selected_uploads:
                selected_uploads.append(file)
        refresh_upload_list()

    def refresh_upload_list():
        upload_listbox.delete(0, tk.END)
        for f in selected_uploads:
            upload_listbox.insert(tk.END, os.path.basename(f))

    def upload_files():
        dest_path = UPLOAD_PATHS.get(dest_path_var.get())
        if not dest_path:
            messagebox.showerror("錯誤", "未能判定上傳路徑。")
            return

        if not selected_uploads:
            messagebox.showwarning("未選擇檔案", "請先選取要上傳的 PDF 檔案。")
            return

        success_count = 0
        failed_files = []
        for src in selected_uploads:
            try:
                filename = os.path.basename(src)
                dest = os.path.join(dest_path, filename)
                shutil.copy(src, dest)
                success_count += 1
            except Exception as e:
                failed_files.append((filename, str(e)))

        msg = f"成功上傳：{success_count} 筆"
        if failed_files:
            msg += f"\n失敗：{len(failed_files)} 筆\n" + "\n".join([f"{f}: {err}" for f, err in failed_files])
        messagebox.showinfo("上傳結果", msg)
        selected_uploads.clear()
        refresh_upload_list()

    tk.Label(left, text="SOP 拼圖區").pack(anchor="w", pady=(15, 0))
    search_frame = tk.Frame(left)
    search_frame.pack(anchor="w")
    entry_keyword = tk.Entry(search_frame, width=40)
    entry_keyword.pack(side="left")
    tk.Button(search_frame, text="搜尋", width=10, command=lambda: search_files()).pack(side="left", padx=10)

    result_frame = tk.Frame(left)
    result_frame.pack(anchor="w")
    result_list = tk.Listbox(result_frame, height=6, width=50)
    result_list.pack(side="left")
    scrollbar = tk.Scrollbar(result_frame, command=result_list.yview)
    scrollbar.pack(side="right", fill="y")
    result_list.config(yscrollcommand=scrollbar.set)

    def search_files():
        keyword = entry_keyword.get().strip().lower()
        search_results.clear()
        result_list.delete(0, tk.END)
        path = UPLOAD_PATHS.get(dest_path_var.get())
        if not path:
            return

        if '&' in keyword:
            terms = [k.strip() for k in keyword.split('&')]
            for f in os.listdir(path):
                if all(term in f.lower() for term in terms) and f.lower().endswith(".pdf"):
                    search_results.append(f)
        elif '/' in keyword:
            terms = [k.strip() for k in keyword.split('/')]
            for f in os.listdir(path):
                if any(term in f.lower() for term in terms) and f.lower().endswith(".pdf"):
                    search_results.append(f)
        else:
            for f in os.listdir(path):
                if keyword in f.lower() and f.lower().endswith(".pdf"):
                    search_results.append(f)

        for f in search_results:
            result_list.insert(tk.END, f)

        if not search_results:
            messagebox.showinfo("查無結果", "查無符合的 PDF 檔案")

    def on_double_click(event):
        idx = result_list.curselection()
        if idx:
            fname = search_results[idx[0]]
            if fname in selected_files:
                selected_files.remove(fname)
            else:
                selected_files.append(fname)
            refresh_sort_list()

    result_list.bind("<Double-Button-1>", on_double_click)

    tk.Label(left, text="已排序項目：").pack(anchor="w", pady=(10, 0))
    sort_frame = tk.Frame(left)
    sort_frame.pack(anchor="w")
    sort_list = tk.Listbox(sort_frame, height=6, width=50)
    sort_list.pack(side="left")
    arrow_frame = tk.Frame(sort_frame)
    arrow_frame.pack(side="left", padx=5)
    tk.Button(arrow_frame, text="▲", command=lambda: move_up()).pack(pady=1)
    tk.Button(arrow_frame, text="▼", command=lambda: move_down()).pack(pady=1)

    def refresh_sort_list():
        sort_list.delete(0, tk.END)
        for i, f in enumerate(selected_files):
            sort_list.insert(tk.END, f"{i+1}. {f}")

    def move_up():
        idx = sort_list.curselection()
        if idx and idx[0] > 0:
            i = idx[0]
            selected_files[i-1], selected_files[i] = selected_files[i], selected_files[i-1]
            refresh_sort_list()
            sort_list.select_set(i-1)

    def move_down():
        idx = sort_list.curselection()
        if idx and idx[0] < len(selected_files)-1:
            i = idx[0]
            selected_files[i+1], selected_files[i] = selected_files[i], selected_files[i+1]
            refresh_sort_list()
            sort_list.select_set(i+1)

    def generate_pdf():
        threading.Thread(target=generate_pdf_thread).start()

    def generate_pdf_thread():
        output_name = entry_filename.get().strip()

        if not output_name:
            entry_filename.after(0, lambda: messagebox.showwarning("請輸入檔名", "請輸入儲檔名稱"))
            return
        if not selected_files:
            entry_filename.after(0, lambda: messagebox.showwarning("未選擇內容", "請先選擇並排序要合併的 PDF"))
            return

        if not re.match(r"^\d{8,12}_.+$", output_name):
            entry_filename.after(0, lambda: messagebox.showerror("錯誤", "請依格式輸入：料號_品名（例：12345678_產品名）"))
            return

        specialty_key = dest_path_var.get()
        save_dir = SOP_SAVE_PATHS.get(specialty_key)

        if not save_dir:
            entry_filename.after(0, lambda: messagebox.showerror("錯誤", f"無法判定專長「{specialty_key}」的儲存路徑"))
            return

        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        final_filename = f"{output_name}_{timestamp}.pdf"
        save_path = os.path.join(save_dir, final_filename)

        try:
            merged_pdf = fitz.open()
            skipped = []
            source_dir = UPLOAD_PATHS.get(dest_path_var.get())
            total_files = len(selected_files)

            def update_progress(value, text):
                progress_bar.after(0, lambda: progress_var.set(value))
                status_label.after(0, lambda: status_var.set(text))

            update_progress(0, "開始合併 PDF...")

            for i, f in enumerate(selected_files):
                full_path = os.path.join(source_dir, f)
                if os.path.exists(full_path):
                    doc = fitz.open(full_path)
                    merged_pdf.insert_pdf(doc)
                    doc.close()
                else:
                    skipped.append(f)

                progress = (i + 1) / total_files * 100
                update_progress(progress, f"處理中: {f}")

            update_progress(100, "完成合併，儲存中...")
            merged_pdf.save(save_path)
            merged_pdf.close()

            log_activity(db_name, current_user.get("user"), "generate_sop", final_filename, module="SOP生成")
            entry_filename.after(0, lambda: messagebox.showinfo("成功", f"已儲存拼圖式 SOP"))

            if skipped:
                skipped_str = "\n".join(skipped)
                entry_filename.after(0, lambda: messagebox.showwarning("部分檔案遺失", f"以下檔案未找到，未合併:\n{skipped_str}"))

            update_progress(100, "SOP 生成完成 ✔")

        except Exception as e:
            entry_filename.after(0, lambda: messagebox.showerror("錯誤", f"儲存失敗: {e}"))

    tk.Label(left, text="存檔名稱：").pack(anchor="w")

    filename_frame = tk.Frame(left)
    filename_frame.pack(anchor="w", pady=5)
    entry_filename = tk.Entry(filename_frame, width=30)
    entry_filename.pack(side="left")

    tk.Button(filename_frame, text="生成 SOP", bg="lightgreen", width=12, command=generate_pdf).pack(side="left", padx=10)
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(left, variable=progress_var, maximum=100, length=300)
    progress_bar.pack(anchor="w", pady=5)

    status_var = tk.StringVar(value="等待操作")
    status_label = tk.Label(left, textvariable=status_var, fg="blue")
    status_label.pack(anchor="w", pady=2)

def build_sop_apply_section(parent_frame, current_user, db_name):
    title_font = ("Arial", 10, "bold")
    tk.Label(parent_frame, text="📚 SOP 套用區", font=title_font, fg="navy").pack(anchor="w", pady=(10, 5))
    role = current_user.get("role", "")
    specialty = current_user.get("specialty", "").lower()

    if role == "engineer" and not specialty:
        lbl = tk.Label(parent_frame,
                       text="您僅有查閱權限，無法執行 SOP 套用。", 
                       fg="red")
        lbl.pack(padx=10, pady=20)
        return

    main_wrapper = tk.Frame(parent_frame)
    main_wrapper.pack(anchor="nw", padx=0, pady=5)
    search_frame = tk.Frame(main_wrapper)
    search_frame.pack(anchor="nw", padx=10, pady=5)

    entry_apply_search = tk.Entry(search_frame, width=20)
    entry_apply_search.pack(side="left", padx=(0, 5))
    search_btn = tk.Button(search_frame, text="來源搜尋", command=lambda: search_apply_files())
    search_btn.pack(side="left", padx=(0, 20))
    source_var = tk.StringVar()
    source_entry = tk.Entry(search_frame, textvariable=source_var, width=25)
    source_entry.pack(side="left", padx=(0, 5))

    tk.Label(search_frame, text="指定來源", anchor="w").pack(side="left", padx=(0, 5))

    role = current_user.get("role", "")
    specialty = current_user.get("specialty", "").lower()
    allow_all = role == "admin"

    tk.Label(main_wrapper, text="來源清單").pack(anchor="nw", padx=10, pady=(10, 0))

    sub_list_frame = tk.Frame(main_wrapper)
    sub_list_frame.pack(anchor="nw", padx=10)
    sub_items = []
    sub_checks = []

    if role == "admin":
        dest_path_var = tk.StringVar(value="dip")
        tk.Label(search_frame, text="專業分類：").pack(side="left", padx=5)
        dropdown = ttk.Combobox(search_frame, textvariable=dest_path_var, values=list(SOP_SAVE_PATHS.keys()), state="readonly", width=10)
        dropdown.pack(side="left")
    else:
        dest_path_var = tk.StringVar(value=specialty)

    keyword_frame = tk.Frame(main_wrapper)
    keyword_frame.pack(anchor="nw", padx=10, pady=(0, 5))

    entry_keyword2 = tk.Entry(keyword_frame, width=20)
    entry_keyword2.pack(side="left", padx=(0, 5))
    tk.Button(keyword_frame, text="套用搜尋", command=lambda: search_apply_targets()).pack(side="left", padx=(0, 20))

    apply_frame = tk.LabelFrame(main_wrapper, text="套用清單", width=720)
    apply_frame.pack(anchor="nw", padx=10, pady=(10, 0), fill="none")
    btn_frame = tk.Frame(main_wrapper)
    btn_frame.pack(anchor="nw", padx=10, pady=(5, 0))
    
    apply_canvas = tk.Canvas(apply_frame)
    scroll_y = ttk.Scrollbar(apply_frame, orient="vertical", command=apply_canvas.yview)
    scroll_x = ttk.Scrollbar(apply_frame, orient="horizontal", command=apply_canvas.xview)

    apply_scroll = tk.Frame(apply_canvas)

    apply_scroll.bind("<Configure>", lambda e: apply_canvas.configure(scrollregion=apply_canvas.bbox("all")))
    apply_canvas.create_window((0, 0), window=apply_scroll, anchor="nw")
    apply_canvas.configure(height=170, width=720, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)


    apply_canvas.pack(side="left", fill="both", expand=True)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")    

    apply_items = []
    apply_checks = []

    def search_apply_targets():
        keyword = entry_keyword2.get().strip().lower()
        if not keyword:
            messagebox.showinfo("提示", "請輸入料號/品名關鍵字")
            return

        for widget in apply_scroll.winfo_children():
            widget.destroy()

        apply_items.clear()
        apply_checks.clear()

        terms_and = [t.strip() for t in keyword.split('&')] if '&' in keyword else []
        terms_or = [t.strip() for t in keyword.split('/')] if '/' in keyword else []
        
        with sqlite3.connect(db_name, timeout=5) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            cursor = conn.cursor()
            cursor.execute("SELECT product_code, product_name FROM issues")
            all_data = cursor.fetchall()
            #conn.close()

        for code, name in all_data:
            combined = f"{code}_{name}".lower()
            matched = False
            if terms_and:
                matched = all(term in combined for term in terms_and)
            elif terms_or:
                matched = any(term in combined for term in terms_or)
            else:
                matched = keyword in combined

            if matched:
                var = tk.BooleanVar()
                cb = tk.Checkbutton(apply_scroll, text=f"{code}_{name}", variable=var, anchor="w", justify="left")
                cb.pack(anchor="w")
                apply_items.append((code, name))
                apply_checks.append(var)
    
    tree = ttk.Treeview(sub_list_frame, columns=("filename",), show="headings", selectmode="extended", height=7)
    tree.heading("filename", text="檔案名稱")
    tree.column("filename", width=720, anchor="w")
    tree.pack(side="left", fill="both", expand=True)

    scrollbar_x = ttk.Scrollbar(sub_list_frame, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")
    tree.configure(xscrollcommand=scrollbar_x.set) 
    scrollbar = ttk.Scrollbar(sub_list_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")

    def on_treeview_double_click(event):
        item = tree.focus()
        if item:
            filename = tree.item(item)['values'][0]
            if source_var.get() == filename:
                source_var.set("")
            else:
                source_var.set(filename)
    
    tree.bind("<Double-1>", on_treeview_double_click)

    def search_apply_files():
        keyword = entry_apply_search.get().strip().lower()
        for row in tree.get_children():
            tree.delete(row)
        sub_items.clear()
        sub_checks.clear()

        selected_key = dest_path_var.get()
        search_path = SOP_SAVE_PATHS.get(selected_key)
        if not search_path:
            messagebox.showerror("錯誤", f"無法決定 {selected_key} 對應的路徑")
            return

        files = []
        if os.path.isdir(search_path):
            files = [os.path.join(search_path, f) for f in os.listdir(search_path) if keyword in f.lower() and f.lower().endswith(".pdf")]

        for f in files:
            tree.insert("", "end", values=(os.path.basename(f),))
            sub_items.append(f)

    def select_all():
        for v in apply_checks:
            v.set(True)

    tk.Button(btn_frame, text="全選", command=select_all).pack(side="left", padx=5)

    def apply_to_all():
        threading.Thread(target=apply_thread).start()
        
    tk.Button(btn_frame, text="套用", command=apply_to_all).pack(side="left", padx=5)

    def apply_thread():
        if role == "engineer" and not specialty:
            messagebox.showerror("權限限制", "您僅有查閱權限，無法執行 SOP 套用。")
            return

        specialty_key = dest_path_var.get()
        if not allow_all and specialty_key != specialty:
            messagebox.showerror("錯誤", f"您只能操作自己的專長：{specialty}，目前選擇的是 {specialty_key}")
            return

        main_filename = source_var.get()
        if not main_filename:
            messagebox.showerror("錯誤", "請先指定資料來源")
            return

        matched_main_path = None
        for path in SOP_SAVE_PATHS.values():
            possible = os.path.join(path, main_filename)
            if os.path.exists(possible):
                matched_main_path = possible
                break

        if not matched_main_path:
            messagebox.showerror("錯誤", "找不到對應的資料來源原始路徑")
            return

        selected = [(code, name) for (code, name), v in zip(apply_items, apply_checks) if v.get()]
        if not selected:
            messagebox.showinfo("提示", "請勾選至少一筆套用資料")
            return

        def update_progress(pct, text):
            progress_bar.after(0, lambda: progress_var.set(pct))
            status_label.after(0, lambda: status_var.set(text))

        update_progress(0, "開始套用...")

        count = 0
        total = len(selected)

        field_map = {
            "dip": "dip_sop",
            "assembly": "assembly_sop",
            "test": "test_sop",
            "packaging": "packaging_sop",
            "oqc": "oqc_checklist"
        }
        field_name = field_map.get(specialty_key)

        if not field_name:
            messagebox.showerror("錯誤", f"無法決定 {specialty_key} 的資料欄位")
            return

        dest_dir = SOP_SAVE_PATHS.get(specialty_key)
        if not dest_dir:
            messagebox.showerror("錯誤", f"無法決定 {specialty_key} 的儲存路徑")
            return

        for code, name in selected:
            try:
                timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
                base_name = f"{code}_{name}_{timestamp}"
                display_name = base_name + ".pdf"

                dest_path = os.path.join(dest_dir, display_name)
                shutil.copy(matched_main_path, dest_path)

                with sqlite3.connect(db_name, timeout=10) as conn:
                    conn.execute("PRAGMA journal_mode=WAL;")
                    cursor = conn.cursor()
                    cursor.execute(f"""
                        UPDATE issues
                        SET {field_name} = ?, created_at = ?
                        WHERE product_code = ?
                    """, (display_name, timestamp, code))
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                log_activity(db_name, current_user.get("user"), "apply_sop", display_name, module="SOP套用")

                count += 1
                update_progress(int(count / total * 100), f"套用中：{display_name}")

            except Exception as e:
                messagebox.showerror("錯誤", f"處理 {code}_{name} 失敗：{e}")

        update_progress(100, "SOP 套用完成 ✔")
        messagebox.showinfo("完成", f"已完成套用，共處理 {count} 筆")

        search_apply_files()

    progress_var = tk.DoubleVar()
    progress_frame = tk.Frame(main_wrapper)
    progress_frame.pack(anchor="w", padx=10, pady=(5, 0))

    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=300)
    progress_bar.pack(side="left")

    status_var = tk.StringVar(value="等待操作")
    status_label = tk.Label(progress_frame, textvariable=status_var, fg="blue")
    status_label.pack(side="left", padx=(10, 0))