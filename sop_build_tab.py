from tkinter import ttk
import tkinter as tk
import threading
import shutil
import os
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger
from datetime import datetime
from utils import log_activity

UPLOAD_PATHS = {
    "dip": r"\\192.120.100.177\工程部\生產管理\SOP生成\DIP",
    "assembly": r"\\192.120.100.177\工程部\生產管理\SOP生成\組裝",
    "test": r"\\192.120.100.177\工程部\生產管理\SOP生成\測試",
    "packaging": r"\\192.120.100.177\工程部\生產管理\SOP生成\包裝",
}
SOP_SAVE_PATHS = {
    "dip": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\DIP_SOP",
    "assembly": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\組裝SOP",
    "test": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\測試SOP",
    "packaging": r"\\192.120.100.177\工程部\生產管理\上齊SOP大禮包\包裝SOP",
}

def build_sop_upload_tab(tab_frame, current_user, db_name):
    role = current_user.get("role", "")
    specialty = current_user.get("specialty", "").lower()
    selected_uploads = []
    selected_files = []
    search_results = []

    main_frame = tk.Frame(tab_frame)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    left = tk.Frame(main_frame)
    left.pack(side="left", fill="y", anchor="n", padx=10)

    right = tk.Frame(main_frame, width=300)
    right.pack(side="left", fill="both", expand=True)

    if role == "admin":
        dest_path_var = tk.StringVar(value="dip")
        select_frame = tk.Frame(left)
        select_frame.pack(anchor="w")
        tk.Label(select_frame, text="選擇上傳分類：").pack(side="left")
        dropdown = ttk.Combobox(select_frame, textvariable=dest_path_var, values=list(UPLOAD_PATHS.keys()), state="readonly", width=10)
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

    tk.Label(left, text="\nSOP 批量上傳區").pack(anchor="w")
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

    tk.Label(left, text="SOP 拼圖搜尋區").pack(anchor="w", pady=(15, 0))
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
            entry_filename.after(0, lambda: messagebox.showwarning("請輸入檔名", "請輸入存檔名稱"))
            return
        if not selected_files:
            entry_filename.after(0, lambda: messagebox.showwarning("未選擇內容", "請先選擇並排序要合併的 PDF"))
            return

        specialty_key = dest_path_var.get()
        save_dir = SOP_SAVE_PATHS.get(specialty_key)

        if not save_dir:
            entry_filename.after(0, lambda: messagebox.showerror("錯誤", f"無法判定專長「{specialty_key}」的儲存路徑"))
            return

        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        final_filename = f"{timestamp}_{output_name}.pdf"
        save_path = os.path.join(save_dir, final_filename)

        try:
            merger = PdfMerger()
            skipped = []
            source_dir = UPLOAD_PATHS.get(dest_path_var.get())

            total_files = len(selected_files)
            if total_files == 0:
                entry_filename.after(0, lambda: messagebox.showwarning("未選擇內容", "請先選擇並排序要合併的 PDF"))
                return

            def update_progress(value, text):
                progress_bar.after(0, lambda: progress_var.set(value))
                status_label.after(0, lambda: status_var.set(text))

            update_progress(0, "開始合併 PDF...")

            for i, f in enumerate(selected_files):
                full_path = os.path.join(source_dir, f)
                if os.path.exists(full_path):
                    merger.append(full_path)
                else:
                    skipped.append(f)

                progress = (i + 1) / total_files * 100
                update_progress(progress, f"處理中：{f}")

            update_progress(100, "完成合併，儲存中...")

            merger.write(save_path)
            merger.close()

            log_activity(db_name, current_user.get("user"), "generate_sop", final_filename, module="SOP生成")
            entry_filename.after(0, lambda: messagebox.showinfo("成功", f"已儲存拼圖式 SOP：\n{save_path}"))

            if skipped:
                skipped_str = "\n".join(skipped)
                entry_filename.after(0, lambda: messagebox.showwarning("部分檔案遺失", f"以下檔案未找到，未合併：\n{skipped_str}"))

            update_progress(100, "SOP 生成完成 ✔")

        except Exception as e:
            entry_filename.after(0, lambda: messagebox.showerror("錯誤", f"儲存失敗：{e}"))

    tk.Label(left, text="存檔名稱：").pack(anchor="w")
    filename_frame = tk.Frame(left)
    filename_frame.pack(anchor="w", pady=5)
    entry_filename = tk.Entry(filename_frame, width=40)
    entry_filename.pack(side="left")
    tk.Button(filename_frame, text="生成 PDF", bg="lightgreen", width=12, command=generate_pdf).pack(side="left", padx=10)
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(left, variable=progress_var, maximum=100, length=300)
    progress_bar.pack(anchor="w", pady=5)

    status_var = tk.StringVar(value="等待操作")
    status_label = tk.Label(left, textvariable=status_var, fg="blue")
    status_label.pack(anchor="w", pady=2)   
    tk.Label(right, text="預留 SOP 套用區（待建置）", fg="gray").pack(anchor="nw", padx=10, pady=20)    
