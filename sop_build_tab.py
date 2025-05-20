import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import shutil
import os
from PyPDF2 import PdfMerger

UPLOAD_PATHS = {
    "dip": r"\\192.120.100.177\工程部\生產管理\SOP生成\DIP",
    "assembly": r"\\192.120.100.177\工程部\生產管理\SOP生成\組裝",
    "test": r"\\192.120.100.177\工程部\生產管理\SOP生成\測試",
    "packaging": r"\\192.120.100.177\工程部\生產管理\SOP生成\包裝",
}

def build_sop_upload_tab(tab_frame, current_user):
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
        search_path = UPLOAD_PATHS["dip"]
    else:
        dest_path = UPLOAD_PATHS.get(specialty)
        if not dest_path:
            tk.Label(left, text="您沒有 SOP 上傳權限", fg="red").pack(pady=20)
            return
        search_path = dest_path

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
        if not selected_uploads:
            messagebox.showwarning("未選擇檔案", "請先選取要上傳的 PDF 檔案。")
            return

        dest_path = UPLOAD_PATHS.get(dest_path_var.get()) if role == "admin" else UPLOAD_PATHS.get(specialty)
        if not dest_path:
            messagebox.showerror("錯誤", "未能判定上傳路徑。")
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
        if '&' in keyword:
            terms = [k.strip() for k in keyword.split('&')]
            for f in os.listdir(search_path):
                if all(term in f.lower() for term in terms) and f.lower().endswith(".pdf"):
                    search_results.append(f)
        elif '/' in keyword:
            terms = [k.strip() for k in keyword.split('/')]
            for f in os.listdir(search_path):
                if any(term in f.lower() for term in terms) and f.lower().endswith(".pdf"):
                    search_results.append(f)
        else:
            for f in os.listdir(search_path):
                if keyword in f.lower() and f.lower().endswith(".pdf"):
                    search_results.append(f)
        result_list.delete(0, tk.END)
        for f in search_results:
            result_list.insert(tk.END, f)

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

        output_name = entry_filename.get().strip()
        if not output_name:
            messagebox.showwarning("請輸入檔名", "請輸入存檔名稱")
            return
        if not selected_files:
            messagebox.showwarning("未選擇內容", "請先選擇並排序要合併的 PDF")
            return

        merger = PdfMerger()
        for f in selected_files:
            full_path = os.path.join(search_path, f)
            merger.append(full_path)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=output_name, filetypes=[("PDF Files", "*.pdf")])
        if save_path:
            merger.write(save_path)
            merger.close()
            messagebox.showinfo("成功", f"已儲存合併 PDF：{save_path}")

    tk.Label(left, text="存檔名稱：").pack(anchor="w")
    filename_frame = tk.Frame(left)
    filename_frame.pack(anchor="w", pady=5)
    entry_filename = tk.Entry(filename_frame, width=40)
    entry_filename.pack(side="left")
    tk.Button(filename_frame, text="生成 PDF", bg="lightgreen", width=12, command=generate_pdf).pack(side="left", padx=10)

    tk.Label(right, text="預留 SOP 套用區（待建置）", fg="gray").pack(anchor="nw", padx=10, pady=20)
