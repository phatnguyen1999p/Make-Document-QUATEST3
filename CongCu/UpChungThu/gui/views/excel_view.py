"""
Excel Processing View for scanning workbooks, selecting/merging sheets,
and exporting to calculation or certificate PDF documents with post-export verification.
"""

import os
import re
import threading
from typing import List, Dict, Any, Optional
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pythoncom
import win32com.client

from gui.theme import TM
from gui.components.base_processor import MappingProcessor
from gui.dialogs.signer_dialog import SignerConfigDialog
from gui.dialogs.bulk_config_dialog import BulkConfigDialog
from gui.dialogs.status_dialog import StatusProgressDialog
from core.pdf_service import PdfService, detect_language
from core.excel_service import ExcelEngine


class ExcelProcessorView(MappingProcessor):
    """
    View for multi-sheet Excel scanning, merging, and PDF exporting.
    """

    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        self.input_folder = ""
        self.output_folder = tk.StringVar(value="Mặc định (Cùng thư mục file gốc)")
        self.file_data = {}
        self.bulk_selected_categories: Dict[str, tk.IntVar] = {}
        self.bulk_connectors: List[tk.IntVar] = []
        self.all_unique_sheets: List[str] = []
        self.col_widths: Dict[str, int] = {}

        self.add_bg_var = tk.BooleanVar(value=True)
        self.batch_signers = None
        self.batch_lock = threading.Lock()

        self.load_excel_settings()
        self.add_bg_var.trace_add("write", lambda *a: self.save_excel_settings())

        # Navigation Bar
        self.nav = tk.Frame(
            self, bg=TM["bg_card"], pady=12, padx=25,
            highlightbackground=TM["border"], highlightthickness=1
        )
        self.nav.pack(fill="x", pady=(0, 15))
        self.btn_back = tk.Button(
            self.nav, text="⬅ Menu", font=TM["font_normal"],
            command=lambda: controller.show_frame("MainMenu"),
            bd=0, bg=TM["accent"], fg="white", cursor="hand2", padx=15, pady=5
        )
        self.btn_back.pack(side="left")
        self.lbl_title = tk.Label(
            self.nav, text="📊 XỬ LÝ FILE EXCEL (GỘP SHEET & KIỂM TRA)",
            font=TM["font_header"], bg=TM["bg_card"], fg=TM["text_main"]
        )
        self.lbl_title.pack(side="left", padx=25)

        # Control Card
        self.control_card = tk.Frame(
            self, bg=TM["bg_card"], padx=25, pady=15,
            highlightbackground=TM["border"], highlightthickness=1
        )
        self.control_card.pack(fill="x", padx=20, pady=5)

        self.bulk_panel = tk.Frame(self, bg=TM["bg_card"], padx=20)
        self.bulk_panel.pack(fill="x")

        self.create_search_widgets(self.control_card)

        # Input folder row
        self.f1 = tk.Frame(self.control_card, bg=TM["bg_main"], padx=15, pady=10, bd=1, relief="solid")
        self.f1.pack(fill="x", pady=5)
        self.btn_sel_folder = tk.Button(
            self.f1, text="📁 1. Chọn thư mục Excel", command=self.load_folder,
            bg=TM["primary"], fg="white", font=("Segoe UI", 11, "bold"), width=25, bd=0, cursor="hand2", pady=5
        )
        self.btn_sel_folder.pack(side="left")

        self.btn_open_in = tk.Button(
            self.f1, text="📂 Mở", command=self.open_input_folder,
            bg=TM["success"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=12, pady=2
        )
        self.btn_open_in.pack(side="right", padx=5)
        self.btn_copy_in = tk.Button(
            self.f1, text="📋 Copy", command=self.copy_input_path,
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=12, pady=2
        )
        self.btn_copy_in.pack(side="right", padx=5)

        self.lbl_in_path = tk.Label(
            self.f1, text="Chưa chọn thư mục...", bg=TM["bg_main"],
            fg=TM["text_muted"], font=TM["font_normal"], anchor="w", justify="left", wraplength=700
        )
        self.lbl_in_path.pack(side="left", fill="x", expand=True, padx=15)

        # Output folder row
        self.f2 = tk.Frame(self.control_card, bg=TM["bg_main"], padx=15, pady=10, bd=1, relief="solid")
        self.f2.pack(fill="x", pady=5)
        self.btn_sel_out = tk.Button(
            self.f2, text="🎯 2. Chọn nơi lưu PDF", command=self.select_output_folder,
            bg=TM["success"], fg="white", font=("Segoe UI", 11, "bold"), width=25, bd=0, cursor="hand2", pady=5
        )
        self.btn_sel_out.pack(side="left")

        self.btn_open_out = tk.Button(
            self.f2, text="📂 Mở", command=self.open_output_folder,
            bg=TM["success"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=12, pady=2
        )
        self.btn_open_out.pack(side="right", padx=5)
        self.btn_copy_out = tk.Button(
            self.f2, text="📋 Copy", command=self.copy_output_path,
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=12, pady=2
        )
        self.btn_copy_out.pack(side="right", padx=5)

        self.lbl_out_path = tk.Label(
            self.f2, textvariable=self.output_folder, bg=TM["bg_main"],
            fg=TM["text_muted"], font=TM["font_normal"], anchor="w", justify="left", wraplength=700
        )
        self.lbl_out_path.pack(side="left", fill="x", expand=True, padx=15)

        # Options Frame
        self.opt_frame = tk.LabelFrame(
            self.control_card, text="✅ Tùy chọn chung", bg=TM["bg_card"],
            font=TM["font_header"], fg=TM["text_main"], pady=10, padx=15, bd=1, relief="solid"
        )
        self.opt_frame.pack(fill="x", pady=(10, 5))

        self.chk_add_bg_excel = tk.Checkbutton(
            self.opt_frame, text="Chèn Giấy Tiêu Đề", variable=self.add_bg_var,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["text_main"],
            activeforeground=TM["text_main"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_add_bg_excel.pack(side="left", padx=15)

        self.btn_signer_config_ex = tk.Button(
            self.opt_frame, text="👤 Cấu hình Chữ ký", command=self.open_signer_config,
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=10, pady=2
        )
        self.btn_signer_config_ex.pack(side="left", padx=15)

        self.lbl_info_merge = tk.Label(
            self.opt_frame, text="💡 Tích [ + ] để GỘP sheet chung vào 1 file PDF.",
            bg=TM["bg_card"], fg=TM["warning"], font=("Segoe UI", 10, "italic")
        )
        self.lbl_info_merge.pack(side="right", padx=10)

        self.lbl_total = tk.Label(
            self, text="Tổng số file: 0", bg=TM["bg_main"],
            fg=TM["text_muted"], font=TM["font_small"], anchor="w"
        )
        self.lbl_total.pack(fill="x", padx=20, pady=(0, 5))

        # Action Buttons
        self.btn_frame = tk.Frame(self, bg=TM["bg_main"])
        self.btn_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.btn_export_calculation = tk.Button(
            self.btn_frame, text="▶ Xuất tính toán (Không kiểm tra)",
            command=lambda: self.process("calculation"), state="disabled", height=2,
            bg=TM["border"], fg="white", font=TM["font_header"], bd=0, cursor="hand2"
        )
        self.btn_export_calculation.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_export_certificate = tk.Button(
            self.btn_frame, text="▶ Xuất chứng thư (Có kiểm tra)",
            command=lambda: self.process("certificate"), state="disabled", height=2,
            bg=TM["border"], fg="white", font=TM["font_header"], bd=0, cursor="hand2"
        )
        self.btn_export_certificate.pack(side="right", fill="x", expand=True, padx=(10, 0))

        # Table Container
        self.table_container = tk.Frame(
            self, bg=TM["bg_card"], bd=1, relief="solid",
            highlightbackground=TM["border"], highlightthickness=1
        )
        self.table_container.pack(fill="both", expand=True, padx=20, pady=(5, 10))

        self.tree = ttk.Treeview(self.table_container, show='headings', selectmode="none")
        self.controller.set_tree_tags(self.tree)
        self.tree.bind("<Button-1>", self.on_tree_click)

        self.v_scrollbar = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        self.h_scrollbar = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        self.h_scrollbar.pack(side="bottom", fill="x")
        self.v_scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

    def update_theme(self):
        self.configure(bg=TM["bg_main"])
        self.nav.configure(bg=TM["bg_card"], highlightbackground=TM["border"])
        self.btn_back.configure(bg=TM["accent"])
        self.lbl_title.configure(bg=TM["bg_card"], fg=TM["text_main"])
        self.control_card.configure(bg=TM["bg_card"], highlightbackground=TM["border"])
        self.bulk_panel.configure(bg=TM["bg_card"])

        self.f1.configure(bg=TM["bg_main"])
        self.btn_sel_folder.configure(bg=TM["primary"])
        self.btn_open_in.configure(bg=TM["success"])
        self.btn_copy_in.configure(bg=TM["accent"])
        self.lbl_in_path.configure(bg=TM["bg_main"], fg=TM["text_main"] if self.input_folder else TM["text_muted"])

        self.f2.configure(bg=TM["bg_main"])
        self.btn_sel_out.configure(bg=TM["success"])
        self.btn_open_out.configure(bg=TM["success"])
        self.btn_copy_out.configure(bg=TM["accent"])
        self.lbl_out_path.configure(bg=TM["bg_main"], fg=TM["text_muted"])

        self.opt_frame.configure(bg=TM["bg_card"], fg=TM["text_main"])
        self.chk_add_bg_excel.configure(
            bg=TM["bg_card"], activebackground=TM["bg_card"],
            fg=TM["text_main"], activeforeground=TM["text_main"], selectcolor=TM["bg_card"]
        )
        self.lbl_info_merge.configure(bg=TM["bg_card"], fg=TM["warning"])

        self.lbl_total.configure(bg=TM["bg_main"], fg=TM["text_muted"])
        self.btn_frame.configure(bg=TM["bg_main"])
        self.table_container.configure(bg=TM["bg_card"], highlightbackground=TM["border"])

        calc_state = self.btn_export_calculation.cget("state")
        calc_text = self.btn_export_calculation.cget("text")
        if calc_state == "normal":
            self.btn_export_calculation.configure(bg=TM["primary"])
        elif "ĐANG" in calc_text:
            self.btn_export_calculation.configure(bg=TM["warning"])
        else:
            self.btn_export_calculation.configure(bg=TM["border"])

        cert_state = self.btn_export_certificate.cget("state")
        cert_text = self.btn_export_certificate.cget("text")
        if cert_state == "normal":
            self.btn_export_certificate.configure(bg=TM["success"])
        elif "ĐANG" in cert_text:
            self.btn_export_certificate.configure(bg=TM["warning"])
        else:
            self.btn_export_certificate.configure(bg=TM["border"])

        self.update_mapping_ui()
        self.controller.set_tree_tags(self.tree)

    def reset_to_defaults(self):
        super().reset_to_defaults()
        self.add_bg_var.set(True)
        self.input_folder = ""
        self.file_data = {}
        self.bulk_selected_categories = {}
        self.bulk_connectors = []
        self.all_unique_sheets = []
        self.output_folder.set("Mặc định (Cùng thư mục file gốc)")
        self.lbl_in_path.config(text="Chưa chọn thư mục...", fg=TM["text_muted"])
        for widget in self.bulk_panel.winfo_children():
            widget.destroy()
        self.tree["columns"] = ()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.btn_export_calculation.config(state="disabled", bg=TM["border"])
        self.btn_export_certificate.config(state="disabled", bg=TM["border"])

    def load_excel_settings(self):
        if not self.config_manager:
            return
        es = self.config_manager.get_excel_settings()
        self.add_bg_var.set(es.get("add_bg_var", True))

    def save_excel_settings(self):
        if not self.config_manager:
            return
        es = {"add_bg_var": self.add_bg_var.get()}
        self.config_manager.save_excel_settings(es)

    def select_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder.set(path)

    def open_input_folder(self):
        if not self.input_folder or not os.path.exists(self.input_folder):
            messagebox.showwarning("Cảnh báo", "Thư mục đầu vào chưa được chọn hoặc không tồn tại!")
            return
        try:
            os.startfile(os.path.normpath(self.input_folder))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")

    def copy_input_path(self):
        if not self.input_folder or "Chưa chọn" in self.input_folder:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục trước!")
            return
        self.clipboard_clear()
        self.clipboard_append(os.path.normpath(self.input_folder))
        orig_text, orig_bg = self.btn_copy_in.cget("text"), self.btn_copy_in.cget("background")
        self.btn_copy_in.config(text="✅ Đã copy!", bg=TM["success"])
        self.after(2000, lambda: self.btn_copy_in.config(text=orig_text, bg=orig_bg))

    def open_output_folder(self):
        path = self.output_folder.get()
        if "Mặc định" in path:
            path = self.input_folder
        if not path or not os.path.exists(path):
            messagebox.showwarning("Cảnh báo", "Thư mục đầu ra không tồn tại!")
            return
        try:
            os.startfile(os.path.normpath(path))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")

    def copy_output_path(self):
        path = self.output_folder.get()
        if "Mặc định" in path:
            path = self.input_folder
        if not path or ("Mặc định" in path and not path):
            messagebox.showwarning("Cảnh báo", "Đường dẫn trống hoặc chưa xác định!")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(os.path.normpath(path))
            orig_text, orig_bg = self.btn_copy_out.cget("text"), self.btn_copy_out.cget("background")
            self.btn_copy_out.config(text="✅ Đã copy!", bg=TM["success"])
            self.after(2000, lambda: self.btn_copy_out.config(text=orig_text, bg=orig_bg))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể copy đường dẫn: {e}")

    def open_signer_config(self):
        SignerConfigDialog(parent=self.controller.root, config_manager=self.config_manager)

    def create_bulk_header(self, all_sheets: List[str]):
        self.bulk_connectors = [tk.IntVar(value=0) for _ in range(len(all_sheets) - 1)]
        for s in all_sheets:
            self.bulk_selected_categories[s] = tk.IntVar(value=0)

        for widget in self.bulk_panel.winfo_children():
            widget.destroy()

        btn_config = tk.Button(
            self.bulk_panel, text="🛠️ Cấu hình Chọn nhanh & Gộp sheet",
            command=lambda: self.open_bulk_config_window(all_sheets),
            bg=TM["primary"], fg="white", font=TM["font_header"],
            bd=0, padx=20, pady=8, cursor="hand2"
        )
        btn_config.pack(pady=5, anchor="w")

    def open_bulk_config_window(self, all_sheets: List[str]):
        BulkConfigDialog(
            parent=self.controller.root,
            all_sheets=all_sheets,
            bulk_selected_categories=self.bulk_selected_categories,
            bulk_connectors=self.bulk_connectors,
            tree=self.tree,
            toggle_column_cb=self.toggle_column
        )

    def toggle_column(self, sheet_name: str, master_var: tk.IntVar):
        new_state = master_var.get()
        char = "☑" if new_state == 1 else "☐"
        col_idx = self.all_unique_sheets.index(sheet_name) + 2

        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, 'values'))
            if col_idx < len(vals) and vals[col_idx] != "":
                vals[col_idx] = char
                self.tree.item(item_id, values=vals)

    def on_folder_selected(self, folder_path: str):
        self.input_folder = folder_path
        self.lbl_in_path.config(text=self.input_folder, fg=TM["text_main"])
        self.start_scanning_folder(folder_path)

    def load_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.on_folder_selected(os.path.normpath(folder))

    def start_scanning_folder(self, folder_path: str):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for widget in self.bulk_panel.winfo_children():
            widget.destroy()
        self.file_data = {}
        self.all_unique_sheets = []
        self.col_widths = {}

        files = [f for f in os.listdir(folder_path) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not files:
            messagebox.showwarning("Thông báo", "Không tìm thấy file Excel.")
            return

        status_dialog = StatusProgressDialog(self.controller.root, "Đang quét cấu trúc file...")
        status_dialog.set_maximum(len(files))

        threading.Thread(
            target=self._scan_excel_thread,
            args=(files, status_dialog),
            daemon=True
        ).start()

    def _scan_excel_thread(self, files: List[str], status_dialog: StatusProgressDialog):
        def progress_cb(fname, count, total):
            self.controller.root.after(0, lambda: status_dialog.update_progress(f"Đang đọc sheets: {fname}", count))

        try:
            temp_data, all_unique_sheets, col_widths = ExcelEngine.scan_files(
                files=files,
                input_folder=self.input_folder,
                progress_callback=progress_cb
            )
            self.all_unique_sheets = all_unique_sheets
            self.col_widths = col_widths
            self.controller.root.after(0, lambda: self._finish_scan(temp_data, status_dialog))
        except Exception as e:
            self.controller.root.after(0, lambda: messagebox.showerror("Lỗi Quét", str(e)))
            self.controller.root.after(0, status_dialog.close)

    def _finish_scan(self, temp_data: List, status_dialog: StatusProgressDialog):
        self.create_bulk_header(self.all_unique_sheets)

        display_names = ["TÊN FILE", "V/E"] + self.all_unique_sheets + ["KẾT QUẢ KIỂM TRA"]
        cols = ["filename", "ve"] + self.all_unique_sheets + ["check"]
        self.tree["columns"] = cols

        for col, dname in zip(cols, display_names):
            self.tree.heading(col, text=dname)
            if col == "filename":
                self.tree.column(col, width=300, anchor="w")
            elif col == "ve":
                self.tree.column(col, width=60, anchor="center")
            elif col in self.all_unique_sheets:
                self.tree.column(col, width=self.col_widths[col], anchor="center")
            elif col == "check":
                self.tree.column(col, width=350, anchor="center")

        self.lbl_total.config(text=f"Tổng số file: {len(temp_data)}")
        for full_path, file, sheet_names in temp_data:
            lang = detect_language(file, sheet_names=sheet_names)
            vals = [file, lang]
            for s in self.all_unique_sheets:
                vals.append("☐" if s in sheet_names else "")
            vals += ["Sẵn sàng"]
            self.tree.insert("", "end", iid=full_path, values=vals, tags=("ready",))

        status_dialog.close()
        self._enable_buttons()

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        col_idx = int(column[1:]) - 1
        num_sheets = len(self.all_unique_sheets)
        if 2 <= col_idx <= num_sheets + 1:
            vals = list(self.tree.item(item_id, 'values'))
            if vals[col_idx] == "☐":
                vals[col_idx] = "☑"
            elif vals[col_idx] == "☑":
                vals[col_idx] = "☐"
            self.tree.item(item_id, values=vals)

    def process(self, export_type: str):
        self.tree["displaycolumns"] = "#all"
        self.btn_export_calculation.config(state="disabled", text="ĐANG XUẤT...", bg=TM["warning"])
        self.btn_export_certificate.config(state="disabled", text="ĐANG XUẤT...", bg=TM["warning"])
        threading.Thread(target=self._run_export_logic, args=(export_type,), daemon=True).start()

    def _run_export_logic(self, export_type: str):
        pythoncom.CoInitialize()
        self.batch_signers = None

        tasks = []
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, 'values')
            path = os.path.abspath(item_id)
            selected_sheets = [s_name for i, s_name in enumerate(self.all_unique_sheets) if vals[i + 2] == "☑"]
            if selected_sheets:
                tasks.append((path, selected_sheets, item_id))

        if not tasks:
            self.controller.root.after(0, lambda: messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một file và sheet để xuất."))
            self.controller.root.after(0, lambda: self._enable_buttons())
            pythoncom.CoUninitialize()
            return

        groups = []
        if self.all_unique_sheets:
            current_group = [self.all_unique_sheets[0]]
            for i, plus_var in enumerate(self.bulk_connectors):
                if plus_var.get() == 1:
                    current_group.append(self.all_unique_sheets[i + 1])
                else:
                    groups.append(current_group)
                    current_group = [self.all_unique_sheets[i + 1]]
            groups.append(current_group)

        out_root = self.output_folder.get()
        save_in_same_folder = (out_root == "Mặc định (Cùng thư mục file gốc)")

        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            if export_type == "calculation":
                final_add_bg = False
                final_check_date = False
                final_check_id = False
                final_check_page = False
            else:
                final_add_bg = self.add_bg_var.get()
                final_check_date = True
                final_check_id = True
                final_check_page = True

            for path, sheets, item_id in tasks:
                original_filename = os.path.splitext(os.path.basename(path))[0]
                wb = None
                try:
                    wb = excel.Workbooks.Open(path, ReadOnly=True, UpdateLinks=0)
                    base_dir = os.path.dirname(path) if save_in_same_folder else out_root

                    for grp in groups:
                        sheets_to_merge = [s for s in grp if s in sheets]
                        if not sheets_to_merge:
                            continue

                        self._update_status(item_id, "Đang xuất...", "warning")

                        save_dir = base_dir
                        bulk_selected_in_group = [
                            s for s in sheets_to_merge
                            if self.bulk_selected_categories.get(s) and self.bulk_selected_categories[s].get() == 1
                        ]
                        if bulk_selected_in_group:
                            folder_name = "+".join(bulk_selected_in_group)
                            save_dir = os.path.join(base_dir, folder_name)
                            if not os.path.exists(save_dir):
                                os.makedirs(save_dir)

                        clean_names = [re.sub(r'[\\/*?:"<>|]', "_", s) for s in sheets_to_merge]
                        pdf_filename = f"{original_filename}_{'+'.join(clean_names)}.pdf"
                        if len(pdf_filename) > 200:
                            pdf_filename = f"{original_filename}_Gop_{len(sheets_to_merge)}_sheets.pdf"

                        full_pdf_path = os.path.abspath(os.path.join(save_dir, pdf_filename))

                        wb.Worksheets(sheets_to_merge).Select()
                        excel.ActiveSheet.ExportAsFixedFormat(0, full_pdf_path)

                        if os.path.exists(full_pdf_path):
                            lang = detect_language(path, sheet_names=sheets_to_merge)
                            res = PdfService.process_pdf_post_export(
                                full_pdf_path,
                                bg_path_vn=self.controller.bg_pdf_path_vn,
                                bg_path_en=self.controller.bg_pdf_path_en,
                                add_bg=final_add_bg,
                                check_date=final_check_date,
                                check_id=final_check_id,
                                check_page=final_check_page,
                                check_signer=final_check_date,
                                target_date=None,
                                inspector_library=self.config_manager.inspector_library,
                                boss_library=self.config_manager.boss_library,
                                expected_insp=self.batch_signers[0] if self.batch_signers else None,
                                expected_boss=self.batch_signers[1] if self.batch_signers else None,
                                lang=lang
                            )
                            overall_check, tag, detected = res

                            with self.batch_lock:
                                if self.batch_signers is None and tag == "good" and detected[0] and detected[1]:
                                    self.batch_signers = detected
                        else:
                            overall_check, tag, detected = "❌ Xuất PDF lỗi", "bad", (None, None)

                        self._update_status(item_id, overall_check, tag)

                except Exception as ex:
                    self._update_status(item_id, f"❌ Lỗi: {ex}", "bad")
                finally:
                    if wb:
                        try:
                            wb.Close(False)
                        except Exception:
                            pass

            self.controller.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã xuất PDF và xử lý xong toàn bộ Excel!"))

        except Exception as e:
            self.controller.root.after(0, lambda err=e: messagebox.showerror("Lỗi xuất PDF", f"Quá trình thất bại: {err}"))
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.controller.root.after(0, lambda: self._enable_buttons())

    def _enable_buttons(self):
        self.btn_export_calculation.config(state="normal", text="▶ Xuất tính toán (Không kiểm tra)", bg=TM["primary"])
        self.btn_export_certificate.config(state="normal", text="▶ Xuất chứng thư (Có kiểm tra)", bg=TM["success"])

    def _update_status(self, item_id: str, check_text: str, tag: str):
        def do_update():
            vals = list(self.tree.item(item_id, 'values'))
            n = len(self.all_unique_sheets)
            if len(vals) > n + 2:
                vals[n + 2] = check_text
                vals = vals[:n + 3]
            else:
                vals.append(check_text)
            new_tags = (tag,) if tag else ()
            self.tree.item(item_id, values=vals, tags=new_tags)

        self.controller.root.after(0, do_update)


# Alias for backwards compatibility
ExcelProcessor = ExcelProcessorView
