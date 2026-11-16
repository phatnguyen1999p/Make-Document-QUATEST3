"""
Word Processing View for automated Word document manipulation,
date updates, PDF conversion, background insertion, and verification.
"""

import os
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pythoncom
import win32com.client

from gui.theme import TM
from gui.components.base_processor import MappingProcessor
from gui.dialogs.signer_dialog import SignerConfigDialog
from gui.dialogs.time_rules_dialog import TimeRulesDialog
from gui.dialogs.calendar_dialog import CalendarDialog
from core.utils import calculate_target_date
from core.pdf_service import PdfService, detect_language
from core.word_service import WordEngine


class WordProcessorView(MappingProcessor):
    """
    View for batch Word document processing.
    """

    def __init__(self, parent, controller):
        super().__init__(parent, controller)

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
            self.nav, text="📝 XỬ LÝ FILE WORD TỰ ĐỘNG",
            font=TM["font_header"], bg=TM["bg_card"], fg=TM["text_main"]
        )
        self.lbl_title.pack(side="left", padx=25)

        # Control Card
        self.control_card = tk.Frame(
            self, bg=TM["bg_card"], padx=25, pady=20,
            highlightbackground=TM["border"], highlightthickness=1
        )
        self.control_card.pack(fill="x", padx=20, pady=5)

        self.create_search_widgets(self.control_card)

        self.target_folder = tk.StringVar(value="Chưa chọn thư mục...")
        self.f_dir = tk.Frame(self.control_card, bg=TM["bg_main"], padx=15, pady=12, bd=1, relief="solid")
        self.f_dir.pack(fill="x", pady=5)
        self.btn_sel_dir = tk.Button(
            self.f_dir, text="📁 Chọn thư mục", command=self.select_folder,
            bg=TM["primary"], fg="white", font=TM["font_small"], bd=0, padx=12, pady=5, cursor="hand2"
        )
        self.btn_sel_dir.pack(side="left")
        self.lbl_proc_tag = tk.Label(
            self.f_dir, text="📂 Đang xử lý:", bg=TM["bg_main"], fg=TM["primary"], font=TM["font_header"]
        )
        self.lbl_proc_tag.pack(side="left", padx=(15, 5))

        self.btn_open_tbkq = tk.Button(
            self.f_dir, text="📂 Mở thư mục", command=self.open_tbkq_folder,
            bg=TM["success"], fg="white", bd=0, relief="flat", font=TM["font_small"], cursor="hand2", padx=10
        )
        self.btn_open_tbkq.pack(side="right", padx=5)

        self.btn_copy_path = tk.Button(
            self.f_dir, text="📋 Copy đường dẫn", command=self.copy_folder_path,
            bg=TM["accent"], fg="white", bd=0, relief="flat", font=TM["font_small"], cursor="hand2", padx=10
        )
        self.btn_copy_path.pack(side="right", padx=5)

        self.lbl_target_path = tk.Label(
            self.f_dir, textvariable=self.target_folder, bg=TM["bg_main"],
            fg=TM["text_main"], font=TM["font_normal"], anchor="w", justify="left", wraplength=700
        )
        self.lbl_target_path.pack(side="left", fill="x", expand=True, padx=10)

        # Options & State Variables
        self.run_date_update = tk.BooleanVar(value=True)
        self.run_conversion = tk.BooleanVar(value=True)
        self.add_bg_var = tk.BooleanVar(value=True)

        self.batch_lock = threading.Lock()
        self.batch_signers = None

        self.use_time_rules = tk.BooleanVar(value=False)
        self.rule_t1 = tk.StringVar(value="09:00")
        self.rule_t2 = tk.StringVar(value="16:00")
        self.rule_x1 = tk.IntVar(value=0)
        self.rule_x2 = tk.IntVar(value=1)
        self.rule_x3 = tk.IntVar(value=2)

        self.use_manual_date = tk.BooleanVar(value=False)
        self.manual_date_str = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))

        self.load_settings()

        # Traces for auto-saving and preview updates
        self.run_date_update.trace_add("write", lambda *a: self.save_settings())
        self.run_conversion.trace_add("write", lambda *a: self.save_settings())
        self.add_bg_var.trace_add("write", lambda *a: self.save_settings())
        self.use_time_rules.trace_add("write", lambda *a: self.update_preview())
        self.use_time_rules.trace_add("write", lambda *a: self.save_settings())
        self.use_manual_date.trace_add("write", lambda *a: self.update_preview())
        self.use_manual_date.trace_add("write", lambda *a: self.save_settings())
        self.manual_date_str.trace_add("write", lambda *a: self.update_preview())
        self.manual_date_str.trace_add("write", lambda *a: self.save_settings())

        # Options Frame UI
        self.f_opt = tk.LabelFrame(
            self.control_card, text="⚙️ Cấu hình xử lý", bg=TM["bg_card"],
            font=TM["font_header"], fg=TM["text_main"], padx=15, pady=12, bd=1, relief="solid"
        )
        self.f_opt.pack(fill="x", pady=(10, 0))

        self.f_opt_r1 = tk.Frame(self.f_opt, bg=TM["bg_card"])
        self.f_opt_r1.pack(fill="x", pady=(0, 5))
        self.f_opt_r2 = tk.Frame(self.f_opt, bg=TM["bg_card"])
        self.f_opt_r2.pack(fill="x")

        self.chk_date_update = tk.Checkbutton(
            self.f_opt_r1, text="Cập nhật Ngày", variable=self.run_date_update,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["text_main"],
            activeforeground=TM["text_main"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_date_update.pack(side="left", padx=(0, 10))

        self.chk_use_time_rules = tk.Checkbutton(
            self.f_opt_r1, text="Theo giờ:", variable=self.use_time_rules,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["text_main"],
            activeforeground=TM["text_main"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_use_time_rules.pack(side="left")

        self.btn_set_rules = tk.Button(
            self.f_opt_r1, text="⚙️", command=self.open_rule_settings,
            bg=TM["primary"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=5, pady=2
        )
        self.btn_set_rules.pack(side="left", padx=(0, 15))

        self.chk_use_manual_date = tk.Checkbutton(
            self.f_opt_r1, text="Nhập tay:", variable=self.use_manual_date,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["text_main"],
            activeforeground=TM["text_main"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_use_manual_date.pack(side="left")

        self.ent_manual_date = tk.Entry(
            self.f_opt_r1, textvariable=self.manual_date_str, width=10,
            font=TM["font_normal"], justify="center", bd=1, relief="solid",
            bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"]
        )
        self.ent_manual_date.pack(side="left", padx=(0, 5))

        self.btn_calendar = tk.Button(
            self.f_opt_r1, text="📅", command=self.open_calendar, bg=TM["bg_card"], bd=0, cursor="hand2"
        )
        self.btn_calendar.pack(side="left", padx=(0, 15))

        self.chk_conversion = tk.Checkbutton(
            self.f_opt_r2, text="Xuất PDF", variable=self.run_conversion,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["text_main"],
            activeforeground=TM["text_main"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_conversion.pack(side="left", padx=(0, 20))

        self.chk_add_bg = tk.Checkbutton(
            self.f_opt_r2, text="Chèn Giấy Tiêu Đề & Kiểm tra PDF", variable=self.add_bg_var,
            bg=TM["bg_card"], activebackground=TM["bg_card"], fg=TM["danger"],
            activeforeground=TM["danger"], selectcolor=TM["bg_card"], font=TM["font_normal"]
        )
        self.chk_add_bg.pack(side="left", padx=(0, 10))

        self.btn_signer_config = tk.Button(
            self.f_opt_r2, text="👤 Cấu hình Chữ ký", command=self.open_signer_config,
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, relief="flat", cursor="hand2", padx=10, pady=2
        )
        self.btn_signer_config.pack(side="left", padx=(0, 15))

        self.lbl_date_preview = tk.Label(
            self.f_opt_r2, text=f"(Sẽ dùng: {self.get_calculated_date().strftime('%d/%m/%Y')})",
            bg=TM["bg_card"], fg=TM["text_muted"], font=TM["font_small"]
        )
        self.lbl_date_preview.pack(side="left")

        # Action Button
        self.btn_run = tk.Button(
            self, text="▶ BẮT ĐẦU XỬ LÝ WORD", command=self.process,
            state="disabled", bg="#cbd5e1", fg="white", font=TM["font_header"], bd=0, height=2, cursor="hand2"
        )
        self.btn_run.pack(fill="x", padx=20, pady=15)

        self.lbl_total = tk.Label(
            self, text="Tổng số file: 0", bg=TM["bg_main"],
            fg=TM["text_muted"], font=TM["font_small"], anchor="w"
        )
        self.lbl_total.pack(fill="x", padx=20, pady=(0, 5))

        # Result Treeview Table
        table_frame = tk.Frame(
            self, bg=TM["bg_card"], padx=2, pady=2,
            highlightbackground=TM["border"], highlightthickness=1
        )
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        columns = ("File Name", "V/E", "Word Status", "Overall Check")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode="browse")
        self.controller.set_tree_tags(self.tree)
        display_names = ["TÊN FILE", "V/E", "TRẠNG THÁI WORD", "KẾT QUẢ KIỂM TRA"]
        for col, dname in zip(columns, display_names):
            self.tree.heading(col, text=dname)
            if col == "File Name":
                self.tree.column(col, anchor="w", width=250)
            elif col == "V/E":
                self.tree.column(col, anchor="center", width=60)
            elif col == "Word Status":
                self.tree.column(col, anchor="center", width=180)
            elif col == "Overall Check":
                self.tree.column(col, anchor="center", width=350)

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

    def update_theme(self):
        self.configure(bg=TM["bg_main"])
        self.nav.configure(bg=TM["bg_card"], highlightbackground=TM["border"])
        self.btn_back.configure(bg=TM["accent"])
        self.lbl_title.configure(bg=TM["bg_card"], fg=TM["text_main"])
        self.control_card.configure(bg=TM["bg_card"], highlightbackground=TM["border"])
        self.f_dir.configure(bg=TM["bg_main"])
        self.btn_sel_dir.configure(bg=TM["primary"])
        self.lbl_proc_tag.configure(bg=TM["bg_main"], fg=TM["primary"])
        self.btn_open_tbkq.configure(bg=TM["success"])
        self.btn_copy_path.configure(bg=TM["accent"])
        self.lbl_target_path.configure(bg=TM["bg_main"], fg=TM["text_main"])
        self.f_opt.configure(bg=TM["bg_card"], fg=TM["text_main"])
        self.btn_set_rules.configure(bg=TM["primary"])
        self.lbl_date_preview.configure(bg=TM["bg_card"], fg=TM["text_muted"])
        self.lbl_total.configure(bg=TM["bg_main"], fg=TM["text_muted"])

        for chk in [self.chk_date_update, self.chk_use_time_rules, self.chk_conversion, self.chk_use_manual_date]:
            chk.configure(
                bg=TM["bg_card"], activebackground=TM["bg_card"],
                fg=TM["text_main"], activeforeground=TM["text_main"], selectcolor=TM["bg_card"]
            )
        self.chk_add_bg.configure(
            bg=TM["bg_card"], activebackground=TM["bg_card"],
            fg=TM["danger"], activeforeground=TM["danger"], selectcolor=TM["bg_card"]
        )
        self.f_opt_r1.configure(bg=TM["bg_card"])
        self.f_opt_r2.configure(bg=TM["bg_card"])
        self.ent_manual_date.configure(bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"])
        self.btn_calendar.configure(bg=TM["bg_card"])

        self.update_mapping_ui()
        self.controller.set_tree_tags(self.tree)
        if self.btn_run.cget("state") == "normal":
            self.btn_run.configure(bg=TM["success"])

    def reset_to_defaults(self):
        super().reset_to_defaults()
        self.run_date_update.set(True)
        self.run_conversion.set(True)
        self.add_bg_var.set(True)
        self.use_time_rules.set(False)
        self.rule_t1.set("09:00")
        self.rule_t2.set("16:00")
        self.rule_x1.set(0)
        self.rule_x2.set(1)
        self.rule_x3.set(2)
        self.target_folder.set("Chưa chọn thư mục...")
        self.use_manual_date.set(False)
        self.manual_date_str.set(datetime.now().strftime('%d/%m/%Y'))
        self.update_preview()
        self.load_files_to_tree("")

    def load_settings(self):
        if not self.config_manager:
            return
        rules = self.config_manager.get_time_rules()
        self.rule_t1.set(rules.get("t1", "09:00"))
        self.rule_t2.set(rules.get("t2", "16:00"))
        self.rule_x1.set(rules.get("x1", 0))
        self.rule_x2.set(rules.get("x2", 1))
        self.rule_x3.set(rules.get("x3", 2))

        ws = self.config_manager.get_word_settings()
        self.run_date_update.set(ws.get("run_date_update", True))
        self.run_conversion.set(ws.get("run_conversion", True))
        self.add_bg_var.set(ws.get("add_bg_var", True))
        self.use_time_rules.set(ws.get("use_time_rules", False))
        self.use_manual_date.set(ws.get("use_manual_date", False))
        self.manual_date_str.set(ws.get("manual_date_str", datetime.now().strftime('%d/%m/%Y')))

    def save_settings(self):
        if not self.config_manager:
            return
        rules = {
            "t1": self.rule_t1.get(),
            "t2": self.rule_t2.get(),
            "x1": self.rule_x1.get(),
            "x2": self.rule_x2.get(),
            "x3": self.rule_x3.get()
        }
        self.config_manager.save_time_rules(rules)

        ws = {
            "run_date_update": self.run_date_update.get(),
            "run_conversion": self.run_conversion.get(),
            "add_bg_var": self.add_bg_var.get(),
            "use_time_rules": self.use_time_rules.get(),
            "use_manual_date": self.use_manual_date.get(),
            "manual_date_str": self.manual_date_str.get()
        }
        self.config_manager.save_word_settings(ws)

    def get_calculated_date(self) -> datetime:
        time_rules = {
            "t1": self.rule_t1.get(),
            "t2": self.rule_t2.get(),
            "x1": self.rule_x1.get(),
            "x2": self.rule_x2.get(),
            "x3": self.rule_x3.get()
        }
        return calculate_target_date(
            time_rules=time_rules,
            use_time_rules=self.use_time_rules.get(),
            use_manual_date=self.use_manual_date.get(),
            manual_date_str=self.manual_date_str.get()
        )

    def open_calendar(self):
        CalendarDialog(
            parent=self.controller.root,
            initial_date_str=self.manual_date_str.get(),
            on_select=lambda d_str: (
                self.manual_date_str.set(d_str),
                self.save_settings(),
                self.update_preview()
            )
        )

    def open_rule_settings(self):
        current_rules = {
            "t1": self.rule_t1.get(),
            "t2": self.rule_t2.get(),
            "x1": self.rule_x1.get(),
            "x2": self.rule_x2.get(),
            "x3": self.rule_x3.get()
        }
        def on_rules_saved(rules):
            self.rule_t1.set(rules.get("t1", "09:00"))
            self.rule_t2.set(rules.get("t2", "16:00"))
            self.rule_x1.set(rules.get("x1", 0))
            self.rule_x2.set(rules.get("x2", 1))
            self.rule_x3.set(rules.get("x3", 2))
            self.save_settings()
            self.update_preview()

        TimeRulesDialog(parent=self.controller.root, current_rules=current_rules, on_save=on_rules_saved)

    def open_signer_config(self):
        SignerConfigDialog(parent=self.controller.root, config_manager=self.config_manager)

    def update_preview(self):
        self.lbl_date_preview.config(text=f"(Sẽ dùng: {self.get_calculated_date().strftime('%d/%m/%Y')})")

    def open_tbkq_folder(self):
        folder_path = self.target_folder.get().strip()
        if not folder_path or "Chưa chọn" in folder_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục trước khi mở!")
            return
        if os.path.isdir(folder_path):
            try:
                os.startfile(os.path.normpath(folder_path))
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")
        else:
            messagebox.showwarning("Cảnh báo", "Thư mục không tồn tại! Vui lòng kiểm tra lại đường dẫn.")

    def copy_folder_path(self):
        folder_path = self.target_folder.get().strip()
        if not folder_path or "Chưa chọn" in folder_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục trước khi copy!")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(os.path.normpath(folder_path))
            orig_text, orig_bg = self.btn_copy_path.cget("text"), self.btn_copy_path.cget("background")
            self.btn_copy_path.config(text="✅ Đã copy!", bg=TM["success"])
            self.after(2000, lambda: self.btn_copy_path.config(text=orig_text, bg=orig_bg))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể copy đường dẫn: {e}")

    def on_folder_selected(self, folder_path: str):
        self.target_folder.set(folder_path)
        self.load_files_to_tree(folder_path)

    def select_folder(self):
        folder_selected = filedialog.askdirectory(title="Chọn thư mục chứa Word")
        if folder_selected:
            self.on_folder_selected(os.path.normpath(self.get_tbkq_subfolder(folder_selected)))

    def load_files_to_tree(self, folder: str):
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            if not folder:
                self.lbl_total.config(text="Tổng số file: 0")
                self.btn_run.config(state="disabled", bg="#cbd5e1")
                return

            files = [f for f in os.listdir(folder) if f.lower().endswith(('.doc', '.docx')) and not f.startswith("~$")]
            self.lbl_total.config(text=f"Tổng số file: {len(files)}")
            if not files:
                messagebox.showinfo("Thông báo", "Không tìm thấy file Word nào trong thư mục này.")
                self.btn_run.config(state="disabled", bg="#cbd5e1")
                return
            for fname in files:
                lang = detect_language(fname)
                self.tree.insert("", "end", iid=fname, values=(fname, lang, "Sẵn sàng", "-"), tags=("ready",))
            self.btn_run.config(state="normal", bg=TM["success"])
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc thư mục: {e}")

    def process(self):
        folder = self.target_folder.get()
        if not os.path.exists(folder) or folder == "Chưa chọn thư mục":
            return messagebox.showwarning("Lỗi", "Vui lòng chọn thư mục hợp lệ!")
        self.btn_run.config(state="disabled", text="ĐANG XỬ LÝ...", bg=TM["warning"])
        threading.Thread(target=self._run_logic, args=(folder,), daemon=True).start()

    def _run_logic(self, folder: str):
        pythoncom.CoInitialize()
        word = None
        futures = []
        self.batch_signers = None
        self.batch_lock = threading.Lock()
        target_dt = self.get_calculated_date()

        with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                items = self.tree.get_children()

                for idx_str, item_id in enumerate(items, 1):
                    fname = self.tree.item(item_id, 'values')[0]
                    lang = detect_language(fname)
                    fpath = os.path.abspath(os.path.join(folder, fname))

                    self.controller.root.after(
                        0, lambda iid=item_id, f=fname, l=lang: self.tree.item(iid, values=(f, l, "Đang xử lý...", "..."))
                    )

                    try:
                        doc = word.Documents.Open(os.path.abspath(fpath))
                        word_status = "➖ Bỏ qua"
                        if self.run_date_update.get():
                            word_status = WordEngine.update_word_date(doc, target_dt)

                        pdf_path = ""
                        if self.run_conversion.get():
                            pdf_path = os.path.splitext(fpath)[0] + ".pdf"
                            WordEngine.export_to_pdf(doc, pdf_path)

                        doc.Close(False)

                        if pdf_path and os.path.exists(pdf_path):
                            do_check = self.add_bg_var.get()

                            def pdf_done_callback(fut, iid=item_id, idx=idx_str, fn=fname, l=lang, ws=word_status):
                                try:
                                    overall_check, tag, detected = fut.result()

                                    with self.batch_lock:
                                        if self.batch_signers is None:
                                            if tag == "good" and detected[0] and detected[1]:
                                                self.batch_signers = detected
                                        else:
                                            if detected[0] and detected[1]:
                                                if (detected[0] != self.batch_signers[0] or detected[1] != self.batch_signers[1]):
                                                    signer_err = f"Chữ ký (KHÁC MẪU ({detected[0] or '?'}/{detected[1] or '?'}))"
                                                    if overall_check == "✅ OK":
                                                        overall_check = f"❌ {signer_err}"
                                                    elif "Chữ ký" not in overall_check:
                                                        overall_check += f", {signer_err}"
                                                    tag = "bad"

                                    self.controller.root.after(0, lambda: self._update_tree(iid, idx, fn, l, ws, overall_check, tag))
                                except Exception as e:
                                    self.controller.root.after(0, lambda: self._update_tree(iid, idx, fn, l, ws, f"❌ Lỗi PDF: {e}", "bad"))

                            future = executor.submit(
                                PdfService.process_pdf_post_export,
                                pdf_path,
                                bg_path_vn=self.controller.bg_pdf_path_vn,
                                bg_path_en=self.controller.bg_pdf_path_en,
                                add_bg=do_check,
                                check_date=do_check,
                                check_id=do_check,
                                check_page=do_check,
                                check_signer=do_check,
                                target_date=target_dt,
                                inspector_library=self.config_manager.inspector_library,
                                boss_library=self.config_manager.boss_library,
                                lang=lang
                            )
                            future.add_done_callback(pdf_done_callback)
                            futures.append(future)
                        else:
                            if self.run_conversion.get():
                                word_status = f"{word_status} (Lỗi xuất PDF)"
                                tag = "bad"
                                ov_check = "❌ Lỗi xuất PDF"
                            else:
                                tag = "good" if "✅" in word_status or "➖" in word_status else "bad"
                                ov_check = "➖ N/A"
                            self.controller.root.after(
                                0, lambda iid=item_id, i=idx_str, f=fname, l=lang, ws=word_status, ov=ov_check, t=tag:
                                self._update_tree(iid, i, f, l, ws, ov, t)
                            )

                    except Exception as e:
                        self.controller.root.after(
                            0, lambda iid=item_id, i=idx_str, f=fname, l=lang, err=str(e):
                            self._update_tree(iid, i, f, l, f"❌ Lỗi Word: {err}", "❌ Lỗi Word", "bad")
                        )

            except Exception as e:
                messagebox.showerror("Lỗi COM", str(e))
            finally:
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()

        self.controller.root.after(
            0, lambda: self.btn_run.config(state="normal", text="▶ BẮT ĐẦU XỬ LÝ WORD", bg=TM["success"])
        )
        self.controller.root.after(
            0, lambda: messagebox.showinfo("Hoàn tất", "Đã xử lý xong toàn bộ file Word và PDF!")
        )

    def _update_tree(self, iid, idx, fname, lang, word_stat, overall_check, tag):
        new_tags = (tag,) if tag else ()
        self.tree.item(iid, values=(fname, lang, word_stat, overall_check), tags=new_tags)


# Alias for backwards compatibility
WordProcessor = WordProcessorView
