"""
Main Application Controller (UnifiedDocumentSystem) managing windows,
styling, navigation routing, and modal dialogs.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict

from gui.theme import TM
from gui.views.main_menu import MainMenu
from gui.views.word_view import WordProcessorView
from gui.views.excel_view import ExcelProcessorView
from gui.dialogs.signer_dialog import SignerConfigDialog
from core.config_manager import ConfigManager
from core.pdf_service import PdfService, detect_language


APP_VERSION = "6.8.1"
APP_TITLE = f"QUATEST 3 - HỆ THỐNG XỬ LÝ VĂN BẢN - ver: {APP_VERSION}"


class UnifiedDocumentSystem:
    """
    Main Tkinter application controller.
    """

    def __init__(self, root: tk.Tk):
        self.root = root

        # Base application directory resolution
        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
            base_dir = getattr(sys, '_MEIPASS', self.app_dir)
        else:
            self.app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            base_dir = self.app_dir

        # Initialize Config Manager
        self.config_manager = ConfigManager(self.app_dir)
        TM.config_manager = self.config_manager
        TM.load_settings()

        self.root.title(APP_TITLE)
        self.root.geometry("1250x850")
        self.root.configure(bg=TM["bg_main"])

        # Template Backgrounds
        self.bg_pdf_path_en = os.path.join(base_dir, "GIAY TIEU DE - E.pdf")
        if not os.path.exists(self.bg_pdf_path_en) and os.path.exists(os.path.join(self.app_dir, "GIAY TIEU DE - E.pdf")):
            self.bg_pdf_path_en = os.path.join(self.app_dir, "GIAY TIEU DE - E.pdf")

        vn_candidate = os.path.join(base_dir, "GIAY TIEU DE .pdf")
        if not os.path.exists(vn_candidate):
            vn_candidate = os.path.join(base_dir, "GIAY TIEU DE.pdf")
        if not os.path.exists(vn_candidate) and os.path.exists(os.path.join(self.app_dir, "GIAY TIEU DE.pdf")):
            vn_candidate = os.path.join(self.app_dir, "GIAY TIEU DE.pdf")
        self.bg_pdf_path_vn = vn_candidate
        self.bg_pdf_path = self.bg_pdf_path_vn

        # Warning for missing PDF background templates
        missing = []
        if not os.path.exists(self.bg_pdf_path_vn):
            missing.append("'GIAY TIEU DE.pdf'")
        if not os.path.exists(self.bg_pdf_path_en):
            missing.append("'GIAY TIEU DE - E.pdf'")
        if missing:
            messagebox.showwarning(
                "Cảnh báo hệ thống",
                f"Không tìm thấy file mẫu nền: {', '.join(missing)}\nĐường dẫn thư mục: {base_dir}\n\n"
                "Tính năng 'Chèn Giấy Tiêu Đề' sẽ báo lỗi nếu thiếu file nền tương ứng."
            )

        self.setup_styles()

        # Main Container
        self.container = tk.Frame(self.root, bg=TM["bg_main"])
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # Theme Toggle Button (top-right)
        btn_text = "☀️ Light Mode" if TM.is_dark_mode else "🌙 Dark Mode"
        self.btn_theme = tk.Button(
            self.root, text=btn_text, command=self.toggle_theme,
            font=TM["font_small"], bd=0, padx=10, pady=5, cursor="hand2",
            bg=TM["bg_card"], fg=TM["text_main"]
        )
        self.btn_theme.place(relx=1.0, x=-20, y=20, anchor="ne")

        # Register Application Frames
        self.frames: Dict[str, tk.Frame] = {}
        for F, page_name in [
            (MainMenu, "MainMenu"),
            (WordProcessorView, "WordProcessor"),
            (ExcelProcessorView, "ExcelProcessor")
        ]:
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.current_frame_name = "MainMenu"
        self.show_frame("MainMenu")

    def show_frame(self, page_name: str):
        """Navigate to a registered view screen."""
        self.current_frame_name = page_name
        frame = self.frames[page_name]
        frame.tkraise()

    def toggle_theme(self):
        """Toggle Dark / Light theme across all views."""
        TM.toggle()
        self.btn_theme.config(
            text="☀️ Light Mode" if TM.is_dark_mode else "🌙 Dark Mode",
            bg=TM["bg_card"], fg=TM["text_main"]
        )
        self.root.configure(bg=TM["bg_main"])
        self.container.configure(bg=TM["bg_main"])
        self.setup_styles()

        for frame in self.frames.values():
            if hasattr(frame, "update_theme"):
                frame.update_theme()

    def setup_styles(self):
        """Configure TTK widget styles for light/dark mode."""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "Treeview",
            rowheight=45,
            font=TM["font_tree"],
            background=TM["bg_card"],
            fieldbackground=TM["bg_card"],
            foreground=TM["text_main"],
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            font=TM["font_header"],
            background=TM["border"],
            foreground=TM["text_main"],
            relief="flat"
        )
        style.map("Treeview.Heading", background=[('active', TM["primary"])])
        style.map(
            'Treeview',
            background=[('selected', TM["primary"])],
            foreground=[('selected', '#ffffff')]
        )
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        style.configure("TProgressbar", thickness=10, background=TM["primary"])

    def set_tree_tags(self, tree: ttk.Treeview):
        """Configure tag color schemes on a Treeview widget."""
        tree.tag_configure(
            "ready",
            background="#e0e7ff" if not TM.is_dark_mode else "#312e81",
            foreground="#3730a3" if not TM.is_dark_mode else "#c7d2fe"
        )
        tree.tag_configure(
            "good",
            background="#065f46" if TM.is_dark_mode else "#22c55e",
            foreground="#ffffff"
        )
        tree.tag_configure(
            "bad",
            background="#991b1b" if TM.is_dark_mode else "#ef4444",
            foreground="#ffffff"
        )
        tree.tag_configure(
            "warning",
            background="#92400e" if TM.is_dark_mode else "#f59e0b",
            foreground="#ffffff"
        )

    def open_signer_config(self):
        """Open the Signer Configuration modal dialog."""
        SignerConfigDialog(parent=self.root, config_manager=self.config_manager)

    def detect_language(self, file_name_or_path: str, sheet_names=None) -> str:
        """Proxy method for language detection."""
        return detect_language(file_name_or_path, sheet_names)

    def process_pdf_post_export(self, pdf_path: str, add_bg=True, check_date=True, check_id=True,
                                check_page=True, check_signer=True, target_date=None,
                                expected_insp=None, expected_boss=None, lang=None):
        """Proxy method for post-export PDF processing."""
        return PdfService.process_pdf_post_export(
            pdf_path=pdf_path,
            bg_path_vn=self.bg_pdf_path_vn,
            bg_path_en=self.bg_pdf_path_en,
            add_bg=add_bg,
            check_date=check_date,
            check_id=check_id,
            check_page=check_page,
            check_signer=check_signer,
            target_date=target_date,
            inspector_library=self.config_manager.inspector_library,
            boss_library=self.config_manager.boss_library,
            expected_insp=expected_insp,
            expected_boss=expected_boss,
            lang=lang
        )
