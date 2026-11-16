"""
Base processor frame providing directory mapping and record search functionality.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, List
from gui.theme import TM
from core.mapping_service import MappingService


class MappingProcessor(tk.Frame):
    """
    Base view component with integrated search autocomplete and base archive mapping.
    """

    def __init__(self, parent, controller):
        super().__init__(parent, bg=TM["bg_main"])
        self.controller = controller
        self.base_search_dirs: List[str] = []
        self.folder_mapping: Dict[str, str] = {}

        self.load_mapping_paths()

    @property
    def config_manager(self):
        if hasattr(self.controller, 'config_manager'):
            return self.controller.config_manager
        return None

    def load_mapping_paths(self):
        """Load list of base directories from ConfigManager."""
        if self.config_manager:
            self.base_search_dirs = list(self.config_manager.mapping_paths)
        else:
            self.base_search_dirs = []

    def reset_to_defaults(self):
        """Reset mapping data and search widgets."""
        self.base_search_dirs = []
        self.folder_mapping = {}
        if hasattr(self, 'lbl_mapping_path'):
            self.lbl_mapping_path.config(text="⚠️ Chưa có thư mục gốc", fg=TM["danger"])
        if hasattr(self, 'search_var'):
            self.search_var.set("")
        if hasattr(self, 'suggestion_box'):
            self.suggestion_box.place_forget()

    def create_search_widgets(self, container):
        self.f_search = tk.Frame(container, bg=TM["bg_card"])
        self.f_search.pack(fill="x", pady=(5, 15))

        self.lbl_search_icon = tk.Label(
            self.f_search, text="🔍 Tìm hồ sơ:", bg=TM["bg_card"],
            fg=TM["text_main"], font=TM["font_header"]
        )
        self.lbl_search_icon.pack(side="left", padx=(0, 10))

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            self.f_search, textvariable=self.search_var, font=TM["font_normal"],
            width=35, bd=1, relief="solid", highlightthickness=0,
            bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"]
        )
        self.search_entry.pack(side="left", ipady=3)

        self.btn_refresh = tk.Button(
            self.f_search, text="🔄 Làm mới Data", command=self.scan_base_directory,
            bg=TM["primary"], fg="white", font=TM["font_small"], bd=0, cursor="hand2", padx=12, pady=4
        )
        self.btn_refresh.pack(side="left", padx=10)

        self.btn_add_root = tk.Button(
            self.f_search, text="➕ Thêm gốc", command=self.select_base_mapping_dir,
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, padx=12, pady=4, cursor="hand2"
        )
        self.btn_add_root.pack(side="left")

        self.lbl_mapping_path = tk.Label(
            self.f_search, text="Đang khởi tạo...", bg=TM["bg_card"],
            fg=TM["text_muted"], font=TM["font_small"]
        )
        self.lbl_mapping_path.pack(side="left", padx=15)

        # Hover effects
        self.btn_refresh.bind("<Enter>", lambda e: self.btn_refresh.config(bg="#2563eb"))
        self.btn_refresh.bind("<Leave>", lambda e: self.btn_refresh.config(bg=TM["primary"]))
        self.btn_add_root.bind("<Enter>", lambda e: self.btn_add_root.config(bg="#4f46e5"))
        self.btn_add_root.bind("<Leave>", lambda e: self.btn_add_root.config(bg=TM["accent"]))

        self.suggestion_box = tk.Listbox(
            self, font=TM["font_normal"], height=6, bg=TM["bg_card"],
            fg=TM["text_main"], selectbackground=TM["primary"], bd=1, relief="solid"
        )
        self.search_entry.bind('<KeyRelease>', self.on_search_type)
        self.search_entry.bind('<Down>', self.on_entry_down)
        self.search_entry.bind('<Return>', self.on_search_select)
        self.search_entry.bind('<FocusIn>', self.on_entry_focus)
        self.search_entry.bind('<FocusOut>', self.on_focus_out)
        self.suggestion_box.bind('<FocusOut>', self.on_focus_out)
        self.suggestion_box.bind('<ButtonRelease-1>', self.on_suggestion_select)
        self.suggestion_box.bind('<Return>', self.on_suggestion_select)

        if self.base_search_dirs:
            self.scan_base_directory()

    def update_mapping_ui(self):
        if hasattr(self, 'f_search'):
            self.f_search.config(bg=TM["bg_card"])
            self.lbl_search_icon.config(bg=TM["bg_card"], fg=TM["text_main"])
            self.search_entry.config(bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"])
            self.btn_refresh.config(bg=TM["primary"])
            self.btn_add_root.config(bg=TM["accent"])
            self.lbl_mapping_path.config(bg=TM["bg_card"])
            self.suggestion_box.config(bg=TM["bg_card"], fg=TM["text_main"], selectbackground=TM["primary"])

    def select_base_mapping_dir(self):
        selected_dir = filedialog.askdirectory(title="Chọn thư mục chứa các hồ sơ để Mapping")
        if selected_dir:
            norm_path = os.path.normpath(selected_dir)
            if self.config_manager:
                self.config_manager.add_mapping_path(norm_path)
                self.load_mapping_paths()
            elif norm_path not in self.base_search_dirs:
                self.base_search_dirs.append(norm_path)
            self.scan_base_directory()

    def scan_base_directory(self):
        self.folder_mapping.clear()
        if not self.base_search_dirs:
            self.lbl_mapping_path.config(text="⚠️ Chưa có thư mục gốc", fg=TM["danger"])
            return
        self.lbl_mapping_path.config(text="⏳ Đang map dữ liệu...", fg=TM["warning"])
        self.update_idletasks()
        try:
            self.folder_mapping, elapsed_ms = MappingService.scan_base_directories(self.base_search_dirs)
            self.lbl_mapping_path.config(
                text=f"✅ Đã map ({elapsed_ms} ms): {len(self.base_search_dirs)} gốc ({len(self.folder_mapping)} hồ sơ)",
                fg=TM["success"]
            )
        except Exception as e:
            self.lbl_mapping_path.config(text=f"❌ Lỗi: {e}", fg=TM["danger"])

    def on_search_type(self, event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape'):
            return
        typed = self.search_var.get().lower()
        if typed == "":
            self.suggestion_box.place_forget()
            return
        hits = MappingService.search_records(self.folder_mapping, typed)
        if hits:
            self.suggestion_box.delete(0, tk.END)
            for h in hits:
                self.suggestion_box.insert(tk.END, h)
            self.suggestion_box.place(in_=self.search_entry, x=0, rely=1.0, relwidth=1.0, y=1)
            self.suggestion_box.lift()
        else:
            self.suggestion_box.place_forget()

    def on_suggestion_select(self, event=None):
        if not self.suggestion_box.curselection():
            return
        selected_key = self.suggestion_box.get(self.suggestion_box.curselection())
        self.search_var.set(selected_key)
        self.suggestion_box.place_forget()
        self.on_search_select()

    def on_search_select(self, event=None):
        selected_key = self.search_var.get().strip()
        if selected_key in self.folder_mapping:
            base_path = self.folder_mapping[selected_key]
            final_path = self.get_tbkq_subfolder(base_path)
            self.on_folder_selected(os.path.normpath(final_path))
            self.focus()

    def get_tbkq_subfolder(self, parent_path: str) -> str:
        return MappingService.get_tbkq_subfolder(parent_path)

    def on_entry_down(self, event):
        if self.suggestion_box.winfo_ismapped() and self.suggestion_box.size() > 0:
            self.suggestion_box.focus_set()
            self.suggestion_box.selection_set(0)

    def on_entry_focus(self, event):
        self.controller.root.after(50, lambda: self.search_entry.selection_range(0, tk.END))

    def on_focus_out(self, event):
        self.controller.root.after(100, self._check_hide_box)

    def _check_hide_box(self):
        focused = self.controller.root.focus_get()
        if focused != self.search_entry and focused != self.suggestion_box:
            self.suggestion_box.place_forget()

    def on_folder_selected(self, folder_path: str):
        """Must be implemented by subclasses."""
        raise NotImplementedError
