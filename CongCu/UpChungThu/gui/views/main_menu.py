"""
Main Menu view for navigation and data source configuration.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from gui.theme import TM


class MainMenu(tk.Frame):
    """
    Main entry navigation view.
    """

    def __init__(self, parent, controller):
        super().__init__(parent, bg=TM["bg_main"])
        self.controller = controller

        self.card = tk.Frame(
            self, bg=TM["bg_card"], padx=60, pady=60, bd=0,
            highlightbackground=TM["border"], highlightthickness=1
        )
        self.card.place(relx=0.5, rely=0.5, anchor="center")

        btn_style = {
            "font": TM["font_header"], "width": 32, "height": 2,
            "cursor": "hand2", "bd": 0, "relief": "flat", "pady": 5
        }

        self.btn_excel = tk.Button(
            self.card, text="📊 THAO TÁC VỚI FILE EXCEL", bg=TM["primary"], fg="white",
            command=lambda: controller.show_frame("ExcelProcessor"), **btn_style
        )
        self.btn_excel.pack(pady=10)

        self.btn_word = tk.Button(
            self.card, text="📝 THAO TÁC VỚI FILE WORD", bg=TM["success"], fg="white",
            command=lambda: controller.show_frame("WordProcessor"), **btn_style
        )
        self.btn_word.pack(pady=10)

        self.btn_reset = tk.Button(
            self.card, text="🗑️ KHÔI PHỤC CÀI ĐẶT MẶC ĐỊNH", bg=TM["danger"], fg="white",
            command=self.reset_all_settings, font=TM["font_small"], bd=0, cursor="hand2", padx=10, pady=5
        )
        self.btn_reset.pack(pady=(20, 0))

        # Data Configuration Box (LAN/Local)
        self.config_box = tk.LabelFrame(
            self.card, text="🔌 CẤU HÌNH DỮ LIỆU", bg=TM["bg_card"],
            fg=TM["primary"], font=TM["font_small"], padx=15, pady=10
        )
        self.config_box.pack(fill="x", pady=(20, 0))

        self.use_lan_var = tk.BooleanVar(value=False)
        self.lan_path_var = tk.StringVar(value="")

        if self.controller.config_manager:
            self.use_lan_var.set(self.controller.config_manager.use_lan)
            self.lan_path_var.set(self.controller.config_manager.lan_path)

        row1 = tk.Frame(self.config_box, bg=TM["bg_card"])
        row1.pack(fill="x")
        self.rb_local = tk.Radiobutton(
            row1, text="Máy này (Local)", variable=self.use_lan_var, value=False,
            bg=TM["bg_card"], fg=TM["text_main"], selectcolor=TM["bg_card"],
            activebackground=TM["bg_card"], command=self.save_bootstrap
        )
        self.rb_local.pack(side="left")

        self.rb_lan = tk.Radiobutton(
            row1, text="Dùng chung (LAN)", variable=self.use_lan_var, value=True,
            bg=TM["bg_card"], fg=TM["text_main"], selectcolor=TM["bg_card"],
            activebackground=TM["bg_card"], command=self.save_bootstrap
        )
        self.rb_lan.pack(side="left", padx=10)

        row2 = tk.Frame(self.config_box, bg=TM["bg_card"])
        row2.pack(fill="x", pady=(5, 0))
        self.ent_lan_path = tk.Entry(
            row2, textvariable=self.lan_path_var, font=TM["font_small"],
            state="readonly", bd=1, relief="solid", bg=TM["bg_main"], fg=TM["text_main"]
        )
        self.ent_lan_path.pack(side="left", fill="x", expand=True)

        tk.Button(
            row2, text="📁", command=self.browse_lan_file,
            bg=TM["bg_card"], fg=TM["text_main"], bd=0, cursor="hand2"
        ).pack(side="left", padx=5)

    def update_theme(self):
        self.configure(bg=TM["bg_main"])
        self.card.configure(bg=TM["bg_card"], highlightbackground=TM["border"])
        self.btn_excel.configure(bg=TM["primary"], fg="white")
        self.btn_word.configure(bg=TM["success"], fg="white")
        self.btn_reset.configure(bg=TM["danger"], fg="white")

        if hasattr(self, 'config_box'):
            self.config_box.configure(bg=TM["bg_card"], fg=TM["primary"])
            for child in self.config_box.winfo_children():
                if isinstance(child, tk.Frame):
                    child.configure(bg=TM["bg_card"])
                    for sub in child.winfo_children():
                        if isinstance(sub, tk.Radiobutton):
                            sub.configure(
                                bg=TM["bg_card"], fg=TM["text_main"],
                                selectcolor=TM["bg_card"], activebackground=TM["bg_card"]
                            )
                        elif isinstance(sub, tk.Label):
                            sub.configure(bg=TM["bg_card"], fg=TM["text_main"])
                        elif isinstance(sub, tk.Entry):
                            sub.configure(bg=TM["bg_main"], fg=TM["text_main"])
                        elif isinstance(sub, tk.Button):
                            sub.configure(bg=TM["bg_card"], fg=TM["text_main"])

    def reset_all_settings(self):
        """Reset all application settings to defaults."""
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ cài đặt (thư mục gốc, quy tắc giờ,...) và quay về mặc định?"):
            self.controller.config_manager.reset_all_settings()

            # Reset Theme
            TM.is_dark_mode = False
            TM.load_settings()
            self.controller.setup_styles()
            self.controller.btn_theme.config(
                text="🌙 Dark Mode", bg=TM["bg_card"], fg=TM["text_main"]
            )
            self.controller.root.configure(bg=TM["bg_main"])
            self.controller.container.configure(bg=TM["bg_main"])

            for frame in self.controller.frames.values():
                if hasattr(frame, 'reset_to_defaults'):
                    frame.reset_to_defaults()
                if hasattr(frame, 'update_theme'):
                    frame.update_theme()

            self.use_lan_var.set(False)
            self.lan_path_var.set("")

            messagebox.showinfo("Hoàn tất", "Đã khôi phục cài đặt mặc định thành công!")

    def browse_lan_file(self):
        f = filedialog.askopenfilename(
            title="Chọn file settings dùng chung (.json)",
            filetypes=[("JSON files", "*.json")]
        )
        if f:
            self.lan_path_var.set(f)
            self.use_lan_var.set(True)
            self.save_bootstrap()
            messagebox.showinfo(
                "Cấu hình",
                f"Đã chuyển sang dùng thư viện tại:\n{f}\n\nLưu ý: Nếu file này chưa tồn tại signer library, bạn hãy tạo mới trong mục cấu hình chữ ký."
            )

    def save_bootstrap(self):
        try:
            self.controller.config_manager.save_bootstrap(
                use_lan=self.use_lan_var.get(),
                lan_path=self.lan_path_var.get()
            )
            for frame in self.controller.frames.values():
                if hasattr(frame, "load_mapping_paths"):
                    frame.load_mapping_paths()
                if hasattr(frame, "load_settings"):
                    frame.load_settings()
                if hasattr(frame, "load_excel_settings"):
                    frame.load_excel_settings()
        except Exception as e:
            messagebox.showerror("Lỗi Lưu", f"Không thể lưu cấu hình bootstrap: {e}")
