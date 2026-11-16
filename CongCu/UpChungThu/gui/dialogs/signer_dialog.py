"""
Dialog for managing signer library (GDV and Boss).
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable
from gui.theme import TM


class SignerConfigDialog:
    """
    Modal window for configuring the library of inspectors (GDV) and supervisors (Boss).
    """

    def __init__(self, parent, config_manager, on_save: Optional[Callable] = None):
        self.parent = parent
        self.config_manager = config_manager
        self.on_save = on_save

        self.win = tk.Toplevel(self.parent)
        self.win.title("Quản lý Thư viện Người ký")
        self.win.geometry("700x550")
        self.win.configure(bg=TM["bg_card"])
        self.win.transient(self.parent)
        self.win.grab_set()

        self._build_ui()

    def _build_ui(self):
        main_container = tk.Frame(self.win, bg=TM["bg_card"], padx=25, pady=20)
        main_container.pack(fill="both", expand=True)

        tk.Label(
            main_container,
            text="THƯ VIỆN NGƯỜI KÝ",
            font=TM["font_header"],
            bg=TM["bg_card"],
            fg=TM["primary"]
        ).pack(pady=(0, 15))

        # --- Section: Add New ---
        add_frame = tk.Frame(main_container, bg=TM["bg_main"], padx=15, pady=15, bd=1, relief="solid")
        add_frame.pack(fill="x", pady=(0, 20))

        tk.Label(add_frame, text="Nhập tên mới:", bg=TM["bg_main"], font=TM["font_normal"]).pack(side="left")
        new_name_var = tk.StringVar()
        ent_name = tk.Entry(add_frame, textvariable=new_name_var, font=TM["font_normal"], width=30)
        ent_name.pack(side="left", padx=10, fill="x", expand=True)

        def add_to(role):
            name = new_name_var.get().strip()
            if not name:
                return
            if self.config_manager.add_signer(name, role):
                new_name_var.set("")
                refresh_lists()

        tk.Button(
            add_frame, text="+ GDV", command=lambda: add_to("gdv"),
            bg=TM["primary"], fg="white", font=TM["font_small"], bd=0, padx=10, cursor="hand2"
        ).pack(side="left", padx=2)

        tk.Button(
            add_frame, text="+ Boss", command=lambda: add_to("boss"),
            bg=TM["accent"], fg="white", font=TM["font_small"], bd=0, padx=10, cursor="hand2"
        ).pack(side="left", padx=2)

        # --- Section: Signer Lists ---
        lists_frame = tk.Frame(main_container, bg=TM["bg_card"])
        lists_frame.pack(fill="both", expand=True)
        lists_frame.columnconfigure(0, weight=1, uniform="group1")
        lists_frame.columnconfigure(1, weight=1, uniform="group1")
        lists_frame.rowconfigure(0, weight=1)

        def create_list_area(parent_f, col, title, color):
            frame = tk.Frame(parent_f, bg=TM["bg_card"], bd=1, relief="solid")
            frame.grid(row=0, column=col, sticky="nsew", padx=5)

            tk.Label(frame, text=title, bg=color, fg="white", font=TM["font_header"], pady=5).pack(fill="x")

            canvas = tk.Canvas(frame, bg=TM["bg_card"], highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=TM["bg_card"])

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            return scrollable_frame

        gdv_list_frame = create_list_area(lists_frame, 0, "DANH SÁCH GDV", TM["primary"])
        boss_list_frame = create_list_area(lists_frame, 1, "DANH SÁCH BOSS", TM["accent"])

        def remove_name(name, role):
            if self.config_manager.remove_signer(name, role):
                refresh_lists()

        def refresh_lists():
            for w in gdv_list_frame.winfo_children():
                w.destroy()
            for w in boss_list_frame.winfo_children():
                w.destroy()

            for name in sorted(self.config_manager.inspector_library):
                row = tk.Frame(gdv_list_frame, bg=TM["bg_card"])
                row.pack(fill="x", pady=2)
                tk.Label(row, text=f"• {name}", bg=TM["bg_card"], font=TM["font_normal"], anchor="w").pack(side="left", fill="x", expand=True)
                tk.Button(row, text="🗑", command=lambda n=name: remove_name(n, "gdv"), bg=TM["bg_card"], fg=TM["danger"], bd=0, cursor="hand2").pack(side="right")

            for name in sorted(self.config_manager.boss_library):
                row = tk.Frame(boss_list_frame, bg=TM["bg_card"])
                row.pack(fill="x", pady=2)
                tk.Label(row, text=f"• {name}", bg=TM["bg_card"], font=TM["font_normal"], anchor="w").pack(side="left", fill="x", expand=True)
                tk.Button(row, text="🗑", command=lambda n=name: remove_name(n, "boss"), bg=TM["bg_card"], fg=TM["danger"], bd=0, cursor="hand2").pack(side="right")

        refresh_lists()

        tk.Label(
            main_container,
            text="* Hệ thống sẽ chốt 2 người ở file đầu tiên làm mẫu cho cả folder.",
            bg=TM["bg_card"], font=TM["font_small"], fg=TM["text_muted"], justify="left"
        ).pack(pady=10)

        def save_and_close():
            self.config_manager.save_settings()
            if self.on_save:
                self.on_save()
            self.win.destroy()

        tk.Button(
            self.win, text="LƯU VÀ ĐÓNG", command=save_and_close,
            bg=TM["success"], fg="white", font=TM["font_header"], bd=0, pady=10, cursor="hand2"
        ).pack(fill="x", padx=25, pady=(0, 20))
