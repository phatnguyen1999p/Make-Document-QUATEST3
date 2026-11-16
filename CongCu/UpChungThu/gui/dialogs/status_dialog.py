"""
Progress and status dialog for long-running operations like folder scanning.
"""

import tkinter as tk
from tkinter import ttk
from gui.theme import TM


class StatusProgressDialog:
    """
    Modal progress window displaying current file progress and a determinate progress bar.
    """

    def __init__(self, parent, title: str):
        self.parent = parent
        self.win = tk.Toplevel(self.parent)
        self.win.title(title)
        self.win.geometry("500x180")
        self.win.configure(bg=TM["bg_card"])
        self.win.attributes('-topmost', True)

        try:
            x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - 250
            y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - 90
            self.win.geometry(f"+{x}+{y}")
        except Exception:
            pass

        self._build_ui(title)

    def _build_ui(self, title: str):
        tk.Label(
            self.win, text=title, font=TM["font_header"],
            bg=TM["bg_card"], pady=15
        ).pack()

        self.lbl_current = tk.Label(
            self.win, text="Khởi tạo...", bg=TM["bg_card"],
            fg=TM["text_muted"], wraplength=450
        )
        self.lbl_current.pack(pady=5)

        self.bar = ttk.Progressbar(self.win, orient="horizontal", length=400, mode='determinate')
        self.bar.pack(pady=10)

    def set_maximum(self, maximum: int):
        self.bar['maximum'] = maximum

    def update_progress(self, text: str, value: int):
        self.lbl_current.config(text=text)
        self.bar.config(value=value)

    def close(self):
        try:
            self.win.destroy()
        except Exception:
            pass
