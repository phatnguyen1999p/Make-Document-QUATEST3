"""
Dialog for configuring time-based document date offset rules.
"""

import tkinter as tk
from typing import Optional, Callable, Dict, Any
from gui.theme import TM


class TimeRulesDialog:
    """
    Modal window for setting time thresholds (t1, t2) and offset day counts (x1, x2, x3).
    """

    def __init__(self, parent, current_rules: Dict[str, Any], on_save: Optional[Callable[[Dict[str, Any]], None]] = None):
        self.parent = parent
        self.on_save = on_save

        self.win = tk.Toplevel(self.parent)
        self.win.title("Cấu hình quy tắc giờ")
        self.win.geometry("420x320")
        self.win.configure(bg=TM["bg_card"])
        self.win.transient(self.parent)
        self.win.grab_set()

        self.rule_t1 = tk.StringVar(value=str(current_rules.get("t1", "09:00")))
        self.rule_t2 = tk.StringVar(value=str(current_rules.get("t2", "16:00")))
        self.rule_x1 = tk.IntVar(value=int(current_rules.get("x1", 0)))
        self.rule_x2 = tk.IntVar(value=int(current_rules.get("x2", 1)))
        self.rule_x3 = tk.IntVar(value=int(current_rules.get("x3", 2)))

        self._build_ui()

    def _build_ui(self):
        container = tk.Frame(self.win, bg=TM["bg_card"], padx=25, pady=20)
        container.pack(fill="both", expand=True)

        tk.Label(
            container, text="CÀI ĐẶT QUY TẮC GIỜ",
            font=TM["font_header"], bg=TM["bg_card"], fg=TM["primary"]
        ).pack(pady=(0, 15))

        def add_combined_row(parent_f, lbl_t, var_t, lbl_x, var_x, is_label=False):
            row = tk.Frame(parent_f, bg=TM["bg_card"])
            row.pack(fill="x", pady=10)

            tk.Label(row, text=lbl_t, bg=TM["bg_card"], font=TM["font_normal"], width=8, anchor="w").pack(side="left")
            if var_t:
                if is_label:
                    tk.Label(row, textvariable=var_t, font=TM["font_normal"], width=8, bg=TM["bg_card"], anchor="center").pack(side="left")
                else:
                    tk.Entry(
                        row, textvariable=var_t, font=TM["font_normal"], width=8, justify="center",
                        bd=1, relief="solid", bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"]
                    ).pack(side="left")
            else:
                tk.Label(row, text="", width=8, bg=TM["bg_card"]).pack(side="left")

            tk.Label(row, text=lbl_x, bg=TM["bg_card"], font=TM["font_normal"], width=10, anchor="e").pack(side="left", padx=(15, 0))
            tk.Entry(
                row, textvariable=var_x, font=TM["font_normal"], width=5, justify="center",
                bd=1, relief="solid", bg=TM["bg_main"], fg=TM["text_main"], insertbackground=TM["text_main"]
            ).pack(side="left", padx=5)
            tk.Label(row, text="ngày", bg=TM["bg_card"], font=TM["font_small"], fg=TM["text_muted"]).pack(side="left")

        add_combined_row(container, "Trước:", self.rule_t1, "Cộng:", self.rule_x1)
        add_combined_row(container, "Trước:", self.rule_t2, "Cộng:", self.rule_x2)
        add_combined_row(container, "Sau:", self.rule_t2, "Cộng:", self.rule_x3, is_label=True)

        tk.Label(
            container, text="* Mốc thời gian theo định dạng HH:MM (VD: 09:00)",
            bg=TM["bg_card"], font=TM["font_small"], fg=TM["text_muted"]
        ).pack(pady=(10, 0))

        def apply_and_save():
            rules = {
                "t1": self.rule_t1.get().strip(),
                "t2": self.rule_t2.get().strip(),
                "x1": self.rule_x1.get(),
                "x2": self.rule_x2.get(),
                "x3": self.rule_x3.get()
            }
            if self.on_save:
                self.on_save(rules)
            self.win.destroy()

        tk.Button(
            self.win, text="LƯU VÀ ĐÓNG", command=apply_and_save,
            bg=TM["primary"], fg="white", font=TM["font_header"], bd=0, pady=10, cursor="hand2"
        ).pack(fill="x", padx=30, pady=20)
