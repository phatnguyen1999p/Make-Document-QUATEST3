"""
Dialog for configuring quick selection and multi-sheet merging in Excel files.
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Callable
from gui.theme import TM


class BulkConfigDialog:
    """
    Modal window to quickly toggle sheet column selections and configure '+' sheet connectors.
    """

    def __init__(self, parent, all_sheets: List[str], bulk_selected_categories: Dict[str, tk.IntVar],
                 bulk_connectors: List[tk.IntVar], tree: ttk.Treeview, toggle_column_cb: Callable):
        self.parent = parent
        self.all_sheets = all_sheets
        self.bulk_selected_categories = bulk_selected_categories
        self.bulk_connectors = bulk_connectors
        self.tree = tree
        self.toggle_column_cb = toggle_column_cb

        self.win = tk.Toplevel(self.parent)
        self.win.title("Cấu hình chọn nhanh & Gộp sheet")
        self.win.geometry("650x550")
        self.win.configure(bg=TM["bg_card"])
        self.win.transient(self.parent)
        self.win.grab_set()

        self._build_ui()

    def _build_ui(self):
        tk.Label(
            self.win, text="CHỌN NHANH VÀ CẤU HÌNH GỘP",
            font=TM["font_header"], bg=TM["bg_card"], fg=TM["primary"]
        ).pack(pady=15)

        container = tk.Frame(self.win, bg=TM["bg_card"])
        container.pack(fill="both", expand=True, padx=20)

        canvas = tk.Canvas(container, bg=TM["bg_card"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=TM["bg_card"])

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Header row
        h_row = tk.Frame(scroll_frame, bg=TM["bg_main"], pady=5)
        h_row.pack(fill="x", pady=(0, 5))
        tk.Label(h_row, text="Tên Sheet", font=TM["font_header"], bg=TM["bg_main"], fg=TM["text_main"], width=35, anchor="w").pack(side="left", padx=5)
        tk.Label(h_row, text="Ghép với sheet sau (+)", font=TM["font_header"], bg=TM["bg_main"], fg=TM["text_main"], width=25).pack(side="left")

        # Rows per sheet
        for i, sheet_name in enumerate(self.all_sheets):
            row = tk.Frame(scroll_frame, bg=TM["bg_card"])
            row.pack(fill="x", pady=2)

            var = self.bulk_selected_categories[sheet_name]
            chk = tk.Checkbutton(
                row, text=f"Tất cả {sheet_name}", variable=var,
                bg=TM["bg_card"], activebackground=TM["bg_card"],
                fg=TM["text_main"], activeforeground=TM["text_main"],
                selectcolor=TM["bg_card"], font=TM["font_normal"],
                width=35, anchor="w",
                command=lambda s=sheet_name, v=var: self.toggle_column_cb(s, v)
            )
            chk.pack(side="left", padx=5)

            if i < len(self.all_sheets) - 1 and i < len(self.bulk_connectors):
                plus_var = self.bulk_connectors[i]
                chk_plus = tk.Checkbutton(
                    row, text="Gộp (+)", variable=plus_var,
                    bg=TM["bg_card"], activebackground=TM["bg_card"],
                    fg=TM["accent"], activeforeground=TM["accent"],
                    selectcolor=TM["bg_card"], font=TM["font_normal"]
                )
                chk_plus.pack(side="left", padx=20)

        # Bottom buttons
        btn_frame = tk.Frame(self.win, bg=TM["bg_card"])
        btn_frame.pack(fill="x", side="bottom", padx=30, pady=20)

        def deselect_all():
            for v in self.bulk_selected_categories.values():
                v.set(0)
            for v in self.bulk_connectors:
                v.set(0)

            num_sheets = len(self.all_sheets)
            for item_id in self.tree.get_children():
                vals = list(self.tree.item(item_id, 'values'))
                for j in range(2, num_sheets + 2):
                    if j < len(vals) and vals[j] != "":
                        vals[j] = "☐"
                self.tree.item(item_id, values=vals)

        tk.Button(
            btn_frame, text="XÓA TẤT CẢ CHỌN", command=deselect_all,
            bg=TM["danger"], fg="white", font=TM["font_header"],
            bd=0, pady=10, cursor="hand2"
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))

        tk.Button(
            btn_frame, text="HOÀN TẤT", command=self.win.destroy,
            bg=TM["success"], fg="white", font=TM["font_header"],
            bd=0, pady=10, cursor="hand2"
        ).pack(side="right", fill="x", expand=True, padx=(10, 0))
