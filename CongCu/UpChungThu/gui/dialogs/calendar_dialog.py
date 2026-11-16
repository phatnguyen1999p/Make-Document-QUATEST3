"""
Calendar popup dialog for visual date selection.
"""

import tkinter as tk
import calendar
from datetime import datetime
from typing import Optional, Callable
from gui.theme import TM


class CalendarDialog:
    """
    Mini calendar popup dialog to select a date in dd/mm/yyyy format.
    """

    def __init__(self, parent, initial_date_str: str = "", on_select: Optional[Callable[[str], None]] = None):
        self.parent = parent
        self.on_select = on_select

        self.top = tk.Toplevel(self.parent)
        self.top.title("Chọn Ngày")
        self.top.geometry("260x280")
        self.top.configure(bg=TM["bg_card"])
        self.top.transient(self.parent)
        self.top.grab_set()

        try:
            self.current_date = datetime.strptime(initial_date_str.strip(), "%d/%m/%Y")
        except Exception:
            self.current_date = datetime.now()

        self.current_year = tk.IntVar(value=self.current_date.year)
        self.current_month = tk.IntVar(value=self.current_date.month)

        self._build_ui()

    def _build_ui(self):
        header_frame = tk.Frame(self.top, bg=TM["bg_card"])
        header_frame.pack(fill="x", pady=10)

        cal_frame = tk.Frame(self.top, bg=TM["bg_card"])
        cal_frame.pack(pady=5)

        def update_calendar():
            for widget in cal_frame.winfo_children():
                widget.destroy()

            y = self.current_year.get()
            m = self.current_month.get()
            lbl_month.config(text=f"Tháng {m} - {y}")

            days = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]
            for i, day in enumerate(days):
                tk.Label(
                    cal_frame, text=day, font=TM["font_small"],
                    bg=TM["bg_card"], fg=TM["primary"], width=4
                ).grid(row=0, column=i)

            cal_data = calendar.monthcalendar(y, m)
            for row_idx, week in enumerate(cal_data):
                for col_idx, day in enumerate(week):
                    if day != 0:
                        btn = tk.Button(
                            cal_frame, text=str(day), bg=TM["bg_card"],
                            fg=TM["text_main"], bd=0, font=TM["font_normal"], cursor="hand2"
                        )
                        if (day == self.current_date.day and
                            m == self.current_date.month and
                            y == self.current_date.year):
                            btn.config(bg=TM["primary"], fg="white")

                        def create_cmd(d, m_val, y_val):
                            return lambda: self._select_date(d, m_val, y_val)

                        btn.config(command=create_cmd(day, m, y))
                        btn.grid(row=row_idx + 1, column=col_idx, padx=2, pady=2)

        def prev_month():
            m = self.current_month.get()
            y = self.current_year.get()
            if m == 1:
                self.current_month.set(12)
                self.current_year.set(y - 1)
            else:
                self.current_month.set(m - 1)
            update_calendar()

        def next_month():
            m = self.current_month.get()
            y = self.current_year.get()
            if m == 12:
                self.current_month.set(1)
                self.current_year.set(y + 1)
            else:
                self.current_month.set(m + 1)
            update_calendar()

        tk.Button(
            header_frame, text="<", command=prev_month,
            bg=TM["bg_card"], fg=TM["text_main"], bd=0, font=TM["font_header"], cursor="hand2"
        ).pack(side="left", padx=15)

        lbl_month = tk.Label(header_frame, text="", bg=TM["bg_card"], font=TM["font_header"], fg=TM["text_main"])
        lbl_month.pack(side="left", expand=True)

        tk.Button(
            header_frame, text=">", command=next_month,
            bg=TM["bg_card"], fg=TM["text_main"], bd=0, font=TM["font_header"], cursor="hand2"
        ).pack(side="right", padx=15)

        update_calendar()

    def _select_date(self, day: int, month: int, year: int):
        formatted = f"{day:02d}/{month:02d}/{year}"
        if self.on_select:
            self.on_select(formatted)
        self.top.destroy()
