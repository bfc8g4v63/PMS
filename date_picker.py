#$ date_picker.py
#% 日期視窗工具
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import calendar

class DatePickerDialog:
    def __init__(self, master, title, initial_date_str, on_selected):
        self.top = tk.Toplevel(master)
        self.top.title(title)
        self.top.transient(master)
        self.top.grab_set()
        self.on_selected = on_selected
        self.selected_date = None

        if initial_date_str:
            try:
                base_date = datetime.strptime(initial_date_str, "%Y-%m-%d")
            except ValueError:
                base_date = datetime.today()
        else:
            base_date = datetime.today()

        self.year = base_date.year
        self.month = base_date.month

        header = ttk.Frame(self.top)
        header.pack(fill="x", padx=5, pady=5)

        self.prev_button = ttk.Button(header, text="<", width=3, command=self.prev_month)
        self.prev_button.pack(side="left")

        self.month_label = ttk.Label(header, text="")
        self.month_label.pack(side="left", expand=True)

        self.next_button = ttk.Button(header, text=">", width=3, command=self.next_month)
        self.next_button.pack(side="right")

        self.days_frame = ttk.Frame(self.top)
        self.days_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.button_frame = ttk.Frame(self.top)
        self.button_frame.pack(fill="x", padx=5, pady=5)

        self.today_button = ttk.Button(self.button_frame, text="今天", command=self.select_today)
        self.today_button.pack(side="left", padx=5)

        self.clear_button = ttk.Button(self.button_frame, text="清除", command=self.clear_selection)
        self.clear_button.pack(side="left", padx=5)

        self.cancel_button = ttk.Button(self.button_frame, text="取消", command=self.close_without_select)
        self.cancel_button.pack(side="right", padx=5)

        self.draw_calendar()

    def draw_calendar(self):
        for widget in self.days_frame.winfo_children():
            widget.destroy()

        self.month_label.config(text=f"{self.year} - {self.month:02d}")

        week_days = ["一", "二", "三", "四", "五", "六", "日"]
        for idx, name in enumerate(week_days):
            lbl = ttk.Label(self.days_frame, text=name, anchor="center", width=4)
            lbl.grid(row=0, column=idx, padx=2, pady=2)

        month_days = calendar.monthcalendar(self.year, self.month)
        for r, week in enumerate(month_days, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    lbl = ttk.Label(self.days_frame, text="", width=4)
                    lbl.grid(row=r, column=c, padx=2, pady=2)
                else:
                    btn = ttk.Button(
                        self.days_frame,
                        text=f"{day:02d}",
                        width=4,
                        command=lambda d=day: self.on_day_clicked(d),
                    )
                    btn.grid(row=r, column=c, padx=2, pady=2)

    def on_day_clicked(self, day):
        self.selected_date = datetime(self.year, self.month, day)
        if self.on_selected:
            date_str = self.selected_date.strftime("%Y-%m-%d")
            self.on_selected(date_str)
        self.top.destroy()

    def prev_month(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self.draw_calendar()

    def next_month(self):
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self.draw_calendar()

    def select_today(self):
        today = datetime.today()
        self.year = today.year
        self.month = today.month
        self.draw_calendar()
        if self.on_selected:
            date_str = today.strftime("%Y-%m-%d")
            self.on_selected(date_str)
        self.top.destroy()

    def clear_selection(self):
        if self.on_selected:
            self.on_selected("")
        self.top.destroy()

    def close_without_select(self):
        self.top.destroy()

def open_date_picker(master, title, current_value, on_selected):
    DatePickerDialog(master, title, current_value, on_selected)