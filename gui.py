
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from zoneinfo import ZoneInfo


TIMEZONE = ZoneInfo("America/New_York")

WEEKDAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]


class SearchReportGUI:
    def __init__(self, root, generate_callback):
        self.root = root
        self.generate_callback = generate_callback

        self.root.title("Search Report Generator")
        self.root.geometry("420x430")
        self.root.resizable(False, False)

        self.employee_count = tk.StringVar()
        self.hourly_rate = tk.StringVar()
        self.search_minutes = tk.StringVar()

        # Default to today's weekday
        today = datetime.now(TIMEZONE)
        self.search_day = tk.StringVar(
            value=WEEKDAYS[today.weekday()]
        )

        self.build_gui()

    def build_gui(self):
        """Create the application interface."""

        main_frame = ttk.Frame(self.root, padding=25)
        main_frame.pack(fill="both", expand=True)

        title = ttk.Label(
            main_frame,
            text="Search Report Generator",
            font=("Segoe UI", 17, "bold")
        )
        title.pack(pady=(0, 20))

        # Day of the week
        ttk.Label(
            main_frame,
            text="Day of the Week:"
        ).pack(anchor="w")

        ttk.Combobox(
            main_frame,
            textvariable=self.search_day,
            values=WEEKDAYS,
            state="readonly"
        ).pack(fill="x", pady=(5, 15))

        # Number of employees
        ttk.Label(
            main_frame,
            text="Number of Employees:"
        ).pack(anchor="w")

        ttk.Entry(
            main_frame,
            textvariable=self.employee_count
        ).pack(fill="x", pady=(5, 15))

        # Hourly rate
        ttk.Label(
            main_frame,
            text="Hourly Rate ($):"
        ).pack(anchor="w")

        ttk.Entry(
            main_frame,
            textvariable=self.hourly_rate
        ).pack(fill="x", pady=(5, 15))

        # Search duration
        ttk.Label(
            main_frame,
            text="Search Duration (Minutes):"
        ).pack(anchor="w")

        ttk.Entry(
            main_frame,
            textvariable=self.search_minutes
        ).pack(fill="x", pady=(5, 20))

        # Generate button
        self.generate_button = ttk.Button(
            main_frame,
            text="Generate Report",
            command=self.handle_generate
        )
        self.generate_button.pack(fill="x")

    def handle_generate(self):
        """Validate inputs and send them to the report system."""

        try:
            employees = int(self.employee_count.get())
            rate = float(self.hourly_rate.get())
            minutes = float(self.search_minutes.get())
            day = self.search_day.get()

            if employees <= 0:
                raise ValueError(
                    "Number of employees must be greater than zero."
                )

            if rate <= 0:
                raise ValueError(
                    "Hourly rate must be greater than zero."
                )

            if minutes <= 0:
                raise ValueError(
                    "Search duration must be greater than zero."
                )

            if day not in WEEKDAYS:
                raise ValueError(
                    "Please select a valid day."
                )

        except ValueError as error:
            messagebox.showerror(
                "Invalid Input",
                f"Please check your inputs.\n\n{error}"
            )
            return

        self.generate_button.config(state="disabled")

        try:
            report_path = self.generate_callback(
                employees,
                rate,
                minutes,
                day
            )

            messagebox.showinfo(
                "Report Generated",
                f"Report successfully generated!\n\n{report_path}"
            )

            self.employee_count.set("")
            self.search_minutes.set("")

            # Keep hourly rate and selected day.

        except Exception as error:
            messagebox.showerror(
                "Report Error",
                f"Unable to generate report:\n\n{error}"
            )

        finally:
            self.generate_button.config(state="normal")


def launch_gui(generate_callback):
    """Launch the Search Report application."""

    root = tk.Tk()
    SearchReportGUI(root, generate_callback)
    root.mainloop()
