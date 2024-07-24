import os
import tkinter as tk
from tkinter import ttk, messagebox
import time
import pandas as pd
import pygetwindow as gw
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
import datetime  # Added for datetime operations
from pynput import mouse, keyboard

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Tracker")
        self.root.geometry("800x600")
        self.root.resizable(False, False)

        self.running = False
        self.data = []
        self.start_time = None
        self.current_app = None
        self.idle_start_time = None
        self.idle_threshold = 30  # seconds
        self.last_activity_time = time.time()

        self.mouse_listener = mouse.Listener(on_move=self.reset_idle_timer, on_click=self.reset_idle_timer)
        self.keyboard_listener = keyboard.Listener(on_press=self.reset_idle_timer)

        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')  # Background color for frames
        self.style.configure('TLabel', font=('Helvetica', 12), background='#f0f0f0')  # Label font and background color
        self.style.configure('TEntry', font=('Helvetica', 12))  # Entry widget font
        self.style.configure('TButton', font=('Helvetica', 12), background='#4cac54', foreground='Black', padding=10)  # Button style
        self.create_login_widgets()

    def create_login_widgets(self):
        self.clear_widgets()
        self.frame_login = ttk.Frame(self.root, padding="10")
        self.frame_login.pack(fill=tk.BOTH, expand=True)

        self.label_username = ttk.Label(self.frame_login, text="Username:")
        self.label_username.pack(pady=10)
        self.entry_username = ttk.Entry(self.frame_login)
        self.entry_username.pack(pady=10)

        self.label_password = ttk.Label(self.frame_login, text="Password:")
        self.label_password.pack(pady=10)
        self.entry_password = ttk.Entry(self.frame_login, show="*")
        self.entry_password.pack(pady=10)

        self.login_button = ttk.Button(self.frame_login, text="Login", command=self.check_login)
        self.login_button.pack(pady=10)

    def check_login(self):
        username = self.entry_username.get()
        password = self.entry_password.get()

        if username == "Chaitanya" and password == "1234":
            self.create_main_widgets()
            self.start_tracking()  # Start tracking immediately after login
        else:
            messagebox.showerror("Login Error", "Invalid username or password")

    def create_main_widgets(self):
        self.clear_widgets()

        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 12), padding=10)
        style.configure('TLabel', font=('Helvetica', 12), padding=10)

        self.frame_controls = ttk.Frame(self.root, padding="10")
        self.frame_controls.pack(fill=tk.BOTH, expand=True)

        self.start_button = ttk.Button(self.frame_controls, text="Start Tracking", command=self.start_tracking, state='disabled')
        self.start_button.pack(pady=5, fill=tk.X)

        self.pause_button = ttk.Button(self.frame_controls, text="Pause Tracking", command=self.pause_tracking)
        self.pause_button.pack(pady=5, fill=tk.X)

        self.resume_button = ttk.Button(self.frame_controls, text="Resume Tracking", command=self.resume_tracking, state='disabled')
        self.resume_button.pack(pady=5, fill=tk.X)

        self.logout_button = ttk.Button(self.frame_controls, text="Logout", command=self.logout)
        self.logout_button.pack(pady=5, fill=tk.X)

        self.frame_status = ttk.Frame(self.root, padding="10")
        self.frame_status.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(self.frame_status, text="Status: Tracking", foreground="green")
        self.status_label.pack(pady=5, fill=tk.X)

        self.active_task_label = ttk.Label(self.frame_status, text="Active Task: None")
        self.active_task_label.pack(pady=5, fill=tk.X)

        self.history_label = ttk.Label(self.frame_status, text="History:")
        self.history_label.pack(pady=5, fill=tk.X)

        self.history_tree = ttk.Treeview(self.frame_status, columns=("Application", "Duration"), show="headings")
        self.history_tree.heading("Application", text="Application")
        self.history_tree.heading("Duration", text="Duration")
        self.history_tree.pack(pady=5, fill=tk.BOTH, expand=True)

    def clear_widgets(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def start_tracking(self):
        self.running = True
        self.start_button.config(state='disabled')
        self.pause_button.config(state='normal')
        self.logout_button.config(state='normal')
        self.status_label.config(text="Status: Tracking", foreground="green")
        
        self.mouse_listener.start()
        self.keyboard_listener.start()
        
        self.track_time()

    def pause_tracking(self):
        self.running = False
        self.pause_button.config(state='disabled')
        self.resume_button.config(state='normal')
        self.status_label.config(text="Status: Paused", foreground="orange")

    def resume_tracking(self):
        self.running = True
        self.resume_button.config(state='disabled')
        self.pause_button.config(state='normal')
        self.status_label.config(text="Status: Tracking", foreground="green")
        self.track_time()

    def track_time(self):
        if self.running:
            active_window_name = self.get_active_window()
            current_time = time.time()

            if current_time - self.last_activity_time >= self.idle_threshold:
                self.handle_active_app("Idle", current_time)
            else:
                self.handle_active_app(active_window_name, current_time)

            self.root.after(1000, self.track_time)

    def handle_active_app(self, active_window_name, current_time):
        if active_window_name != "Idle":
            self.reset_idle_timer()

        if self.current_app and self.current_app != active_window_name:
            duration = current_time - self.start_time
            self.data.append((self.current_app, duration))
            self.update_history_tree()

        if self.current_app != active_window_name:
            self.current_app = active_window_name
            self.start_time = current_time
            self.active_task_label.config(text=f"Active Task: {self.current_app}", foreground="black")
            
        if self.current_app and any(app in self.current_app for app in ["YouTube", "Facebook", "Twitter", "Instagram"]):
            self.active_task_label.config(foreground="red")  # Non-productive apps in red

    def get_active_window(self):
        active_window = gw.getActiveWindow()
        if active_window:
            return active_window.title
        else:
            return "Unknown"

    def update_history_tree(self):
        self.history_tree.delete(*self.history_tree.get_children())
        for entry in self.data:
            duration_formatted = self.format_duration(entry[1])
            self.history_tree.insert("", tk.END, values=(entry[0], duration_formatted))

    def format_duration(self, duration_seconds):
        # Convert duration from seconds to timedelta and then to HH:MM:SS format
        duration_timedelta = datetime.timedelta(seconds=duration_seconds)
        hours, remainder = divmod(duration_timedelta.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    def reset_idle_timer(self, *args):
        self.last_activity_time = time.time()
        if self.current_app == "Idle":
            self.current_app = None
            self.start_time = self.last_activity_time

    def logout(self):
        self.running = False
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
        self.generate_report()
        self.open_report_and_exit()

    def generate_report(self):
        productive_apps = ["Python", "PyCharm", "VSCode", "Microsoft Word", "Excel"]
        non_productive_apps = ["YouTube", "Facebook", "Twitter", "Instagram"]

        # Aggregate and sum durations for repeated tasks
        aggregated_data = {}
        for app, duration in self.data:
            if app in aggregated_data:
                aggregated_data[app] += duration
            else:
                aggregated_data[app] = duration

        # Create DataFrame from aggregated data
        aggregated_data = [(app, duration) for app, duration in aggregated_data.items()]

        # Convert duration from seconds to timedelta format (HH:MM:SS)
        df = pd.DataFrame(aggregated_data, columns=["Application", "Duration"])
        df["Duration"] = df["Duration"].apply(lambda x: self.format_duration(x))

        # Add a column with numeric durations (in seconds) for charting
        df["Duration (seconds)"] = [duration.total_seconds() for duration in pd.to_timedelta(df["Duration"])]

        # Categorize applications
        df['Category'] = df['Application'].apply(lambda x: 'Productive' if any(app in x for app in productive_apps) else ('Non-Productive' if any(app in x for app in non_productive_apps) else 'Other'))

        # Generate report
        report_path = "time_report.xlsx"
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Time Report')

            # Load workbook and worksheet
            wb = writer.book
            ws = wb['Time Report']

            # Create Pie Chart
            chart = PieChart()

            # Correct ranges for labels and data
            labels = Reference(ws, min_col=1, min_row=2, max_row=len(df) + 1)  # Applications as labels
            data = Reference(ws, min_col=3, min_row=1, max_row=len(df) + 1)    # Numeric durations as data

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
            chart.title = 'Productive vs Non-Productive Time'

            # Adjusting the height and width of the chart
            chart.height = 15
            chart.width = 30

            # Add chart to worksheet
            ws.add_chart(chart, "G2")

        self.status_label.config(text=f"Report saved to {report_path}", foreground="blue")
        print(f"Report saved to {report_path}")

    def open_report_and_exit(self):
        report_path = "time_report.xlsx"
        if os.path.exists(report_path):
            os.startfile(report_path)
            self.root.destroy()  # Close the Tkinter GUI window after opening the report
        else:
            messagebox.showerror("File Not Found", "Report file not found.")

root = tk.Tk()
app = TimeTrackerApp(root)
root.mainloop()
