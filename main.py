import tkinter as tk
from tkinter import ttk, messagebox
import datetime

class DelayStudyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Intersection Delay Study Data Collector")
        self.data = []

        # --- Input Frame ---
        input_frame = ttk.LabelFrame(root, text="Enter Vehicle Counts")
        input_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(input_frame, text="Time (Minute):").grid(row=0, column=0, padx=5, pady=5)
        self.minute_var = tk.StringVar()
        self.minute_entry = ttk.Entry(input_frame, textvariable=self.minute_var, width=10)
        self.minute_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Stopped Vehicles:").grid(row=1, column=0, padx=5, pady=5)
        self.stopped_var = tk.StringVar()
        self.stopped_entry = ttk.Entry(input_frame, textvariable=self.stopped_var, width=10)
        self.stopped_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Not Stopped Vehicles:").grid(row=2, column=0, padx=5, pady=5)
        self.notstopped_var = tk.StringVar()
        self.notstopped_entry = ttk.Entry(input_frame, textvariable=self.notstopped_var, width=10)
        self.notstopped_entry.grid(row=2, column=1, padx=5, pady=5)

        ttk.Button(input_frame, text="Add Entry", command=self.add_entry).grid(row=3, column=0, columnspan=2, pady=10)

        # --- Data Table ---
        self.tree = ttk.Treeview(root, columns=("minute", "stopped", "notstopped"), show="headings")
        self.tree.heading("minute", text="Minute")
        self.tree.heading("stopped", text="Stopped")
        self.tree.heading("notstopped", text="Not Stopped")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

        # --- Calculation Buttons ---
        btn_frame = ttk.Frame(root)
        btn_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(btn_frame, text="Calculate Results", command=self.calculate_results).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear Data", command=self.clear_data).pack(side="right", padx=5)

    def add_entry(self):
        try:
            minute = int(self.minute_var.get())
            stopped = int(self.stopped_var.get())
            notstopped = int(self.notstopped_var.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers.")
            return

        self.data.append((minute, stopped, notstopped))
        self.tree.insert("", "end", values=(minute, stopped, notstopped))

        self.minute_var.set("")
        self.stopped_var.set("")
        self.notstopped_var.set("")

    def calculate_results(self):
        if not self.data:
            messagebox.showwarning("No Data", "Please add some entries first.")
            return

        total_stopped = sum(d[1] for d in self.data)
        total_notstopped = sum(d[2] for d in self.data)
        total_approach = total_stopped + total_notstopped

        total_delay_seconds = total_stopped * 15  # each count = 15 sec
        avg_delay_stopped = total_delay_seconds / total_stopped if total_stopped else 0
        avg_delay_approach = total_delay_seconds / total_approach if total_approach else 0
        percent_stopped = (total_stopped / total_approach * 100) if total_approach else 0

        results = (
            f"Total Delay: {total_delay_seconds} vehicle-seconds\n"
            f"Total Delay (hours): {total_delay_seconds/3600:.2f} vehicle-hours\n"
            f"Avg Delay per Stopped Vehicle: {avg_delay_stopped:.2f} sec\n"
            f"Avg Delay per Approach Vehicle: {avg_delay_approach:.2f} sec\n"
            f"Percent of Vehicles Stopped: {percent_stopped:.2f}%"
        )

        messagebox.showinfo("Results", results)

    def clear_data(self):
        self.data.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)

if __name__ == "__main__":
    root = tk.Tk()
    app = DelayStudyApp(root)
    root.mainloop()
