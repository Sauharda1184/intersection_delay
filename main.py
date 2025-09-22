import tkinter as tk
from tkinter import ttk, messagebox


class DelayStudyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Intersection Delay Study Data Collector")
        self.root.geometry("700x500")  # set a reasonable window size
        self.data = []

        # --- Title ---
        title_label = ttk.Label(root, text="Traffic Delay Study Data Collection",
                                font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # --- Input Frame ---
        input_frame = ttk.LabelFrame(root, text="Enter Vehicle Counts")
        input_frame.pack(fill="x", padx=10, pady=5)

        # Configure grid columns to be centered
        input_frame.grid_columnconfigure(0, weight=1)
        input_frame.grid_columnconfigure(1, weight=1)

        # Time input
        time_label = ttk.Label(input_frame, text="Time (Minute):")
        time_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.minute_var = tk.StringVar()
        self.minute_entry = ttk.Entry(input_frame, textvariable=self.minute_var, width=10)
        self.minute_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Stopped vehicles input
        stopped_label = ttk.Label(input_frame, text="Stopped Vehicles:")
        stopped_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.stopped_var = tk.StringVar()
        self.stopped_entry = ttk.Entry(input_frame, textvariable=self.stopped_var, width=10)
        self.stopped_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Not stopped vehicles input
        notstopped_label = ttk.Label(input_frame, text="Not Stopped Vehicles:")
        notstopped_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.notstopped_var = tk.StringVar()
        self.notstopped_entry = ttk.Entry(input_frame, textvariable=self.notstopped_var, width=10)
        self.notstopped_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Add Entry button
        add_button = ttk.Button(input_frame, text="Add Entry", command=self.add_entry)
        add_button.grid(row=3, column=0, columnspan=2, pady=10)

        # --- Data Table ---
        table_frame = ttk.LabelFrame(root, text="Collected Data")
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side="right", fill="y")

        self.tree = ttk.Treeview(table_frame, columns=("minute", "stopped", "notstopped"), 
                                show="headings", yscrollcommand=scrollbar.set)
        self.tree.heading("minute", text="Minute")
        self.tree.heading("stopped", text="Stopped")
        self.tree.heading("notstopped", text="Not Stopped")
        
        # Configure column widths
        self.tree.column("minute", width=100)
        self.tree.column("stopped", width=100)
        self.tree.column("notstopped", width=100)
        
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar.config(command=self.tree.yview)

        # --- Calculation Buttons ---
        btn_frame = ttk.Frame(root)
        btn_frame.pack(fill="x", padx=10, pady=5)

        # Use grid for button layout
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        calc_button = ttk.Button(btn_frame, text="Calculate Results", command=self.calculate_results)
        calc_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        clear_button = ttk.Button(btn_frame, text="Clear Data", command=self.clear_data)
        clear_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    def add_entry(self):
        """Add a new row of data (minute, stopped, not stopped)."""
        try:
            minute = int(self.minute_var.get())
            stopped = int(self.stopped_var.get())
            notstopped = int(self.notstopped_var.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers.")
            return

        self.data.append((minute, stopped, notstopped))
        self.tree.insert("", "end", values=(minute, stopped, notstopped))

        # clear entry boxes
        self.minute_var.set("")
        self.stopped_var.set("")
        self.notstopped_var.set("")

    def calculate_results(self):
        """Perform delay study calculations and show results."""
        if not self.data:
            messagebox.showwarning("No Data", "Please add some entries first.")
            return

        total_stopped = sum(d[1] for d in self.data)
        total_notstopped = sum(d[2] for d in self.data)
        total_approach = total_stopped + total_notstopped

        # Each "stopped" count represents 15 seconds of delay
        total_delay_seconds = total_stopped * 15
        avg_delay_stopped = total_delay_seconds / total_stopped if total_stopped else 0
        avg_delay_approach = total_delay_seconds / total_approach if total_approach else 0
        percent_stopped = (total_stopped / total_approach * 100) if total_approach else 0

        results = (
            f"Total Delay: {total_delay_seconds} vehicle-seconds\n"
            f"Total Delay (hours): {total_delay_seconds / 3600:.2f} vehicle-hours\n"
            f"Average Delay per Stopped Vehicle: {avg_delay_stopped:.2f} sec\n"
            f"Average Delay per Approach Vehicle: {avg_delay_approach:.2f} sec\n"
            f"Percent of Vehicles Stopped: {percent_stopped:.2f}%"
        )

        messagebox.showinfo("Results", results)

    def clear_data(self):
        """Clear all collected data and reset the table."""
        self.data.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)


if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    app = DelayStudyApp(root)
    root.mainloop()
