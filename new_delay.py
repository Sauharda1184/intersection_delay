import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd
import csv

class DelayStudyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Intersection Delay Study Data Collector")
        self.root.geometry("1200x800")
        self.data = []
        # Stores observed approach volume per (period, direction)
        self.approach_volume = {}
        # Aggregated unique counts per (period, direction)
        self.unique_map = {}
        # Temporary per-minute unique results from Unique Assist
        self.current_unique_stopped = None
        self.current_unique_notstopped = None
        # Inline unique per-minute counter storage: {(period, direction, minute): {"stopped": int, "notstopped": int}}
        self.unique_minute_map = {}

        # --- Title ---
        title_label = ttk.Label(root, text="Traffic Delay Study Data Collection",
                              font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # --- Study Info Frame ---
        info_frame = ttk.LabelFrame(root, text="Study Information")
        info_frame.pack(fill="x", padx=10, pady=5)

        # Date and Time
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        date_label = ttk.Label(info_frame, text="Date:")
        date_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        date_entry = ttk.Entry(info_frame, textvariable=self.date_var, width=12)
        date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Intersection Name
        self.intersection_var = tk.StringVar()
        intersection_label = ttk.Label(info_frame, text="Intersection:")
        intersection_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        intersection_entry = ttk.Entry(info_frame, textvariable=self.intersection_var, width=25)
        intersection_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        # Weather Conditions
        self.weather_var = tk.StringVar()
        weather_label = ttk.Label(info_frame, text="Weather:")
        weather_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        weather_entry = ttk.Entry(info_frame, textvariable=self.weather_var, width=12)
        weather_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Start Time
        self.start_hour_var = tk.StringVar(value="08")
        self.start_minute_var = tk.StringVar(value="00")
        self.start_second_var = tk.StringVar(value="00")
        start_time_label = ttk.Label(info_frame, text="Start Time:")
        start_time_label.grid(row=1, column=2, padx=5, pady=5, sticky="e")
        start_hour_entry = ttk.Entry(info_frame, textvariable=self.start_hour_var, width=3)
        start_hour_entry.grid(row=1, column=3, padx=2, pady=5, sticky="w")
        ttk.Label(info_frame, text=":").grid(row=1, column=4, padx=0, pady=5)
        start_minute_entry = ttk.Entry(info_frame, textvariable=self.start_minute_var, width=3)
        start_minute_entry.grid(row=1, column=5, padx=2, pady=5, sticky="w")
        ttk.Label(info_frame, text=":").grid(row=1, column=6, padx=0, pady=5)
        start_second_entry = ttk.Entry(info_frame, textvariable=self.start_second_var, width=3)
        start_second_entry.grid(row=1, column=7, padx=2, pady=5, sticky="w")

        # --- Input Frame ---
        input_frame = ttk.LabelFrame(root, text="Vehicle Count Data Entry")
        input_frame.pack(fill="x", padx=10, pady=5)

        # Study Period and Time inputs
        time_dir_frame = ttk.Frame(input_frame)
        time_dir_frame.pack(fill="x", padx=5, pady=5)

        # 15-minute period selection
        period_label = ttk.Label(time_dir_frame, text="15-Min Period:")
        period_label.pack(side="left", padx=5)
        self.period_var = tk.StringVar(value="1")
        self.period_combo = ttk.Combobox(time_dir_frame, textvariable=self.period_var,
                                        values=["1 (0-15 min)", "2 (15-30 min)", "3 (30-45 min)", "4 (45-60 min)"], 
                                        width=15, state="readonly")
        self.period_combo.pack(side="left", padx=5)

        # Current minute within period
        minute_label = ttk.Label(time_dir_frame, text="Minute in Period:")
        minute_label.pack(side="left", padx=5)
        self.minute_var = tk.StringVar(value="1")
        self.minute_entry = ttk.Entry(time_dir_frame, textvariable=self.minute_var, width=8, state="readonly")
        self.minute_entry.pack(side="left", padx=5)

        # Direction selection
        direction_label = ttk.Label(time_dir_frame, text="Direction:")
        direction_label.pack(side="left", padx=5)
        self.direction_var = tk.StringVar(value="North")
        self.direction_combo = ttk.Combobox(time_dir_frame, textvariable=self.direction_var,
                                          values=["North", "South", "East", "West"], width=10, state="readonly")
        self.direction_combo.pack(side="left", padx=5)

        # Auto-progress button
        self.progress_button = ttk.Button(time_dir_frame, text="Next Minute", command=self.next_minute)
        self.progress_button.pack(side="left", padx=5)

        # Keep the tally matched with selected period/direction
        self.period_combo.bind("<<ComboboxSelected>>", self.on_period_or_direction_change)
        self.direction_combo.bind("<<ComboboxSelected>>", self.on_period_or_direction_change)

        # Stopped vehicle counts frame
        stopped_frame = ttk.LabelFrame(input_frame, text="Vehicles Stopped at Each Interval")
        stopped_frame.pack(fill="x", padx=5, pady=5)

        # Initialize variables for interval counts
        self.interval_vars = []
        interval_labels = ["0-15 sec", "15-30 sec", "30-45 sec", "45-60 sec"]
        
        for i, label in enumerate(interval_labels):
            interval_frame = ttk.Frame(stopped_frame)
            interval_frame.pack(side="left", expand=True, padx=5, pady=5)
            
            ttk.Label(interval_frame, text=label).pack()
            count_var = tk.StringVar(value="0")
            self.interval_vars.append(count_var)
            
            count_entry = ttk.Entry(interval_frame, textvariable=count_var, width=5, justify="center")
            count_entry.pack(pady=2)
            
            btn_frame = ttk.Frame(interval_frame)
            btn_frame.pack()
            
            ttk.Button(btn_frame, text="+", width=3,
                      command=lambda v=count_var: self.increment_count(v)).pack(side="left", padx=1)
            ttk.Button(btn_frame, text="-", width=3,
                      command=lambda v=count_var: self.decrement_count(v)).pack(side="left", padx=1)

        # Not stopped vehicles frame
        notstopped_frame = ttk.LabelFrame(input_frame, text="Vehicles Not Stopped at Each Interval")
        notstopped_frame.pack(fill="x", padx=5, pady=5)

        # Initialize variables for not-stopped interval counts
        self.notstopped_vars = []
        
        for i, label in enumerate(interval_labels):
            interval_frame = ttk.Frame(notstopped_frame)
            interval_frame.pack(side="left", expand=True, padx=5, pady=5)
            
            ttk.Label(interval_frame, text=label).pack()
            count_var = tk.StringVar(value="0")
            self.notstopped_vars.append(count_var)
            
            count_entry = ttk.Entry(interval_frame, textvariable=count_var, width=5, justify="center")
            count_entry.pack(pady=2)
            
            btn_frame = ttk.Frame(interval_frame)
            btn_frame.pack()
            
            ttk.Button(btn_frame, text="+", width=3,
                      command=lambda v=count_var: self.increment_count(v)).pack(side="left", padx=1)
            ttk.Button(btn_frame, text="-", width=3,
                      command=lambda v=count_var: self.decrement_count(v)).pack(side="left", padx=1)

        # Inline Unique Vehicle Counters (per-minute)
        unique_inline = ttk.LabelFrame(input_frame, text="Unique Vehicles This Minute (Approach Volume)")
        unique_inline.pack(fill="x", padx=5, pady=5)

        ttk.Label(unique_inline, text="Unique Stopped").pack(side="left", padx=5)
        self.unique_minute_stopped_var = tk.StringVar(value="0")
        ttk.Entry(unique_inline, textvariable=self.unique_minute_stopped_var, width=7, justify="center").pack(side="left", padx=2)
        ttk.Button(unique_inline, text="+", width=3, command=lambda: self.increment_unique_minute(True)).pack(side="left", padx=2)

        ttk.Label(unique_inline, text="Unique Not Stopped").pack(side="left", padx=10)
        self.unique_minute_notstopped_var = tk.StringVar(value="0")
        ttk.Entry(unique_inline, textvariable=self.unique_minute_notstopped_var, width=7, justify="center").pack(side="left", padx=2)
        ttk.Button(unique_inline, text="+", width=3, command=lambda: self.increment_unique_minute(False)).pack(side="left", padx=2)

        ttk.Button(unique_inline, text="Reset Unique", command=self.reset_unique_minute).pack(side="left", padx=10)

        # Approach volume tally frame (15-minute)
        volume_frame = ttk.LabelFrame(input_frame, text="Approach Volume Tally (15-minute)")
        volume_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(volume_frame, text="Observed vehicles in current 15-min period and direction").pack(side="left", padx=5)
        self.approach_volume_var = tk.StringVar(value="0")
        self.approach_volume_entry = ttk.Entry(volume_frame, textvariable=self.approach_volume_var, width=7, justify="center")
        self.approach_volume_entry.pack(side="left", padx=5)
        ttk.Button(volume_frame, text="+", width=3, command=lambda: self.increment_count(self.approach_volume_var)).pack(side="left", padx=2)
        ttk.Button(volume_frame, text="-", width=3, command=lambda: self.decrement_count(self.approach_volume_var)).pack(side="left", padx=2)
        ttk.Button(volume_frame, text="Reset", command=lambda: self.approach_volume_var.set("0")).pack(side="left", padx=6)
        ttk.Button(volume_frame, text="Save 15-min Volume", command=self.save_approach_volume).pack(side="left", padx=6)

        # Add Entry and Reset Counts buttons
        btn_frame = ttk.Frame(input_frame)
        btn_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(btn_frame, text="Add Entry", command=self.add_entry).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Reset Counts", command=self.reset_counts).pack(side="left", padx=5)

        # --- Data Table ---
        table_frame = ttk.LabelFrame(root, text="Collected Data")
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side="right", fill="y")

        # Configure treeview
        self.tree = ttk.Treeview(table_frame, 
            columns=(
                "period", "minute", "direction", 
                "stopped_0", "stopped_15", "stopped_30", "stopped_45", "total_stopped",
                "notstopped_0", "notstopped_15", "notstopped_30", "notstopped_45", "total_notstopped",
                "total"
            ),
            show="headings",
            yscrollcommand=scrollbar.set)
        
        # Configure headers and columns
        headers = {
            "period": "Interval",
            "minute": "Minute",
            "direction": "Direction",
            "stopped_0": "Stopped 0-15s",
            "stopped_15": "Stopped 15-30s",
            "stopped_30": "Stopped 30-45s",
            "stopped_45": "Stopped 45-60s",
            "total_stopped": "Total Stopped",
            "notstopped_0": "Not Stopped 0-15s",
            "notstopped_15": "Not Stopped 15-30s",
            "notstopped_30": "Not Stopped 30-45s",
            "notstopped_45": "Not Stopped 45-60s",
            "total_notstopped": "Total Not Stopped",
            "total": "Total Volume"
        }
        
        for col, heading in headers.items():
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=80)
        
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar.config(command=self.tree.yview)

        # --- Action Buttons ---
        action_frame = ttk.Frame(root)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(action_frame, text="Calculate Results", 
                  command=self.calculate_results).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Generate Form 2", 
                  command=self.generate_form2).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Save Data", 
                  command=self.save_data).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Clear All Data", 
                  command=self.clear_data).pack(side="left", padx=5)

    def calculate_actual_time(self, period, minute):
        """Calculate actual time based on start time, period, and minute."""
        try:
            start_hour = int(self.start_hour_var.get())
            start_minute = int(self.start_minute_var.get())
            start_second = int(self.start_second_var.get())
            
            # Calculate total minutes from start
            total_minutes = (period - 1) * 15 + (minute - 1)
            
            # Add to start time
            current_minute = start_minute + total_minutes
            current_hour = start_hour + (current_minute // 60)
            current_minute = current_minute % 60
            
            return f"{current_hour:02d}:{current_minute:02d}:{start_second:02d}"
        except ValueError:
            return "00:00:00"

    def increment_count(self, var):
        try:
            current = int(var.get())
            var.set(str(current + 1))
        except ValueError:
            var.set("1")

    def decrement_count(self, var):
        try:
            current = int(var.get())
            if current > 0:
                var.set(str(current - 1))
        except ValueError:
            var.set("0")

    def reset_counts(self):
        for var in self.interval_vars:
            var.set("0")
        for var in self.notstopped_vars:
            var.set("0")

    def next_minute(self):
        """Automatically progress to the next minute within the current period."""
        current_minute = int(self.minute_var.get())
        if current_minute < 15:  # 1-15 minutes in each 15-minute period
            self.minute_var.set(str(current_minute + 1))
            # Load unique counter for new minute
            self.load_unique_counter_for_current_context()
        else:
            # Move to next period
            current_period = int(self.period_var.get().split()[0])
            if current_period < 4:
                next_period = current_period + 1
                self.period_var.set(f"{next_period} ({15*(next_period-1)}-{15*next_period} min)")
                self.minute_var.set("1")
                # Reset tally display for new period; load if previously saved
                self.on_period_or_direction_change()
            else:
                messagebox.showinfo("Study Complete", "All 4 periods (60 minutes) have been completed!")

    def add_entry(self):
        try:
            period = int(self.period_var.get().split()[0])
            minute = int(self.minute_var.get())
            direction = self.direction_var.get()
            
            # Get all interval counts for stopped vehicles
            stopped_counts = [int(var.get()) for var in self.interval_vars]
            total_stopped = sum(stopped_counts)
            
            # Get all interval counts for non-stopped vehicles
            notstopped_counts = [int(var.get()) for var in self.notstopped_vars]
            total_notstopped = sum(notstopped_counts)
            
            # Get unique counts from inline counters (current values in the UI)
            try:
                unique_stopped_count = int(self.unique_minute_stopped_var.get() or 0)
                unique_notstopped_count = int(self.unique_minute_notstopped_var.get() or 0)
            except ValueError:
                unique_stopped_count = 0
                unique_notstopped_count = 0
            
            # Store these unique counts for this minute
            self.set_unique_minute(period, direction, minute, {
                "stopped": unique_stopped_count, 
                "notstopped": unique_notstopped_count
            })
            
            unique_total_minute = unique_stopped_count + unique_notstopped_count

            # Create data tuple with all values including period, using unique totals for 'Total Stopped' and 'Total Volume'
            entry_data = (
                period, minute, direction, 
                *stopped_counts, unique_stopped_count,
                *notstopped_counts, unique_notstopped_count,
                unique_total_minute
            )
            self.data.append(entry_data)
            # Recompute aggregates per (period, direction)
            self.recompute_unique_aggregates()
            # Clear temp unique once used
            self.current_unique_stopped = None
            self.current_unique_notstopped = None

            # Add to treeview
            self.tree.insert("", "end", values=entry_data)

            # Auto-progress to next minute and reset counts
            self.next_minute()
            self.reset_counts()
            
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers.")

    def open_unique_assist(self):
        """Dialog to help deduplicate counts across 15s snapshots for this minute."""
        # Gather current counts
        try:
            s0, s15, s30, s45 = [int(v.get()) for v in self.interval_vars]
            n0, n15, n30, n45 = [int(v.get()) for v in self.notstopped_vars]
        except ValueError:
            messagebox.showerror("Invalid Input", "Please ensure counts are numbers before using Unique Assist.")
            return

        win = tk.Toplevel(self.root)
        win.title("Unique Assist (Same vs Different Vehicles)")
        win.geometry("520x340")

        container = ttk.Frame(win)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(container, text="Enter how many at each later snapshot were the SAME as previous snapshot").grid(row=0, column=0, columnspan=4, pady=5)

        # Headers
        ttk.Label(container, text="Interval").grid(row=1, column=0, padx=4)
        ttk.Label(container, text="Stopped: same as previous").grid(row=1, column=1, padx=4)
        ttk.Label(container, text="Not Stopped: same as previous").grid(row=1, column=2, padx=4)

        # For +15, +30, +45 we need same-as-previous entries
        same_stopped_vars = {}
        same_notstopped_vars = {}
        intervals = [("+15 sec", s15, n15), ("+30 sec", s30, n30), ("+45 sec", s45, n45)]
        for idx, (label, sval, nval) in enumerate(intervals, start=2):
            ttk.Label(container, text=label).grid(row=idx, column=0, sticky="w", padx=4, pady=2)
            sv = tk.StringVar(value="0")
            nv = tk.StringVar(value="0")
            same_stopped_vars[label] = sv
            same_notstopped_vars[label] = nv
            ttk.Entry(container, textvariable=sv, width=10, justify="center").grid(row=idx, column=1, padx=4)
            ttk.Entry(container, textvariable=nv, width=10, justify="center").grid(row=idx, column=2, padx=4)

        # Results labels
        result_lbl = ttk.Label(container, text="")
        result_lbl.grid(row=6, column=0, columnspan=4, pady=10)

        def compute_unique():
            def safe_int(v):
                try:
                    return max(0, int(v))
                except Exception:
                    return 0
            same15_s = safe_int(same_stopped_vars["+15 sec"].get())
            same30_s = safe_int(same_stopped_vars["+30 sec"].get())
            same45_s = safe_int(same_stopped_vars["+45 sec"].get())
            same15_n = safe_int(same_notstopped_vars["+15 sec"].get())
            same30_n = safe_int(same_notstopped_vars["+30 sec"].get())
            same45_n = safe_int(same_notstopped_vars["+45 sec"].get())

            u_stopped = max(0, s0) + max(0, s15 - same15_s) + max(0, s30 - same30_s) + max(0, s45 - same45_s)
            u_notstopped = max(0, n0) + max(0, n15 - same15_n) + max(0, n30 - same30_n) + max(0, n45 - same45_n)

            self.current_unique_stopped = u_stopped
            self.current_unique_notstopped = u_notstopped
            result_lbl.config(text=f"Unique Stopped: {u_stopped}    Unique Not Stopped: {u_notstopped}    Total: {u_stopped + u_notstopped}")

        def use_and_close():
            if self.current_unique_stopped is None:
                compute_unique()
            win.destroy()

        btns = ttk.Frame(container)
        btns.grid(row=7, column=0, columnspan=4, pady=8)
        ttk.Button(btns, text="Compute", command=compute_unique).pack(side="left", padx=5)
        ttk.Button(btns, text="Use These", command=use_and_close).pack(side="left", padx=5)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=5)

    def save_approach_volume(self):
        """Save observed approach volume for current 15-minute period and direction."""
        try:
            period = int(self.period_var.get().split()[0])
            direction = self.direction_var.get()
            count = int(self.approach_volume_var.get())
            self.approach_volume[(period, direction)] = count
            messagebox.showinfo("Saved", f"Saved {count} vehicles for Period {period} - {direction}")
        except ValueError:
            messagebox.showerror("Invalid Input", "Approach volume must be a whole number.")

    def on_period_or_direction_change(self, event=None):
        """Load saved tally for selected period/direction, else clear to 0."""
        try:
            period = int(self.period_var.get().split()[0])
        except Exception:
            period = 1
        direction = self.direction_var.get()
        val = self.approach_volume.get((period, direction))
        self.approach_volume_var.set(str(val) if val is not None else "0")
        # Load the unique minute counter for current minute context
        self.load_unique_counter_for_current_context()

    def get_unique_minute(self, period, direction, minute):
        return self.unique_minute_map.get((period, direction, minute))

    def set_unique_minute(self, period, direction, minute, value):
        # value is a dict {"stopped": int, "notstopped": int}
        self.unique_minute_map[(period, direction, minute)] = {
            "stopped": max(0, int(value.get("stopped", 0))),
            "notstopped": max(0, int(value.get("notstopped", 0)))
        }

    def load_unique_counter_for_current_context(self):
        try:
            period = int(self.period_var.get().split()[0])
            minute = int(self.minute_var.get())
        except Exception:
            period = 1
            minute = 1
        direction = self.direction_var.get()
        v = self.get_unique_minute(period, direction, minute)
        if v is None:
            self.unique_minute_stopped_var.set("0")
            self.unique_minute_notstopped_var.set("0")
        else:
            self.unique_minute_stopped_var.set(str(v.get("stopped", 0)))
            self.unique_minute_notstopped_var.set(str(v.get("notstopped", 0)))

    def increment_unique_minute(self, is_stopped):
        if is_stopped:
            try:
                current = int(self.unique_minute_stopped_var.get())
            except ValueError:
                current = 0
            new_val = current + 1
            self.unique_minute_stopped_var.set(str(new_val))
        else:
            try:
                current = int(self.unique_minute_notstopped_var.get())
            except ValueError:
                current = 0
            new_val = current + 1
            self.unique_minute_notstopped_var.set(str(new_val))
        # Persist and aggregate
        period = int(self.period_var.get().split()[0])
        minute = int(self.minute_var.get())
        direction = self.direction_var.get()
        self.set_unique_minute(period, direction, minute, {
            "stopped": int(self.unique_minute_stopped_var.get() or 0),
            "notstopped": int(self.unique_minute_notstopped_var.get() or 0)
        })
        self.recompute_unique_aggregates()

    def reset_unique_minute(self):
        self.unique_minute_stopped_var.set("0")
        self.unique_minute_notstopped_var.set("0")
        try:
            period = int(self.period_var.get().split()[0])
            minute = int(self.minute_var.get())
        except Exception:
            return
        direction = self.direction_var.get()
        self.set_unique_minute(period, direction, minute, {"stopped": 0, "notstopped": 0})
        self.recompute_unique_aggregates()

    def recompute_unique_aggregates(self):
        # Sum per (period, direction) from per-minute unique totals
        agg = {}
        for (p, d, m), val in self.unique_minute_map.items():
            if (p, d) not in agg:
                agg[(p, d)] = {"stopped": 0, "notstopped": 0}
            agg[(p, d)]["stopped"] += int(val.get("stopped", 0))
            agg[(p, d)]["notstopped"] += int(val.get("notstopped", 0))
        self.unique_map = agg
        # Keep approach_volume in sync for compatibility and display
        self.approach_volume = {k: v["stopped"] + v["notstopped"] for k, v in self.unique_map.items()}
        # Update displayed approach tally for current selection
        try:
            period = int(self.period_var.get().split()[0])
        except Exception:
            period = 1
        direction = self.direction_var.get()
        self.approach_volume_var.set(str(self.approach_volume.get((period, direction), 0)))

    def calculate_results(self):
        if not self.data:
            messagebox.showwarning("No Data", "Please add some entries first.")
            return

        # Calculate results by period and direction
        results_by_period_direction = {}
        results_by_period_overall = {}
        overall_results_by_direction = {}
        
        for period in range(1, 5):
            results_by_period_direction[period] = {}
            # Aggregate period totals across directions
            period_rows = [d for d in self.data if d[0] == period]
            if period_rows:
                # Calculate total stopped vehicles (sum of all 15-second interval counts) for this period
                p_total_stopped_all_intervals = 0
                for direction in ["North", "South", "East", "West"]:
                    direction_data = [d for d in self.data if d[0] == period and d[2] == direction]
                    for row in direction_data:
                        # Sum all stopped counts: stopped_0 + stopped_15 + stopped_30 + stopped_45
                        p_total_stopped_all_intervals += sum(row[3:7])  # indices 3-6 are stopped counts
                
                # Use UNIQUE stopped totals for approach volume calculation
                p_total_stopped_unique = 0
                for direction in ["North", "South", "East", "West"]:
                    u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                    p_total_stopped_unique += u["stopped"]
                
                # Approach volume for period = sum of unique per direction for this period
                p_total_vehicles = 0
                for direction in ["North", "South", "East", "West"]:
                    u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                    p_total_vehicles += (u["stopped"] + u["notstopped"])
                
                if p_total_vehicles > 0:
                    # Total Delay = Total Number Stopped at Time Intervals * 15 seconds
                    p_total_delay = p_total_stopped_all_intervals * 15
                    # Average Delay per Stopped Vehicle = Total Delay / Number Stopped (total stopped, not unique)
                    p_avg_delay_stopped = p_total_delay / p_total_stopped_all_intervals if p_total_stopped_all_intervals > 0 else 0
                    # Average Delay per Approach Vehicle = Total Delay / Total Approach Volume (unique vehicles)
                    p_avg_delay_approach = p_total_delay / p_total_vehicles
                    # Percent of Vehicles Stopped = Number Stopped (total) / Approach Volume (unique)
                    p_percent_stopped = (p_total_stopped_all_intervals / p_total_vehicles) * 100
                    results_by_period_overall[period] = {
                        "Total Vehicles": p_total_vehicles,
                        "Total Stopped": p_total_stopped_all_intervals,
                        "Total Delay (sec)": p_total_delay,
                        "Avg Delay per Stopped (sec)": p_avg_delay_stopped,
                        "Avg Delay per Approach (sec)": p_avg_delay_approach,
                        "Percent Stopped": p_percent_stopped
                    }
            for direction in ["North", "South", "East", "West"]:
                # Filter data for this period and direction (period is now index 0, direction is index 2)
                period_direction_data = [d for d in self.data if d[0] == period and d[2] == direction]
                if period_direction_data:
                    # Calculate total stopped vehicles (sum of all 15-second interval counts) for this period+direction
                    total_stopped_all_intervals = 0
                    for row in period_direction_data:
                        # Sum all stopped counts: stopped_0 + stopped_15 + stopped_30 + stopped_45
                        total_stopped_all_intervals += sum(row[3:7])  # indices 3-6 are stopped counts
                    
                    # Use UNIQUE stopped and UNIQUE approach volume for this period+direction
                    u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                    total_vehicles = u["stopped"] + u["notstopped"]
                    
                    if total_vehicles > 0:
                        # Total Delay = Total Number Stopped at Time Intervals * 15 seconds
                        total_delay = total_stopped_all_intervals * 15
                        # Average Delay per Stopped Vehicle = Total Delay / Number Stopped (total stopped, not unique)
                        avg_delay_stopped = total_delay / total_stopped_all_intervals if total_stopped_all_intervals > 0 else 0
                        # Average Delay per Approach Vehicle = Total Delay / Total Approach Volume (unique vehicles)
                        avg_delay_approach = total_delay / total_vehicles
                        # Percent of Vehicles Stopped = Number Stopped (total) / Approach Volume (unique)
                        percent_stopped = (total_stopped_all_intervals / total_vehicles) * 100
                        
                        results_by_period_direction[period][direction] = {
                            "Total Vehicles": total_vehicles,
                            "Total Stopped": total_stopped_all_intervals,
                            "Total Delay (sec)": total_delay,
                            "Avg Delay per Stopped (sec)": avg_delay_stopped,
                            "Avg Delay per Approach (sec)": avg_delay_approach,
                            "Percent Stopped": percent_stopped
                        }
        
        # Calculate overall results by direction
        for direction in ["North", "South", "East", "West"]:
            direction_data = [d for d in self.data if d[2] == direction]
            if direction_data:
                # Calculate total stopped vehicles (sum of all 15-second interval counts) for this direction across all periods
                total_stopped_all_intervals = 0
                for row in direction_data:
                    # Sum all stopped counts: stopped_0 + stopped_15 + stopped_30 + stopped_45
                    total_stopped_all_intervals += sum(row[3:7])  # indices 3-6 are stopped counts
                
                # Use UNIQUE stopped and UNIQUE approach volume overall for direction across all periods
                total_vehicles = 0
                for period in range(1, 5):
                    u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                    total_vehicles += (u["stopped"] + u["notstopped"])
                
                if total_vehicles > 0:
                    # Total Delay = Total Number Stopped at Time Intervals * 15 seconds
                    total_delay = total_stopped_all_intervals * 15
                    # Average Delay per Stopped Vehicle = Total Delay / Number Stopped (total stopped, not unique)
                    avg_delay_stopped = total_delay / total_stopped_all_intervals if total_stopped_all_intervals > 0 else 0
                    # Average Delay per Approach Vehicle = Total Delay / Total Approach Volume (unique vehicles)
                    avg_delay_approach = total_delay / total_vehicles
                    # Percent of Vehicles Stopped = Number Stopped (total) / Approach Volume (unique)
                    percent_stopped = (total_stopped_all_intervals / total_vehicles) * 100
                    
                    overall_results_by_direction[direction] = {
                        "Total Vehicles": total_vehicles,
                        "Total Stopped": total_stopped_all_intervals,
                        "Total Delay (sec)": total_delay,
                        "Avg Delay per Stopped (sec)": avg_delay_stopped,
                        "Avg Delay per Approach (sec)": avg_delay_approach,
                        "Percent Stopped": percent_stopped
                    }

        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile="Intersection_Delay_Results.xlsx"
            )
            if filename:
                # Create DataFrame for overall results
                overall_results_data = []
                for direction, data in overall_results_by_direction.items():
                    overall_results_data.append({
                        "Direction": direction,
                        **data
                    })
                
                overall_df = pd.DataFrame(overall_results_data)
                
                # Create DataFrame for period-by-period results (by direction)
                period_results_data = []
                for period in range(1, 5):
                    for direction, data in results_by_period_direction[period].items():
                        period_results_data.append({
                            "Period": period,
                            "Direction": direction,
                            **data
                        })
                
                period_df = pd.DataFrame(period_results_data)

                # Create DataFrame for period summaries (all directions combined)
                period_overall_results_data = []
                for period in sorted(results_by_period_overall.keys()):
                    period_overall_results_data.append({
                        "Period": period,
                        **results_by_period_overall[period]
                    })
                period_overall_df = pd.DataFrame(period_overall_results_data)
                
                # Write to Excel with proper formatting
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # Write study information
                    info_df = pd.DataFrame([
                        ["Intersection Delay Study Results"],
                        ["Date:", self.date_var.get()],
                        ["Intersection:", self.intersection_var.get()],
                        ["Weather:", self.weather_var.get()],
                        [""]
                    ])
                    info_df.to_excel(writer, sheet_name='Overall Results', index=False, header=False)
                    
                    # Write overall results data
                    overall_df.to_excel(writer, sheet_name='Overall Results', index=False, startrow=len(info_df))
                    
                    # Write period-by-period results to separate sheet
                    period_df.to_excel(writer, sheet_name='Period Results', index=False)

                    # Write period summary sheet
                    period_overall_df.to_excel(writer, sheet_name='Period Summary', index=False)
                    
                    # Auto-adjust columns width
                    for sheet_name in ['Overall Results', 'Period Results', 'Period Summary']:
                        worksheet = writer.sheets[sheet_name]
                        if sheet_name == 'Overall Results':
                            df_to_adjust = overall_df
                        elif sheet_name == 'Period Results':
                            df_to_adjust = period_df
                        else:
                            df_to_adjust = period_overall_df
                        for idx, col in enumerate(df_to_adjust.columns):
                            worksheet.column_dimensions[chr(65 + idx)].width = 20

                messagebox.showinfo("Success", f"Results saved to {filename}")
                self.show_results_popup(overall_results_by_direction, results_by_period_direction, results_by_period_overall)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving results: {str(e)}")

    def show_results_popup(self, overall_results_by_direction, results_by_period_direction, results_by_period_overall):
        results = f"Intersection: {self.intersection_var.get()}\n"
        results += f"Date: {self.date_var.get()}\n"
        results += f"Weather: {self.weather_var.get()}\n\n"
        
        # Overall Results
        results += "=== OVERALL STUDY RESULTS ===\n"
        for direction, data in overall_results_by_direction.items():
            results += f"\n{direction} Approach:\n"
            results += f"Total Vehicles: {data['Total Vehicles']}\n"
            results += f"Total Stopped: {data['Total Stopped']}\n"
            results += f"Total Delay: {data['Total Delay (sec)']:.1f} seconds\n"
            results += f"Avg Delay per Stopped Vehicle: {data['Avg Delay per Stopped (sec)']:.1f} seconds\n"
            results += f"Avg Delay per Approach Vehicle: {data['Avg Delay per Approach (sec)']:.1f} seconds\n"
            results += f"Percent Stopped: {data['Percent Stopped']:.1f}%\n"
        
        # Period-by-Period Results (by direction)
        results += "\n\n=== PERIOD-BY-PERIOD RESULTS (by direction) ===\n"
        for period in range(1, 5):
            results += f"\nPeriod {period} ({15*(period-1)}-{15*period} minutes):\n"
            for direction, data in results_by_period_direction[period].items():
                results += f"  {direction}: {data['Total Vehicles']} vehicles, "
                results += f"{data['Total Stopped']} stopped, "
                results += f"{data['Avg Delay per Approach (sec)']:.1f}s avg delay\n"

        # Period Summaries (all directions combined)
        results += "\n\n=== PERIOD SUMMARIES (all directions) ===\n"
        for period in sorted(results_by_period_overall.keys()):
            data = results_by_period_overall[period]
            results += (
                f"Period {period}: total={data['Total Vehicles']}, stopped={data['Total Stopped']}, "
                f"delay={data['Total Delay (sec)']:.0f}s, avgStopped={data['Avg Delay per Stopped (sec)']:.1f}s, "
                f"avgApproach={data['Avg Delay per Approach (sec)']:.1f}s, %stopped={data['Percent Stopped']:.1f}%\n"
            )

        messagebox.showinfo("Delay Study Results", results)

    def save_data(self):
        if not self.data:
            messagebox.showwarning("No Data", "No data to save.")
            return
            
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if filename:
                # Create DataFrame from collected data, but override totals with UNIQUE split
                columns = [
                    "Time", "Interval", "Minute", "Direction",
                    "Stopped 0-15s", "Stopped 15-30s", "Stopped 30-45s", "Stopped 45-60s", "Total Stopped",
                    "Not Stopped 0-15s", "Not Stopped 15-30s", "Not Stopped 30-45s", "Not Stopped 45-60s",
                    "Total Not Stopped", "Total Volume"
                ]
                # Rebuild rows to ensure Total Stopped / Total Not Stopped are unique per minute
                rebuilt_rows = []
                for row in self.data:
                    period = row[0]
                    minute = row[1]
                    direction = row[2]
                    # Calculate actual time
                    actual_time = self.calculate_actual_time(period, minute)
                    # Unique split for this minute
                    usplit = self.unique_minute_map.get((period, direction, minute))
                    if usplit is None:
                        # Fallback to original totals if unique not present
                        u_st = row[7]
                        u_ns = row[12]
                    else:
                        u_st = int(usplit.get("stopped", 0))
                        u_ns = int(usplit.get("notstopped", 0))
                    u_total = u_st + u_ns
                    rebuilt_rows.append([
                        actual_time, period, minute, direction,
                        row[3], row[4], row[5], row[6], u_st,
                        row[8], row[9], row[10], row[11], u_ns, u_total
                    ])
                df = pd.DataFrame(rebuilt_rows, columns=columns)
                
                if filename.endswith('.xlsx'):
                    # Add study information at the top
                    info_df = pd.DataFrame([
                        ["Intersection Delay Study Raw Data"],
                        ["Date:", self.date_var.get()],
                        ["Intersection:", self.intersection_var.get()],
                        ["Weather:", self.weather_var.get()],
                        ["Start Time:", f"{self.start_hour_var.get()}:{self.start_minute_var.get()}:{self.start_second_var.get()}"],
                        [""]
                    ])
                    
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        # Write info
                        info_df.to_excel(writer, sheet_name='Data', index=False, header=False)
                        # Write data starting after info rows
                        df.to_excel(writer, sheet_name='Data', index=False, startrow=len(info_df))
                        
                        # Auto-adjust column widths
                        worksheet = writer.sheets['Data']
                        for idx, col in enumerate(df.columns):
                            max_length = max(
                                df[col].astype(str).apply(len).max(),
                                len(str(col))
                            )
                            worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
                
                elif filename.endswith('.csv'):
                    with open(filename, 'w', newline='') as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(["Intersection Delay Study Raw Data"])
                        writer.writerow(["Date:", self.date_var.get()])
                        writer.writerow(["Intersection:", self.intersection_var.get()])
                        writer.writerow(["Weather:", self.weather_var.get()])
                        writer.writerow(["Start Time:", f"{self.start_hour_var.get()}:{self.start_minute_var.get()}:{self.start_second_var.get()}"])
                        writer.writerow([])
                        writer.writerow(columns)
                        for row in rebuilt_rows:
                            writer.writerow(row)
                
                messagebox.showinfo("Success", f"Data saved successfully to {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving file: {str(e)}")

    def clear_data(self):
        if messagebox.askyesno("Clear Data", "Are you sure you want to clear all data?"):
            self.data.clear()
            for row in self.tree.get_children():
                self.tree.delete(row)
            self.period_var.set("1 (0-15 min)")
            self.minute_var.set("1")
            self.reset_counts()
            self.approach_volume.clear()
            self.approach_volume_var.set("0")
            self.unique_map.clear()
            self.current_unique_stopped = None
            self.current_unique_notstopped = None

    def generate_form2(self):
        """Generate Form 2 with aggregated 15-minute period data."""
        if not self.data:
            messagebox.showwarning("No Data", "Please collect some data first.")
            return
        
        # Create Form 2 window
        form2_window = Form2Window(self.root, self.data, self.date_var.get(), 
                                  self.intersection_var.get(), self.weather_var.get(),
                                  self.approach_volume, self.unique_map)

class Form2Window:
    def __init__(self, parent, data, date, intersection, weather, approach_volume_map, unique_map):
        self.data = data
        self.date = date
        self.intersection = intersection
        self.weather = weather
        self.approach_volume_map = approach_volume_map
        self.unique_map = unique_map
        
        # Create new window
        self.window = tk.Toplevel(parent)
        self.window.title("Form 2: Intersection Delay Study")
        self.window.geometry("1000x700")
        
        # Study Information Frame
        info_frame = ttk.LabelFrame(self.window, text="Study Information")
        info_frame.pack(fill="x", padx=10, pady=5)
        
        # Study details
        self.location_var = tk.StringVar(value=intersection)
        self.approach_var = tk.StringVar(value="All Approaches")
        self.movement_var = tk.StringVar(value="All Movements")
        self.lanes_var = tk.StringVar(value="All Lanes")
        self.delay_observer_var = tk.StringVar()
        self.count_observer_var = tk.StringVar()
        self.recorder_var = tk.StringVar()
        
        # Peak hour times
        self.begin_hour_var = tk.StringVar(value="8")
        self.begin_minute_var = tk.StringVar(value="00")
        self.end_hour_var = tk.StringVar(value="9")
        self.end_minute_var = tk.StringVar(value="00")
        
        # Layout study information
        ttk.Label(info_frame, text="Date:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
        ttk.Label(info_frame, text=date).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Location:").grid(row=0, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.location_var, width=20).grid(row=0, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Approach:").grid(row=1, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.approach_var, width=20).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Movement(s):").grid(row=1, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.movement_var, width=20).grid(row=1, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Lanes:").grid(row=2, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.lanes_var, width=20).grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Weather:").grid(row=2, column=2, padx=5, pady=2, sticky="e")
        ttk.Label(info_frame, text=weather).grid(row=2, column=3, padx=5, pady=2, sticky="w")
        
        # Peak hour
        ttk.Label(info_frame, text="Peak Hour:").grid(row=3, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.begin_hour_var, width=3).grid(row=3, column=1, padx=2, pady=2, sticky="w")
        ttk.Label(info_frame, text=":").grid(row=3, column=2, padx=0, pady=2)
        ttk.Entry(info_frame, textvariable=self.begin_minute_var, width=3).grid(row=3, column=3, padx=2, pady=2, sticky="w")
        ttk.Label(info_frame, text="to").grid(row=3, column=4, padx=5, pady=2)
        ttk.Entry(info_frame, textvariable=self.end_hour_var, width=3).grid(row=3, column=5, padx=2, pady=2, sticky="w")
        ttk.Label(info_frame, text=":").grid(row=3, column=6, padx=0, pady=2)
        ttk.Entry(info_frame, textvariable=self.end_minute_var, width=3).grid(row=3, column=7, padx=2, pady=2, sticky="w")
        
        # Observers
        ttk.Label(info_frame, text="Delay Observer:").grid(row=4, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.delay_observer_var, width=20).grid(row=4, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Count Observer:").grid(row=4, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.count_observer_var, width=20).grid(row=4, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(info_frame, text="Recorder:").grid(row=5, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(info_frame, textvariable=self.recorder_var, width=20).grid(row=5, column=1, padx=5, pady=2, sticky="w")
        
        # Form 2 Data Table
        table_frame = ttk.LabelFrame(self.window, text="Form 2: 15-Minute Period Data")
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side="right", fill="y")
        
        # Create treeview for Form 2 data
        self.form2_tree = ttk.Treeview(table_frame, 
            columns=("period", "time_range", "direction", "total_stopped", "total_notstopped", "total_volume", "observed_stopped", "observed_notstopped", "observed_volume"),
            show="headings",
            yscrollcommand=scrollbar.set)
        
        # Configure headers
        headers = {
            "period": "Interval",
            "time_range": "Time Range",
            "direction": "Direction",
            "total_stopped": "Total Stopped",
            "total_notstopped": "Total Not Stopped", 
            "total_volume": "Total Volume",
            "observed_stopped": "Observed Stopped (unique)",
            "observed_notstopped": "Observed Not Stopped (unique)",
            "observed_volume": "Observed Volume (unique)"
        }
        
        for col, heading in headers.items():
            self.form2_tree.heading(col, text=heading)
            self.form2_tree.column(col, width=120)
        
        self.form2_tree.pack(fill="both", expand=True, side="left")
        scrollbar.config(command=self.form2_tree.yview)
        
        # Populate Form 2 data
        self.populate_form2_data()
        
        # Calculated Results Frame
        results_frame = ttk.LabelFrame(self.window, text="Calculated Results")
        results_frame.pack(fill="x", padx=10, pady=5)
        
        # Results display
        self.results_text = tk.Text(results_frame, height=8, wrap="word")
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scrollbar.set)
        
        self.results_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        results_scrollbar.pack(side="right", fill="y", pady=5)
        
        # Calculate and display results
        self.calculate_form2_results()
        
        # Action buttons
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(button_frame, text="Export Form 2", 
                  command=self.export_form2).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", 
                  command=self.window.destroy).pack(side="right", padx=5)

    def populate_form2_data(self):
        """Populate Form 2 table with aggregated 15-minute period data."""
        # Aggregate data by period and direction
        period_data = {}
        
        for row in self.data:
            period = row[0]  # Period number
            direction = row[2]  # Direction
            total_stopped = row[7]  # Total stopped
            total_notstopped = row[12]  # Total not stopped
            
            if period not in period_data:
                period_data[period] = {}
            if direction not in period_data[period]:
                period_data[period][direction] = {"stopped": 0, "notstopped": 0}
            
            period_data[period][direction]["stopped"] += total_stopped
            period_data[period][direction]["notstopped"] += total_notstopped
        
        # Populate treeview
        for period in sorted(period_data.keys()):
            time_range = f"{15*(period-1):02d}:00-{15*period:02d}:00"
            for direction in ["North", "South", "East", "West"]:
                if direction in period_data[period]:
                    stopped = period_data[period][direction]["stopped"]
                    notstopped = period_data[period][direction]["notstopped"]
                    total = stopped + notstopped
                    u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                    observed_st = u["stopped"]
                    observed_ns = u["notstopped"]
                    observed = observed_st + observed_ns
                    
                    self.form2_tree.insert("", "end", values=(
                        period, time_range, direction, stopped, notstopped, total, observed_st, observed_ns, observed
                    ))

    def calculate_form2_results(self):
        """Calculate and display Form 2 results."""
        # Calculate total stopped vehicles (sum of all 15-second interval counts) across all periods and directions
        total_stopped_all_intervals = 0
        for row in self.data:
            # Sum all stopped counts: stopped_0 + stopped_15 + stopped_30 + stopped_45
            total_stopped_all_intervals += sum(row[3:7])  # indices 3-6 are stopped counts
        
        # Calculate total unique vehicles (approach volume) across all periods and directions
        total_vehicles = 0
        for period in range(1, 5):
            for direction in ["North", "South", "East", "West"]:
                u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                total_vehicles += (u["stopped"] + u["notstopped"])
        
        # Calculate metrics
        # Total Delay = Total Number Stopped at Time Intervals * 15 seconds
        total_delay_seconds = total_stopped_all_intervals * 15
        total_delay_hours = total_delay_seconds / 3600
        # Average Delay per Stopped Vehicle = Total Delay / Number Stopped (total stopped, not unique)
        avg_delay_stopped = total_delay_seconds / total_stopped_all_intervals if total_stopped_all_intervals > 0 else 0
        # Average Delay per Approach Vehicle = Total Delay / Total Approach Volume (unique vehicles)
        avg_delay_approach = total_delay_seconds / total_vehicles if total_vehicles > 0 else 0
        # Percent of Vehicles Stopped = Number Stopped (total) / Approach Volume (unique)
        percent_stopped = (total_stopped_all_intervals / total_vehicles * 100) if total_vehicles > 0 else 0
        
        # Format results
        results = f"""FORM 2 CALCULATED RESULTS
{'='*50}

Total Delay: {total_delay_seconds} vehicle-seconds
Total Delay: {total_delay_hours:.2f} vehicle-hours
Average Delay per Stopped Vehicle: {avg_delay_stopped:.2f} vehicle-seconds/vehicle
Average Delay per Approach Vehicle: {avg_delay_approach:.2f} vehicle-seconds/vehicle
Percent of Vehicles Stopped: {percent_stopped:.2f}%

STUDY SUMMARY
{'='*50}
Total Vehicles Observed: {total_vehicles}
Total Vehicles Stopped: {total_stopped_all_intervals}
Total Vehicles Not Stopped: {max(total_vehicles - total_stopped_all_intervals, 0)}
Study Duration: 60 minutes (4 x 15-minute periods)
Sampling Interval: 15 seconds

PERIOD BREAKDOWN
{'='*50}"""
        
        # Add period-by-period breakdown
        period_data = {}
        for row in self.data:
            period = row[0]
            if period not in period_data:
                period_data[period] = {"stopped_all_intervals": 0}
            # Sum all stopped counts: stopped_0 + stopped_15 + stopped_30 + stopped_45
            period_data[period]["stopped_all_intervals"] += sum(row[3:7])  # indices 3-6 are stopped counts
        
        for period in sorted(period_data.keys()):
            stopped_all_intervals = period_data[period]["stopped_all_intervals"]
            # unique total approach volume for this period across all directions
            total = 0
            for direction in ["North", "South", "East", "West"]:
                u = self.unique_map.get((period, direction), {"stopped": 0, "notstopped": 0})
                total += (u["stopped"] + u["notstopped"])
            period_delay = stopped_all_intervals * 15
            period_avg_delay = period_delay / total if total > 0 else 0
            
            results += f"""
Period {period} ({15*(period-1):02d}:00-{15*period:02d}:00):
  Total Vehicles: {total}
  Stopped: {stopped_all_intervals}
  Not Stopped: {max(total - stopped_all_intervals, 0)}
  Total Delay: {period_delay} vehicle-seconds
  Average Delay: {period_avg_delay:.2f} seconds/vehicle"""
        
        self.results_text.insert("1.0", results)
        self.results_text.config(state="disabled")

    def export_form2(self):
        """Export Form 2 to Excel file."""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile="Form2_Intersection_Delay_Study.xlsx"
            )
            if filename:
                # Prepare Form 2 data
                form2_data = []
                for item in self.form2_tree.get_children():
                    values = self.form2_tree.item(item)["values"]
                    form2_data.append(values)
                
                # Create DataFrame
                columns = ["Interval", "Time Range", "Direction", "Total Stopped", "Total Not Stopped", "Total Volume", "Observed Stopped (unique)", "Observed Not Stopped (unique)", "Observed Volume (unique)"]
                df = pd.DataFrame(form2_data, columns=columns)
                
                # Write to Excel
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # Study information
                    info_df = pd.DataFrame([
                        ["Form 2: Intersection Delay Study"],
                        ["Date:", self.date],
                        ["Location:", self.location_var.get()],
                        ["Approach:", self.approach_var.get()],
                        ["Movement(s):", self.movement_var.get()],
                        ["Lanes:", self.lanes_var.get()],
                        ["Weather:", self.weather],
                        ["Peak Hour:", f"{self.begin_hour_var.get()}:{self.begin_minute_var.get()} - {self.end_hour_var.get()}:{self.end_minute_var.get()}"],
                        ["Delay Observer:", self.delay_observer_var.get()],
                        ["Count Observer:", self.count_observer_var.get()],
                        ["Recorder:", self.recorder_var.get()],
                        [""]
                    ])
                    info_df.to_excel(writer, sheet_name='Form2', index=False, header=False)
                    
                    # Form 2 data
                    df.to_excel(writer, sheet_name='Form2', index=False, startrow=len(info_df))
                    
                    # Auto-adjust columns
                    worksheet = writer.sheets['Form2']
                    for idx, col in enumerate(df.columns):
                        worksheet.column_dimensions[chr(65 + idx)].width = 20
                
                messagebox.showinfo("Success", f"Form 2 exported to {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting Form 2: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    app = DelayStudyApp(root)
    root.mainloop()