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
        self.minute_var = tk.StringVar(value="0")
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

        # Stopped vehicle counts frame
        stopped_frame = ttk.LabelFrame(input_frame, text="Vehicles Stopped at Each Interval")
        stopped_frame.pack(fill="x", padx=5, pady=5)

        # Initialize variables for interval counts
        self.interval_vars = []
        interval_labels = ["+0 sec", "+15 sec", "+30 sec", "+45 sec"]
        
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
            "period": "Period",
            "minute": "Minute",
            "direction": "Direction",
            "stopped_0": "Stopped +0s",
            "stopped_15": "Stopped +15s",
            "stopped_30": "Stopped +30s",
            "stopped_45": "Stopped +45s",
            "total_stopped": "Total Stopped",
            "notstopped_0": "Not Stopped +0s",
            "notstopped_15": "Not Stopped +15s",
            "notstopped_30": "Not Stopped +30s",
            "notstopped_45": "Not Stopped +45s",
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
        ttk.Button(action_frame, text="Save Data", 
                  command=self.save_data).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Clear All Data", 
                  command=self.clear_data).pack(side="left", padx=5)

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
        if current_minute < 14:  # 0-14 minutes in each 15-minute period
            self.minute_var.set(str(current_minute + 1))
        else:
            # Move to next period
            current_period = int(self.period_var.get().split()[0])
            if current_period < 4:
                next_period = current_period + 1
                self.period_var.set(f"{next_period} ({15*(next_period-1)}-{15*next_period} min)")
                self.minute_var.set("0")
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
            
            total = total_stopped + total_notstopped
            
            # Create data tuple with all values including period
            entry_data = (
                period, minute, direction, 
                *stopped_counts, total_stopped,
                *notstopped_counts, total_notstopped,
                total
            )
            self.data.append(entry_data)
            
            # Add to treeview
            self.tree.insert("", "end", values=entry_data)

            # Auto-progress to next minute and reset counts
            self.next_minute()
            self.reset_counts()
            
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers.")

    def calculate_results(self):
        if not self.data:
            messagebox.showwarning("No Data", "Please add some entries first.")
            return

        # Calculate results by period and direction
        results_by_period_direction = {}
        overall_results_by_direction = {}
        
        for period in range(1, 5):
            results_by_period_direction[period] = {}
            for direction in ["North", "South", "East", "West"]:
                # Filter data for this period and direction (period is now index 0, direction is index 2)
                period_direction_data = [d for d in self.data if d[0] == period and d[2] == direction]
                if period_direction_data:
                    total_stopped = sum(d[7] for d in period_direction_data)  # Index 7 is total_stopped
                    total_notstopped = sum(d[12] for d in period_direction_data)  # Index 12 is total_notstopped
                    total_vehicles = total_stopped + total_notstopped
                    
                    if total_vehicles > 0:
                        total_delay = total_stopped * 15  # Total Delay formula
                        avg_delay_stopped = total_delay / total_stopped if total_stopped > 0 else 0
                        avg_delay_approach = total_delay / total_vehicles
                        percent_stopped = (total_stopped / total_vehicles) * 100
                        
                        results_by_period_direction[period][direction] = {
                            "Total Vehicles": total_vehicles,
                            "Total Stopped": total_stopped,
                            "Total Delay (sec)": total_delay,
                            "Avg Delay per Stopped (sec)": avg_delay_stopped,
                            "Avg Delay per Approach (sec)": avg_delay_approach,
                            "Percent Stopped": percent_stopped
                        }
        
        # Calculate overall results by direction
        for direction in ["North", "South", "East", "West"]:
            direction_data = [d for d in self.data if d[2] == direction]
            if direction_data:
                total_stopped = sum(d[7] for d in direction_data)  # Index 7 is total_stopped
                total_notstopped = sum(d[12] for d in direction_data)  # Index 12 is total_notstopped
                total_vehicles = total_stopped + total_notstopped
                
                if total_vehicles > 0:
                    total_delay = total_stopped * 15  # Total Delay formula
                    avg_delay_stopped = total_delay / total_stopped if total_stopped > 0 else 0
                    avg_delay_approach = total_delay / total_vehicles
                    percent_stopped = (total_stopped / total_vehicles) * 100
                    
                    overall_results_by_direction[direction] = {
                        "Total Vehicles": total_vehicles,
                        "Total Stopped": total_stopped,
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
                
                # Create DataFrame for period-by-period results
                period_results_data = []
                for period in range(1, 5):
                    for direction, data in results_by_period_direction[period].items():
                        period_results_data.append({
                            "Period": period,
                            "Direction": direction,
                            **data
                        })
                
                period_df = pd.DataFrame(period_results_data)
                
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
                    
                    # Auto-adjust columns width
                    for sheet_name in ['Overall Results', 'Period Results']:
                        worksheet = writer.sheets[sheet_name]
                        if sheet_name == 'Overall Results':
                            df_to_adjust = overall_df
                        else:
                            df_to_adjust = period_df
                        for idx, col in enumerate(df_to_adjust.columns):
                            worksheet.column_dimensions[chr(65 + idx)].width = 20

                messagebox.showinfo("Success", f"Results saved to {filename}")
                self.show_results_popup(overall_results_by_direction, results_by_period_direction)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving results: {str(e)}")

    def show_results_popup(self, overall_results_by_direction, results_by_period_direction):
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
        
        # Period-by-Period Results
        results += "\n\n=== PERIOD-BY-PERIOD RESULTS ===\n"
        for period in range(1, 5):
            results += f"\nPeriod {period} ({15*(period-1)}-{15*period} minutes):\n"
            for direction, data in results_by_period_direction[period].items():
                results += f"  {direction}: {data['Total Vehicles']} vehicles, "
                results += f"{data['Total Stopped']} stopped, "
                results += f"{data['Avg Delay per Approach (sec)']:.1f}s avg delay\n"

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
                # Create DataFrame from collected data
                columns = [
                    "Period", "Minute", "Direction",
                    "Stopped +0s", "Stopped +15s", "Stopped +30s", "Stopped +45s", "Total Stopped",
                    "Not Stopped +0s", "Not Stopped +15s", "Not Stopped +30s", "Not Stopped +45s",
                    "Total Not Stopped", "Total Volume"
                ]
                
                df = pd.DataFrame(self.data, columns=columns)
                
                if filename.endswith('.xlsx'):
                    # Add study information at the top
                    info_df = pd.DataFrame([
                        ["Intersection Delay Study Raw Data"],
                        ["Date:", self.date_var.get()],
                        ["Intersection:", self.intersection_var.get()],
                        ["Weather:", self.weather_var.get()],
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
                        writer.writerow([])
                        writer.writerow(columns)
                        for row in self.data:
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
            self.minute_var.set("0")
            self.reset_counts()

if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    app = DelayStudyApp(root)
    root.mainloop()