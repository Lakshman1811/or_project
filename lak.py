import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pulp
import numpy as np
import sys
import traceback
from time import time
import threading

class PackingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dorm Packing Planner")
        self.root.geometry("960x540")  # Half-screen size
        self.root.minsize(960, 540)  # Enforce minimum dimensions
        self.items = []
        self.dp_result = None
        self.dp_running = False
        # Set a default style for visibility
        style = ttk.Style()
        style.theme_use('alt')  # Use 'alt' theme
        style.configure("TLabel", foreground="black", background="white", font=("Helvetica", 12))
        style.configure("TButton", foreground="black", background="#e0e0e0", font=("Helvetica", 12))
        style.map("TButton", foreground=[("active", "blue")], background=[("active", "#d0d0d0")])
        style.configure("Treeview", foreground="black", background="white", font=("Helvetica", 10))
        style.configure("TEntry", foreground="black", background="white", font=("Helvetica", 10))
        style.configure("TText", foreground="black", background="white", font=("Helvetica", 10))
        self.progress = ttk.Label(self.root, text="Processing... 0s", style="TLabel")
        try:
            self.setup_gui()
            print("GUI initialized successfully")
        except Exception as e:
            print(f"Error initializing GUI: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to initialize GUI: {str(e)}")
            sys.exit(1)

    def setup_gui(self):
        # Item Input Frame
        input_frame = ttk.LabelFrame(self.root, text="Item Details", style="TLabel")
        input_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        # Treeview for items
        columns = ("Name", "Value", "Weight", "Volume", "Age", "Dep%", "Type")
        self.tree = ttk.Treeview(input_frame, columns=columns, show="headings", style="Treeview")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.grid(row=0, column=0, columnspan=2, sticky="nsew")
        input_frame.columnconfigure(0, weight=1)
        # Load Excel Button
        self.load_button = ttk.Button(input_frame, text="Load Excel", command=self.load_excel, style="TButton")
        self.load_button.grid(row=1, column=0, pady=5, sticky="ew")
        print(f"Load Excel button created at ({self.load_button.winfo_x()}, {self.load_button.winfo_y()})")
        # Suitcase Input
        suitcase_frame = ttk.LabelFrame(self.root, text="Suitcase Constraints", style="TLabel")
        suitcase_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        ttk.Label(suitcase_frame, text="Cabin Weight (kg):", style="TLabel").grid(row=0, column=0)
        self.cabin_w = ttk.Entry(suitcase_frame, width=10, style="TEntry")
        self.cabin_w.insert(0, "7")
        self.cabin_w.grid(row=0, column=1)
        ttk.Label(suitcase_frame, text="Cabin Volume (L):", style="TLabel").grid(row=0, column=2)
        self.cabin_v = ttk.Entry(suitcase_frame, width=10, style="TEntry")
        self.cabin_v.insert(0, "46")
        self.cabin_v.grid(row=0, column=3)
        ttk.Label(suitcase_frame, text="Check-in Weight (kg):", style="TLabel").grid(row=1, column=0)
        self.check_w = ttk.Entry(suitcase_frame, width=10, style="TEntry")
        self.check_w.insert(0, "23")
        self.check_w.grid(row=1, column=1)
        ttk.Label(suitcase_frame, text="Check-in Volume (L):", style="TLabel").grid(row=1, column=2)
        self.check_v = ttk.Entry(suitcase_frame, width=10, style="TEntry")
        self.check_v.insert(0, "100")
        self.check_v.grid(row=1, column=3)
        # Solver Selection
        solver_frame = ttk.LabelFrame(self.root, text="Solver", style="TLabel")
        solver_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.solver_var = tk.StringVar(value="IP")
        ttk.Radiobutton(solver_frame, text="Dynamic Programming", variable=self.solver_var, value="DP", style="TRadiobutton").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(solver_frame, text="Integer Programming", variable=self.solver_var, value="IP", style="TRadiobutton").grid(row=0, column=1, padx=5)
        # Output Frame
        output_frame = ttk.LabelFrame(self.root, text="Packing Plan", style="TLabel")
        output_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        self.output_text = tk.Text(output_frame, height=10, width=60, bg="white", fg="black", font=("Helvetica", 10))
        self.output_text.grid(row=0, column=0, sticky="nsew")
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        # Solve Button
        self.solve_button = ttk.Button(self.root, text="Solve", command=self.start_solve, style="TButton")
        self.solve_button.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
        print(f"Solve button created at ({self.solve_button.winfo_x()}, {self.solve_button.winfo_y()})")
        # Export Button
        self.export_button = ttk.Button(self.root, text="Export Plan", command=self.export, style="TButton")
        self.export_button.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
        print(f"Export button created at ({self.export_button.winfo_x()}, {self.export_button.winfo_y()})")
        # Progress Label (hidden by default)
        self.progress.grid(row=6, column=0, pady=5)
        self.progress.grid_remove()  # Hide until needed
        # Configure root to expand properly
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)  # Item frame
        self.root.rowconfigure(1, weight=1)  # Suitcase frame
        self.root.rowconfigure(2, weight=1)  # Solver frame
        self.root.rowconfigure(3, weight=2)  # Output frame
        self.root.rowconfigure(4, weight=1)  # Solve button
        self.root.rowconfigure(5, weight=1)  # Export button
        self.root.rowconfigure(6, weight=1)  # Progress label

    def load_excel(self):
        try:
            file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if file:
                df = pd.read_excel(file, engine="openpyxl")
                expected_columns = {
                    "Item": str,
                    "Monetary Value (INR)": float,
                    "Weight (kg)": float,
                    "Volume (liters)": float,
                    "Age (years)": float,
                    "Depreciation per Year (%)": float,
                    "Airline Baggage Type": str
                }
                missing_cols = [col for col in expected_columns if col not in df.columns]
                if missing_cols:
                    self.output_text.delete(1.0, tk.END)
                    self.output_text.insert(tk.END, f"Error: Missing columns: {', '.join(missing_cols)}")
                    return
                for col, dtype in expected_columns.items():
                    if dtype == float:
                        df[col] = pd.to_numeric(df[col], errors="coerce")
                        if df[col].isna().any():
                            self.output_text.delete(1.0, tk.END)
                            self.output_text.insert(tk.END, f"Error: Invalid or missing values in '{col}'")
                            return
                        if (df[col] < 0).any():  # Changed to < 0 to allow zero
                            self.output_text.delete(1.0, tk.END)
                            self.output_text.insert(tk.END, f"Error: Negative values in '{col}'")
                            return
                    else:
                        df[col] = df[col].astype(str).fillna("")
                self.items = df.to_dict("records")
                for item in self.tree.get_children():
                    self.tree.delete(item)
                for item in self.items:
                    self.tree.insert("", "end", values=(
                        item["Item"],
                        item["Monetary Value (INR)"],
                        item["Weight (kg)"],
                        item["Volume (liters)"],
                        item["Age (years)"],
                        item["Depreciation per Year (%)"],
                        item["Airline Baggage Type"]
                    ))
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, f"Loaded {len(self.items)} items successfully.\n")
                print(f"Loaded {len(self.items)} items from {file}")
        except Exception as e:
            print(f"Error loading Excel: {str(e)}")
            traceback.print_exc()
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, f"Error loading Excel: {str(e)}")

    def start_solve(self):
        if not self.dp_running:
            self.dp_running = True
            self.progress.grid()
            self.root.update()
            threading.Thread(target=self.solve_thread, daemon=True).start()

    def solve_thread(self):
        try:
            print("Solve thread started")
            if not self.items:
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, "Error: No items loaded.")
                self.root.after(0, lambda: self.update_error("No items loaded"))
                return
            W_cabin = int(float(self.cabin_w.get()))
            U_cabin = int(float(self.cabin_v.get()))
            W_check = int(float(self.check_w.get()))
            U_check = int(float(self.check_v.get()))
            if W_cabin <= 0 or U_cabin <= 0 or W_check <= 0 or U_check <= 0:
                raise ValueError("Constraints must be positive integers.")
            solver = self.solver_var.get()
            values = [max(0, item["Monetary Value (INR)"] * (1 - item["Depreciation per Year (%)"]/100) ** item["Age (years)"])
                      for item in self.items]
            weights = [int(item["Weight (kg)"]) for item in self.items]
            volumes = [int(item["Volume (liters)"]) for item in self.items]
            compat = [{"cabin": item["Airline Baggage Type"] == "Hand Baggage",
                       "check-in": item["Airline Baggage Type"] == "Check-in Baggage",
                       "movers": item["Item"] not in ["Laptop", "Smartphone", "Headphones", "Mirror"]}
                      for item in self.items]
            start_time = time()
            if solver == "DP":
                max_value, selected_items = self.solve_dp(values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check)
                elapsed = time() - start_time
                self.root.after(0, lambda: self.update_result_dp(max_value, selected_items, elapsed))
            else:
                plan = self.solve_ip(values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check)
                total_value = sum(values[i] for i, item in enumerate(self.items)
                                if item["Item"] in (plan["cabin"] + plan["check-in"] + plan["movers"]))
                self.root.after(0, lambda: self.update_result_ip(plan, total_value))
        except ValueError as ve:
            print(f"Value Error in solve: {str(ve)}")
            traceback.print_exc()
            self.root.after(0, lambda: self.update_error(str(ve)))
        except Exception as e:
            print(f"Error in solve: {str(e)}")
            traceback.print_exc()
            self.root.after(0, lambda: self.update_error(str(e)))
        finally:
            self.dp_running = False

    def solve_dp(self, values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check):
        try:
            n = len(values)
            max_w = min(W_cabin + 1, 20)  # Cap at 20 to reduce complexity
            max_v = min(U_cabin + 1, 50)  # Cap at 50 to reduce complexity
            dp = {}  # (i, wc, vc, wk, vk) -> max value
            keep = {}  # (i, wc, vc, wk, vk) -> decision (None, "cabin", "check-in", "movers")
            def dp_key(i, wc, vc, wk, vk):
                return (i, wc, vc, wk, vk)
            for i in range(n + 1):
                for wc in range(max_w):
                    for vc in range(max_v):
                        for wk in range(W_check + 1):
                            for vk in range(U_check + 1):
                                if i == 0:
                                    dp[dp_key(i, wc, vc, wk, vk)] = 0 if wc <= W_cabin and vc <= U_cabin and wk <= W_check and vk <= U_check else float('-inf')
                                    keep[dp_key(i, wc, vc, wk, vk)] = None
                                else:
                                    prev = dp.get(dp_key(i - 1, wc, vc, wk, vk), float('-inf'))
                                    curr_val, curr_w, curr_v = values[i - 1], weights[i - 1], volumes[i - 1]
                                    curr_comp = compat[i - 1]
                                    dp[dp_key(i, wc, vc, wk, vk)] = prev
                                    keep[dp_key(i, wc, vc, wk, vk)] = None
                                    if curr_comp["cabin"] and wc >= curr_w and vc >= curr_v and wc <= W_cabin and vc <= U_cabin:
                                        new_val = dp.get(dp_key(i - 1, wc - curr_w, vc - curr_v, wk, vk), float('-inf')) + curr_val
                                        if new_val > dp[dp_key(i, wc, vc, wk, vk)]:
                                            dp[dp_key(i, wc, vc, wk, vk)] = new_val
                                            keep[dp_key(i, wc, vc, wk, vk)] = "cabin"
                                    if curr_comp["check-in"] and wk >= curr_w and vk >= curr_v and wk <= W_check and vk <= U_check:
                                        new_val = dp.get(dp_key(i - 1, wc, vc, wk - curr_w, vk - curr_v), float('-inf')) + curr_val
                                        if new_val > dp[dp_key(i, wc, vc, wk, vk)]:
                                            dp[dp_key(i, wc, vc, wk, vk)] = new_val
                                            keep[dp_key(i, wc, vc, wk, vk)] = "check-in"
                                    if curr_comp["movers"]:
                                        new_val = dp.get(dp_key(i - 1, wc, vc, wk, vk), float('-inf')) + curr_val
                                        if new_val > dp[dp_key(i, wc, vc, wk, vk)]:
                                            dp[dp_key(i, wc, vc, wk, vk)] = new_val
                                            keep[dp_key(i, wc, vc, wk, vk)] = "movers"
            max_value = dp[dp_key(n, W_cabin, U_cabin, W_check, U_check)]
            selected = {"cabin": [], "check-in": [], "movers": []}
            i, wc, vc, wk, vk = n, W_cabin, U_cabin, W_check, U_check
            while i > 0:
                decision = keep[dp_key(i, wc, vc, wk, vk)]
                if decision:
                    item_name = self.items[i - 1]["Item"]
                    if decision == "cabin":
                        selected["cabin"].append(item_name)
                        wc -= weights[i - 1]
                        vc -= volumes[i - 1]
                    elif decision == "check-in":
                        selected["check-in"].append(item_name)
                        wk -= weights[i - 1]
                        vk -= volumes[i - 1]
                    elif decision == "movers":
                        selected["movers"].append(item_name)
                i -= 1
            return max_value, selected
        except RecursionError:
            raise Exception("Dynamic Programming exceeded recursion limit; try fewer items or Integer Programming.")
        except Exception as e:
            raise Exception(f"DP error: {str(e)}")

    def solve_ip(self, values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check):
        try:
            prob = pulp.LpProblem("Packing", pulp.LpMaximize)
            x_cabin = [pulp.LpVariable(f"x_c_{i}", cat="Binary") for i in range(len(values))]
            x_check = [pulp.LpVariable(f"x_k_{i}", cat="Binary") for i in range(len(values))]
            x_movers = [pulp.LpVariable(f"x_m_{i}", cat="Binary") for i in range(len(values))]
            prob += pulp.lpSum(values[i] * (x_cabin[i] + x_check[i] + x_movers[i]) for i in range(len(values)))
            for i in range(len(values)):
                prob += x_cabin[i] + x_check[i] + x_movers[i] <= 1
                if not compat[i]["cabin"]:
                    prob += x_cabin[i] == 0
                if not compat[i]["check-in"]:
                    prob += x_check[i] == 0
                if not compat[i]["movers"]:
                    prob += x_movers[i] == 0
            prob += pulp.lpSum(weights[i] * x_cabin[i] for i in range(len(values))) <= W_cabin
            prob += pulp.lpSum(volumes[i] * x_cabin[i] for i in range(len(values))) <= U_cabin
            prob += pulp.lpSum(weights[i] * x_check[i] for i in range(len(values))) <= W_check
            prob += pulp.lpSum(volumes[i] * x_check[i] for i in range(len(values))) <= U_check
            prob.solve(pulp.PULP_CBC_CMD(msg=0))
            if pulp.LpStatus[prob.status] != "Optimal":
                raise ValueError(f"Solver status: {pulp.LpStatus[prob.status]}")
            plan = {"cabin": [], "check-in": [], "movers": []}
            for i in range(len(values)):
                if pulp.value(x_cabin[i]) == 1:
                    plan["cabin"].append(self.items[i]["Item"])
                elif pulp.value(x_check[i]) == 1:
                    plan["check-in"].append(self.items[i]["Item"])
                elif pulp.value(x_movers[i]) == 1:
                    plan["movers"].append(self.items[i]["Item"])
            return plan
        except Exception as e:
            raise Exception(f"IP solver failed: {str(e)}")

    def update_result_dp(self, max_value, selected_items, elapsed):
        self.progress.config(text=f"Processing... {elapsed:.2f}s")
        self.root.update()
        if elapsed > 5:
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, f"Maximum Value: {max_value:.2f} INR (Note: DP took {elapsed:.2f} seconds)\n\n")
            self.output_text.insert(tk.END, f"Cabin ({len(selected_items['cabin'])} items):\n" + "\n".join(selected_items['cabin']) + "\n\n")
            self.output_text.insert(tk.END, f"Check-in ({len(selected_items['check-in'])} items):\n" + "\n".join(selected_items['check-in']) + "\n\n")
            self.output_text.insert(tk.END, f"Movers ({len(selected_items['movers'])} items):\n" + "\n".join(selected_items['movers']) + "\n")
            self.output_text.insert(tk.END, "Consider using Integer Programming for faster results.")
        else:
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, f"Maximum Value: {max_value:.2f} INR\n\n")
            self.output_text.insert(tk.END, f"Cabin ({len(selected_items['cabin'])} items):\n" + "\n".join(selected_items['cabin']) + "\n\n")
            self.output_text.insert(tk.END, f"Check-in ({len(selected_items['check-in'])} items):\n" + "\n".join(selected_items['check-in']) + "\n\n")
            self.output_text.insert(tk.END, f"Movers ({len(selected_items['movers'])} items):\n" + "\n".join(selected_items['movers']) + "\n")
        self.progress.grid_remove()
        self.dp_running = False
        self.root.update()

    def update_result_ip(self, plan, total_value):
        self.progress.grid_remove()
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, f"Packing Plan (Total Value: {total_value:.2f} INR):\n\n")
        self.output_text.insert(tk.END, f"Cabin ({len(plan['cabin'])} items):\n" + "\n".join(plan["cabin"]) + "\n\n")
        self.output_text.insert(tk.END, f"Check-in ({len(plan['check-in'])} items):\n" + "\n".join(plan["check-in"]) + "\n\n")
        self.output_text.insert(tk.END, f"Movers ({len(plan['movers'])} items):\n" + "\n".join(plan["movers"]) + "\n")
        self.dp_running = False
        self.root.update()

    def update_error(self, error_msg):
        self.progress.grid_remove()
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, f"Error solving: {error_msg}\n")
        self.dp_running = False
        self.root.update()

    def export(self):
        try:
            if self.solver_var.get() == "DP":
                with open("dynamic_programming_result.txt", "w") as f:
                    f.write(self.output_text.get(1.0, tk.END))
                self.output_text.insert(tk.END, "\nPlan exported to dynamic_programming_result.txt\n")
            else:
                with open("integer_programming_result.txt", "w") as f:
                    f.write(self.output_text.get(1.0, tk.END))
                self.output_text.insert(tk.END, "\nPlan exported to integer_programming_result.txt\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"\nError exporting: {str(e)}\n")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = PackingGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting GUI: {str(e)}")
        traceback.print_exc()
        tk.Tk().withdraw()
        messagebox.showerror("Error", f"Failed to start GUI: {str(e)}")