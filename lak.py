import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pulp
import sys
import traceback
from time import time
import threading
import numpy as np

class PackingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dorm Packing Planner")
        self.root.geometry("960x540")
        self.root.minsize(960, 540)
        self.items = []
        self.dp_running = False
        self.solver_var = tk.StringVar(value="IP")

        style = ttk.Style()
        style.theme_use('alt')
        style.configure("TLabel", foreground="black", background="white", font=("Helvetica", 12))
        style.configure("TButton", foreground="black", background="#e0e0e0", font=("Helvetica", 12))
        style.map("TButton", foreground=[("active", "blue")], background=[("active", "#d0d0d0")])
        style.configure("Treeview", foreground="black", background="white", font=("Helvetica", 10))
        style.configure("TEntry", foreground="black", background="white", font=("Helvetica", 10))
        style.configure("TText", foreground="black", background="white", font=("Helvetica", 10))

        self.progress = ttk.Label(self.root, text="Processing... 0s", style="TLabel")
        self.setup_gui()

    def setup_gui(self):
        input_frame = ttk.LabelFrame(self.root, text="Item Details", style="TLabel")
        input_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

        columns = ("Name", "Value", "Weight", "Volume", "Age", "Dep%", "Type")
        self.tree = ttk.Treeview(input_frame, columns=columns, show="headings", style="Treeview")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.grid(row=0, column=0, columnspan=2, sticky="nsew")
        input_frame.columnconfigure(0, weight=1)

        self.load_button = ttk.Button(input_frame, text="Load Excel", command=self.load_excel, style="TButton")
        self.load_button.grid(row=1, column=0, pady=5, sticky="ew")

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

        solver_frame = ttk.LabelFrame(self.root, text="Solver", style="TLabel")
        solver_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        ttk.Radiobutton(solver_frame, text="Integer Programming", variable=self.solver_var, value="IP").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(solver_frame, text="Dynamic Programming", variable=self.solver_var, value="DP").grid(row=0, column=1, padx=5)

        output_frame = ttk.LabelFrame(self.root, text="Packing Plan", style="TLabel")
        output_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        self.output_text = tk.Text(output_frame, height=10, width=60, bg="white", fg="black", font=("Helvetica", 10))
        self.output_text.grid(row=0, column=0, sticky="nsew")
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)

        self.solve_button = ttk.Button(self.root, text="Solve", command=self.start_solve, style="TButton")
        self.solve_button.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        self.export_button = ttk.Button(self.root, text="Export Plan", command=self.export, style="TButton")
        self.export_button.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

        self.progress.grid(row=6, column=0, pady=5)
        self.progress.grid_remove()

        for i in range(7):
            self.root.rowconfigure(i, weight=1)
        self.root.columnconfigure(0, weight=1)

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
                for col, dtype in expected_columns.items():
                    if col not in df.columns:
                        raise ValueError(f"Missing column: {col}")
                    if dtype == float:
                        df[col] = pd.to_numeric(df[col], errors="coerce")
                        if df[col].isna().any() or (df[col] < 0).any():
                            raise ValueError(f"Invalid or negative values in column: {col}")
                    else:
                        df[col] = df[col].astype(str).fillna("")

                self.items = df.to_dict("records")
                self.tree.delete(*self.tree.get_children())
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
        except Exception as e:
            traceback.print_exc()
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, f"Error loading Excel: {str(e)}\n")

    def start_solve(self):
        if not self.dp_running:
            self.dp_running = True
            self.progress.grid()
            self.root.update()
            threading.Thread(target=self.solve_thread, daemon=True).start()

    def solve_thread(self):
        try:
            if not self.items:
                raise ValueError("No items loaded")

            W_cabin = int(float(self.cabin_w.get()))
            U_cabin = int(float(self.cabin_v.get()))
            W_check = int(float(self.check_w.get()))
            U_check = int(float(self.check_v.get()))

            values = [max(0, item["Monetary Value (INR)"] * (1 - item["Depreciation per Year (%)"] / 100) ** item["Age (years)"])
                      for item in self.items]
            weights = [int(item["Weight (kg)"]) for item in self.items]
            volumes = [int(item["Volume (liters)"]) for item in self.items]
            compat = [{
                "cabin": item["Airline Baggage Type"] == "Hand Baggage",
                "check-in": item["Airline Baggage Type"] == "Check-in Baggage",
                "movers": item["Item"] not in ["Laptop", "Smartphone", "Headphones", "Mirror"]
            } for item in self.items]

            start_time = time()
            if self.solver_var.get() == "IP":
                plan = self.solve_ip(values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check)
            else:
                plan = self.solve_dp(values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check)
            total_value = sum(values[i] for i, item in enumerate(self.items)
                              if item["Item"] in (plan["cabin"] + plan["check-in"] + plan["movers"]))
            elapsed = time() - start_time
            self.root.after(0, lambda: self.update_result(plan, total_value, elapsed))
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, lambda: self.update_error(str(e)))
        finally:
            self.dp_running = False

    def solve_ip(self, values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check):
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

    def solve_dp(self, values, weights, volumes, compat, W_cabin, U_cabin, W_check, U_check):
        n = len(values)
        def knapsack_dp(max_w, max_v, allowed):
            dp = np.zeros((max_w + 1, max_v + 1))
            keep = [[set() for _ in range(max_v + 1)] for _ in range(max_w + 1)]
            for i in range(n):
                if not allowed[i]:
                    continue
                w, v, val = weights[i], volumes[i], values[i]
                for cw in range(max_w, w - 1, -1):
                    for cv in range(max_v, v - 1, -1):
                        if dp[cw - w][cv - v] + val > dp[cw][cv]:
                            dp[cw][cv] = dp[cw - w][cv - v] + val
                            keep[cw][cv] = keep[cw - w][cv - v].copy()
                            keep[cw][cv].add(i)
            max_val = 0
            best_set = set()
            for cw in range(max_w + 1):
                for cv in range(max_v + 1):
                    if dp[cw][cv] > max_val:
                        max_val = dp[cw][cv]
                        best_set = keep[cw][cv]
            return best_set

        assigned = set()
        plan = {"cabin": [], "check-in": [], "movers": []}
        allowed_cabin = [compat[i]["cabin"] for i in range(n)]
        cabin_set = knapsack_dp(W_cabin, U_cabin, [a and (i not in assigned) for i, a in enumerate(allowed_cabin)])
        for i in cabin_set:
            plan["cabin"].append(self.items[i]["Item"])
            assigned.add(i)
        allowed_check = [compat[i]["check-in"] for i in range(n)]
        check_set = knapsack_dp(W_check, U_check, [a and (i not in assigned) for i, a in enumerate(allowed_check)])
        for i in check_set:
            plan["check-in"].append(self.items[i]["Item"])
            assigned.add(i)
        for i in range(n):
            if i not in assigned and compat[i]["movers"]:
                plan["movers"].append(self.items[i]["Item"])
        return plan

    def update_result(self, plan, total_value, elapsed):
        self.progress.grid_remove()
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, f"Packing Plan (Total Value: {total_value:.2f} INR) [Solved in {elapsed:.2f}s]:\n\n")
        self.output_text.insert(tk.END, f"Cabin ({len(plan['cabin'])} items):\n" + "\n".join(plan["cabin"]) + "\n\n")
        self.output_text.insert(tk.END, f"Check-in ({len(plan['check-in'])} items):\n" + "\n".join(plan["check-in"]) + "\n\n")
        self.output_text.insert(tk.END, f"Movers ({len(plan['movers'])} items):\n" + "\n".join(plan["movers"]) + "\n")
        self.root.update()

    def update_error(self, error_msg):
        self.progress.grid_remove()
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, f"Error solving: {error_msg}\n")
        self.root.update()

    def export(self):
        try:
            with open("packing_plan_result.txt", "w") as f:
                f.write(self.output_text.get(1.0, tk.END))
            self.output_text.insert(tk.END, "\nPlan exported to packing_plan_result.txt\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"\nError exporting: {str(e)}\n")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = PackingGUI(root)
        root.mainloop()
    except Exception as e:
        traceback.print_exc()
        tk.Tk().withdraw()
        messagebox.showerror("Error", f"Failed to start GUI: {str(e)}")
