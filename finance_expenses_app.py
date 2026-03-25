import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class ExpenseApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)

        self.master = master
        self.master.title("Finance Dashboard")
        self.master.geometry("1100x650")

        # ===== DARK THEME =====
        style = ttk.Style()
        style.theme_use("clam")

        self.master.configure(bg="#0f0f0f")

        style.configure("TFrame", background="#0f0f0f")
        style.configure("TLabel", background="#0f0f0f", foreground="white")

        style.configure("Treeview",
                        background="#1c1c1c",
                        foreground="white",
                        fieldbackground="#1c1c1c")

        style.configure("Treeview.Heading",
                        background="#2a2a2a",
                        foreground="white")

        # ===== DATA =====
        self.df = None
        self.expenses = {}
        self.detail_entries = {}
        self.city_expenses = {}

        self.income = 0.0
        self.net = 0.0

        self.categories = {
            "Monthly Taxes": ['SOFIYSKA VODA', 'OVERGAS', 'PB PERSONAL', 'YETTEL', 'ELEKTROHOLD'],
            "Revolut": ['REVOLUT'],
            "ATM Withdrawals": [],
            "Fuel": ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', 'SHELL'],
            "Food": ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET', 'LIDAL', 'MINIMARKET'],
            "Other": []
        }

        self.create_widgets()
        self.pack(fill="both", expand=True)

    # ================= UI =================

    def create_widgets(self):
        menubar = tk.Menu(self.master)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open", command=self.load_file)
        file_menu.add_command(label="Save Report", command=self.save_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        self.master.config(menu=menubar)

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True)

        # Tabs
        self.tab_data = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_data, text="Raw Data")
        self._build_raw_tab()

        self.tab_summary = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_summary, text="Summary")
        self._build_summary_tab()

        self.tab_cities = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_cities, text="Cities")
        self._build_cities_tab()

        self.tab_charts = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_charts, text="Charts")
        self._build_charts_tab()

        self.category_tabs = {}
        for cat in self.categories:
            frame = ttk.Frame(self.tabs)
            self.tabs.add(frame, text=cat)
            self._build_category_tab(frame, cat)

        self.status = ttk.Label(self, text="Open an .xls file to begin.",
                                relief="sunken", anchor="w")
        self.status.pack(fill="x", side="bottom")

    def _build_raw_tab(self):
        ttk.Button(self.tab_data, text="Browse…", command=self.load_file).pack(anchor="nw", pady=5)

        cols = ("Date", "Amount", "Method", "Description")
        self.tree_data = ttk.Treeview(self.tab_data, columns=cols, show="headings")

        for col in cols:
            anchor = "e" if col == "Amount" else "w"
            self.tree_data.heading(col, text=col)
            self.tree_data.column(col, anchor=anchor, width=200)

        self.tree_data.pack(fill="both", expand=True)

    def _build_summary_tab(self):
        self.tree_summary = ttk.Treeview(self.tab_summary,
                                         columns=("Category", "Total"),
                                         show="headings")

        self.tree_summary.heading("Category", text="Category")
        self.tree_summary.heading("Total", text="Amount")
        self.tree_summary.pack(fill="both", expand=True)

    def _build_cities_tab(self):
        self.tree_cities = ttk.Treeview(self.tab_cities,
                                        columns=("City", "Total"),
                                        show="headings")

        self.tree_cities.heading("City", text="City")
        self.tree_cities.heading("Total", text="Amount")
        self.tree_cities.pack(fill="both", expand=True)

    def _build_category_tab(self, frame, cat):
        tree = ttk.Treeview(frame,
                            columns=("Amount", "Description"),
                            show="headings")

        tree.heading("Amount", text="Amount")
        tree.heading("Description", text="Description")
        tree.pack(fill="both", expand=True)

        self.category_tabs[cat] = tree

    def _build_charts_tab(self):
        self.fig = Figure(figsize=(10, 5), facecolor="#0f0f0f")

        self.ax1 = self.fig.add_subplot(121)
        self.ax2 = self.fig.add_subplot(122)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_charts)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    # ================= LOGIC =================

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xls")])
        if not path:
            return

        try:
            self.df = pd.read_excel(path, sheet_name="Sheet")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.status.config(text=f"Loaded: {os.path.basename(path)}")
        self._populate_raw()
        self._analyze()

    def _populate_raw(self):
        self.tree_data.delete(*self.tree_data.get_children())

        for i in range(9, len(self.df)):
            row = self.df.iloc[i]
            if pd.notna(row.iloc[3]):
                self.tree_data.insert("", "end", values=(
                    row.iloc[1],
                    f"{row.iloc[3]:.2f}",
                    row.iloc[5],
                    row.iloc[7]
                ))

    def add_city_expense(self, city, amount):
        if not city:
            city = "SOFIA"
        self.city_expenses.setdefault(city, 0.0)
        self.city_expenses[city] += amount

    def _analyze(self):
        self.expenses = {cat: 0.0 for cat in self.categories}
        self.detail_entries = {cat: [] for cat in self.categories}
        self.city_expenses = {}

        self.income = 0.0

        city_variants = {
            'SOFIYA': 'SOFIA', 'SOFIA': 'SOFIA',
            'PLEVEN': 'PLEVEN', 'VARNA': 'VARNA',
            'BURGAS': 'BURGAS', 'PLOVDIV': 'PLOVDIV',
            'RUSE': 'RUSE', 'STARA ZAGORA': 'STARA ZAGORA',
            'SEVLIEVO': 'SEVLIEVO'
        }

        for i in range(9, len(self.df)):
            row = self.df.iloc[i]

            amount = row.iloc[3]
            credit = row.iloc[4]

            # Income
            if pd.notna(credit) and float(credit) > 0:
                self.income += float(credit)
                continue

            if pd.isna(amount):
                continue

            amount = float(amount)

            method = str(row.iloc[5]).upper() if pd.notna(row.iloc[5]) else ""
            desc = str(row.iloc[7]).upper() if pd.notna(row.iloc[7]) else ""

            city = None
            for var, norm in city_variants.items():
                if var in desc:
                    city = norm
                    break

            chosen = "Other"
            for cat, keywords in self.categories.items():
                if keywords and any(k in desc for k in keywords):
                    chosen = cat
                    break
                elif cat == "ATM Withdrawals" and "ATM" in method:
                    chosen = cat
                    break

            self.expenses[chosen] += amount
            self.detail_entries[chosen].append((amount, desc))

            self.add_city_expense(city, amount)

        total_expenses = sum(self.expenses.values())
        self.net = self.income - total_expenses

        # ===== SUMMARY =====
        self.tree_summary.delete(*self.tree_summary.get_children())

        for cat, total in self.expenses.items():
            self.tree_summary.insert("", "end", values=(cat, f"{total:.2f}"))

        self.tree_summary.insert("", "end", values=("", ""))
        self.tree_summary.insert("", "end", values=("Total Expenses", f"{total_expenses:.2f}"))
        self.tree_summary.insert("", "end", values=("Total Income", f"{self.income:.2f}"))
        self.tree_summary.insert("", "end", values=("NET", f"{self.net:.2f}"))

        # ===== CITIES =====
        self.tree_cities.delete(*self.tree_cities.get_children())

        for city, total in sorted(self.city_expenses.items(), key=lambda x: x[1], reverse=True):
            self.tree_cities.insert("", "end", values=(city, f"{total:.2f}"))

        # ===== DETAILS =====
        for cat, tree in self.category_tabs.items():
            tree.delete(*tree.get_children())
            for amt, desc in self.detail_entries[cat]:
                tree.insert("", "end", values=(f"{amt:.2f}", desc))

        # ===== CHARTS =====
        self._draw_charts()

    def _draw_charts(self):
        self.ax1.clear()
        self.ax2.clear()

        self.ax1.set_facecolor("#0f0f0f")
        self.ax2.set_facecolor("#0f0f0f")

        # Pie chart
        labels = [c for c, v in self.expenses.items() if v > 0]
        values = [v for v in self.expenses.values() if v > 0]

        if values:
            self.ax1.pie(values, labels=labels, autopct='%1.1f%%')
            self.ax1.set_title("Expenses by Category", color="white")

        # Bar chart
        cities = list(self.city_expenses.keys())
        amounts = list(self.city_expenses.values())

        if cities:
            self.ax2.bar(cities, amounts)
            self.ax2.set_title("Expenses by City", color="white")
            self.ax2.tick_params(axis='x', rotation=45)

        self.fig.tight_layout()
        self.canvas.draw()

    # ================= SAVE =================

    def save_report(self):
        if not self.expenses:
            messagebox.showwarning("No data", "Load data first.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".txt")
        if not path:
            return

        total_expenses = sum(self.expenses.values())

        with open(path, "w", encoding="utf-8") as f:
            f.write("== SUMMARY ==\n")
            for cat, total in self.expenses.items():
                f.write(f"{cat}: {total:.2f}\n")

            f.write("\n")
            f.write(f"Total Expenses: {total_expenses:.2f}\n")
            f.write(f"Total Income: {self.income:.2f}\n")
            f.write(f"NET: {self.net:.2f}\n\n")

            f.write("== CITIES ==\n")
            for city, total in self.city_expenses.items():
                f.write(f"{city}: {total:.2f}\n")

        messagebox.showinfo("Saved", "Report saved successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExpenseApp(root)
    root.mainloop()