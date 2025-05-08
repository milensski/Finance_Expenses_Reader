import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


class ExpenseApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.master = master
        self.master.title("Modern Finance Expenses App")
        self.master.geometry("1000x600")

        # Data
        self.df = None

        # Totals and details
        self.expenses = {}
        self.detail_entries = {}

        # City totals
        self.city_expenses = {}

        # Define categories in desired priority order
        self.categories = {
            "Monthly Taxes": ['SOFIYSKA VODA', 'OVERGAS', 'PB PERSONAL', 'YETTEL', 'ELEKTROHOLD'],
            "Revolut": ['REVOLUT'],
            "ATM Withdrawals": [],  # will match by method
            "Fuel": ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', 'SHELL'],
            "Food": ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET'],
            "Other": []  # everything else
        }

        self.create_widgets()
        self.pack(fill="both", expand=True)

    def create_widgets(self):
        # — Menu Bar —
        menubar = tk.Menu(self.master)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open", command=self.load_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Save Report", command=self.save_report, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit, accelerator="Ctrl+Q")
        menubar.add_cascade(label="File", menu=file_menu)
        self.master.config(menu=menubar)
        self.master.bind_all("<Control-o>", lambda e: self.load_file())
        self.master.bind_all("<Control-s>", lambda e: self.save_report())
        self.master.bind_all("<Control-q>", lambda e: self.master.quit())

        # — Notebook (Tabs) —
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True)

        # Raw Data Tab
        self.tab_data = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_data, text="Raw Data")
        self._build_raw_tab()

        # Summary Tab
        self.tab_summary = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_summary, text="Summary")
        self._build_summary_tab()

        # Cities Tab
        self.tab_cities = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_cities, text="Cities")
        self._build_cities_tab()

        # One tab per category
        self.category_tabs = {}
        for cat in self.categories:
            frame = ttk.Frame(self.tabs)
            self.tabs.add(frame, text=cat)
            self._build_category_tab(frame, cat)

        # — Status bar —
        self.status = ttk.Label(self, text="Welcome! Open an .xls file to begin.",
                                relief="sunken", anchor="w")
        self.status.pack(fill="x", side="bottom")

    def _build_raw_tab(self):
        btn = ttk.Button(self.tab_data, text="Browse…", command=self.load_file)
        btn.pack(anchor="nw", pady=5)

        columns = ("Date", "Amount", "Method", "Description")
        self.tree_data = ttk.Treeview(self.tab_data, columns=columns, show="headings")
        for col in columns:
            anchor = "e" if col == "Amount" else "w"
            width = 100 if col == "Amount" else 200
            self.tree_data.heading(col, text=col)
            self.tree_data.column(col, anchor=anchor, width=width)
        self.tree_data.pack(fill="both", expand=True)

    def _build_summary_tab(self):
        self.tree_summary = ttk.Treeview(self.tab_summary,
                                         columns=("Category", "Total"),
                                         show="headings")
        self.tree_summary.heading("Category", text="Category")
        self.tree_summary.heading("Total", text="Total (BGN)")
        self.tree_summary.column("Category", anchor="w", width=300)
        self.tree_summary.column("Total", anchor="e", width=100)
        self.tree_summary.pack(fill="both", expand=True, pady=10)

    def _build_cities_tab(self):
        self.tree_cities = ttk.Treeview(self.tab_cities,
                                        columns=("City", "Total"),
                                        show="headings")
        self.tree_cities.heading("City", text="City")
        self.tree_cities.heading("Total", text="Total (BGN)")
        self.tree_cities.column("City", anchor="w", width=300)
        self.tree_cities.column("Total", anchor="e", width=100)
        self.tree_cities.pack(fill="both", expand=True, pady=10)

    def _build_category_tab(self, frame, cat):
        tree = ttk.Treeview(frame,
                            columns=("Amount", "Description"),
                            show="headings")
        tree.heading("Amount", text="Amount (BGN)")
        tree.heading("Description", text="Description")
        tree.column("Amount", anchor="e", width=100)
        tree.column("Description", anchor="w", width=600)
        tree.pack(fill="both", expand=True, pady=10)
        self.category_tabs[cat] = tree

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel 97-2003 Workbook", "*.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.df = pd.read_excel(path, sheet_name="Sheet")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")
            return

        self.status.config(text=f"Loaded: {os.path.basename(path)}")
        self._populate_raw()
        self._analyze()

    def _populate_raw(self):
        for item in self.tree_data.get_children():
            self.tree_data.delete(item)
        for idx in range(9, len(self.df)):
            row = self.df.iloc[idx]
            date = row.iloc[1]
            amt = row.iloc[3]
            method = row.iloc[5]
            desc = row.iloc[7]
            if pd.notna(amt):
                self.tree_data.insert("", "end",
                                      values=(date, f"{amt:.2f}", method, desc))

    def _analyze(self):
        # reset all trackers
        self.expenses = {cat: 0.0 for cat in self.categories}
        self.detail_entries = {cat: [] for cat in self.categories}
        self.city_expenses = {"SOFIA": 0.0}  # ensure Sofia key exists

        # city name normalization
        city_variants = {
            'SOFIYA': 'SOFIA', 'SOFIA': 'SOFIA',
            'PLEVEN': 'PLEVEN', 'VARNA': 'VARNA',
            'BURGAS': 'BURGAS', 'PLOVDIV': 'PLOVDIV',
            'RUSE': 'RUSE', 'STARA ZAGORA': 'STARA ZAGORA',
            'SEVLIEVO': 'SEVLIEVO'
        }

        # classify each transaction
        for idx in range(9, len(self.df)):
            row = self.df.iloc[idx]
            amt = row.iloc[3]
            if pd.isna(amt):
                continue
            amt = float(amt)
            method = str(row.iloc[5]).upper() if pd.notna(row.iloc[5]) else ""
            desc = str(row.iloc[7]).upper() if pd.notna(row.iloc[7]) else ""

            # extract city if present
            city = None
            for var, norm in city_variants.items():
                if var in desc:
                    city = norm
                    break

            # pick category in defined order
            chosen = "Other"
            for cat, keywords in self.categories.items():
                if keywords:
                    if any(kw in desc for kw in keywords):
                        chosen = cat
                        break
                else:
                    # ATM withdrawals
                    if cat == "ATM Withdrawals" and "ATM" in method:
                        chosen = cat
                        break

            # accumulate totals and details
            self.expenses[chosen] += amt
            self.detail_entries[chosen].append((amt, desc))

            # accumulate by city
            if city:
                self.city_expenses.setdefault(city, 0.0)
                self.city_expenses[city] += amt
            elif chosen == "ATM Withdrawals":
                # no city → attribute ATM withdrawal to Sofia
                self.city_expenses['SOFIA'] += amt

        # — Populate Summary tab —
        # clear old rows
        for item in self.tree_summary.get_children():
            self.tree_summary.delete(item)
        # insert category totals
        for cat, total in self.expenses.items():
            self.tree_summary.insert("", "end", values=(cat, f"{total:.2f}"))
        # blank separator
        self.tree_summary.insert("", "end", values=("", ""))
        # grand total
        grand_total = sum(self.expenses.values())
        self.tree_summary.insert("", "end", values=("Grand Total", f"{grand_total:.2f}"))

        # — Populate Cities tab —
        for item in self.tree_cities.get_children():
            self.tree_cities.delete(item)
        for city, total in sorted(self.city_expenses.items()):
            self.tree_cities.insert("", "end", values=(city, f"{total:.2f}"))

        # — Populate Category tabs —
        for cat, tree in self.category_tabs.items():
            for item in tree.get_children():
                tree.delete(item)
            for amt, desc in self.detail_entries[cat]:
                tree.insert("", "end", values=(f"{amt:.2f}", desc))

    def save_report(self):
        if not self.expenses:
            messagebox.showwarning("No data", "Load data first.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Text files", "*.txt")])
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("Category,Total (BGN)\n")
                for cat, total in self.expenses.items():
                    f.write(f"{cat},{total:.2f}\n")
                f.write("\nExpenses per City:\nCity,Total (BGN)\n")
                for city, total in sorted(self.city_expenses.items()):
                    f.write(f"{city},{total:.2f}\n")
            messagebox.showinfo("Saved", f"Report saved to:\n{path}")
            self.status.config(text=f"Report saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save report:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    ExpenseApp(root)
    root.mainloop()
