import pandas as pd
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class ExpenseApp(tb.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.master = master
        self.pack(fill=BOTH, expand=YES)

        self.df = None
        self.expenses = {}
        self.city_expenses = {}
        self.income = 0.0
        self.net = 0.0

        self.category_transactions = {}

        self.categories = {
            "Monthly Taxes": ['SOFIYSKA VODA', 'OVERGAS', 'PB PERSONAL', 'YETTEL', 'ELEKTROHOLD'],
            "Revolut": ['REVOLUT'],
            "ATM Withdrawals": [],
            "Fuel": ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', 'SHELL'],
            "Food": ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET', 'LIDAL', 'MINIMARKET'],
            "Other": []
        }

        self.create_widgets()

    # ================= UI =================

    def create_widgets(self):
        top_frame = tb.Frame(self)
        top_frame.pack(fill=X, pady=5)

        tb.Button(top_frame, text="Open File", bootstyle="primary", command=self.load_file).pack(side=LEFT, padx=5)
        tb.Button(top_frame, text="Save Report", bootstyle="success", command=self.save_report).pack(side=LEFT)

        # Main notebook
        self.tabs = tb.Notebook(self)
        self.tabs.pack(fill=BOTH, expand=YES)

        self.tab_summary = tb.Frame(self.tabs)
        self.tab_cities = tb.Frame(self.tabs)
        self.tab_charts = tb.Frame(self.tabs)

        self.tabs.add(self.tab_summary, text="Summary")
        self.tabs.add(self.tab_cities, text="Cities")
        self.tabs.add(self.tab_charts, text="Charts")

        # Category tabs inside Summary
        self.category_tabs = tb.Notebook(self.tab_summary)
        self.category_tabs.pack(fill=BOTH, expand=YES)

        self.category_trees = {}

        self._build_summary()
        self._build_cities()
        self._build_charts()

    def _build_category_tabs(self):
        for tab in self.category_tabs.tabs():
            self.category_tabs.forget(tab)

        self.category_trees = {}

        for cat in self.categories.keys():
            frame = tb.Frame(self.category_tabs)
            self.category_tabs.add(frame, text=cat)

            tree = tb.Treeview(frame, columns=("Desc", "Amount"), show="headings")
            tree.heading("Desc", text="Description")
            tree.heading("Amount", text="Amount")

            tree.column("Desc", anchor="w", width=400)
            tree.column("Amount", anchor="e", width=120)

            tree.pack(fill=BOTH, expand=YES)

            self.category_trees[cat] = tree

    def _build_summary(self):
        self.tree_summary = tb.Treeview(self.tab_summary, columns=("Category", "Amount"), show="headings")
        self.tree_summary.heading("Category", text="Category")
        self.tree_summary.heading("Amount", text="Amount")

        self.tree_summary.column("Category", anchor="w", width=250)
        self.tree_summary.column("Amount", anchor="e", width=120)

        self.tree_summary.pack(fill=BOTH, expand=YES)

    def _build_cities(self):
        self.tree_cities = tb.Treeview(self.tab_cities, columns=("City", "Amount"), show="headings")
        self.tree_cities.heading("City", text="City")
        self.tree_cities.heading("Amount", text="Amount")

        self.tree_cities.pack(fill=BOTH, expand=YES)

    def _build_charts(self):
        self.fig = Figure(figsize=(10, 5), facecolor="#1e1e1e")
        self.ax1 = self.fig.add_subplot(121)
        self.ax2 = self.fig.add_subplot(122)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_charts)
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=YES)

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

        self._analyze()

    def _analyze(self):
        self.expenses = {cat: 0.0 for cat in self.categories}
        self.city_expenses = {}
        self.income = 0.0
        self.category_transactions = {cat: [] for cat in self.categories}

        for i in range(9, len(self.df)):
            row = self.df.iloc[i]

            amount = row.iloc[3]
            credit = row.iloc[4]
            payment_method = row.iloc[5]
            description = row.iloc[7]

            # income
            if pd.notna(credit) and float(credit) > 0:
                self.income += float(credit)
                continue

            if pd.isna(amount):
                continue

            amount = float(amount)

            desc_str = str(description).upper() if isinstance(description, str) else ""
            pay_str = str(payment_method).upper() if isinstance(payment_method, str) else ""

            # city detection
            if "SOFIA" in desc_str:
                city = "SOFIA"
            elif "PLEVEN" in desc_str:
                city = "PLEVEN"
            elif "BURGAS" in desc_str:
                city = "BURGAS"
            else:
                city = "OTHER"

            # classification
            if any(kw in desc_str for kw in self.categories["Monthly Taxes"]):
                category = "Monthly Taxes"
            elif 'REVOLUT' in desc_str or 'REVOLUT' in pay_str:
                category = "Revolut"
            elif 'ATM' in pay_str:
                category = "ATM Withdrawals"
            elif any(comp in desc_str for comp in self.categories["Fuel"]):
                category = "Fuel"
            elif any(comp in desc_str for comp in self.categories["Food"]):
                category = "Food"
            else:
                category = "Other"

            self.expenses[category] += amount
            self.category_transactions[category].append((desc_str, amount))

            if city:
                self.city_expenses.setdefault(city, 0)
                self.city_expenses[city] += amount

        total_expenses = sum(self.expenses.values())
        self.net = self.income - total_expenses

        self._update_summary()
        self._update_cities()
        self._draw_charts()

        self._build_category_tabs()
        self._update_category_tabs()

    # ================= UI UPDATE =================

    def _update_summary(self):
        self.tree_summary.delete(*self.tree_summary.get_children())

        for cat, val in self.expenses.items():
            self.tree_summary.insert("", "end", values=(cat, f"{val:.2f}"))

        self.tree_summary.insert("", "end", values=("", ""))
        self.tree_summary.insert("", "end", values=("Expenses", f"{sum(self.expenses.values()):.2f}"))
        # self.tree_summary.insert("", "end", values=("Income", f"{self.income:.2f}"))
        self.tree_summary.insert("", "end", values=("NET", f"{self.net:.2f}"))

    def _update_cities(self):
        self.tree_cities.delete(*self.tree_cities.get_children())

        for city, val in self.city_expenses.items():
            self.tree_cities.insert("", "end", values=(city, f"{val:.2f}"))

    def _update_category_tabs(self):
        for cat, tree in self.category_trees.items():
            tree.delete(*tree.get_children())

            for desc, amount in self.category_transactions.get(cat, []):
                tree.insert("", "end", values=(desc, f"{amount:.2f}"))

    # ================= CHARTS =================

    def _draw_charts(self):
        self.ax1.clear()
        self.ax2.clear()

        self.ax1.set_facecolor("#1e1e1e")
        self.ax2.set_facecolor("#1e1e1e")

        labels = list(self.expenses.keys())
        values = list(self.expenses.values())

        self.ax1.pie(values, labels=labels, autopct='%1.1f%%', textprops={'color': 'white'})
        self.ax1.set_title("Expenses by Category", color="white")

        cities = list(self.city_expenses.keys())
        amounts = list(self.city_expenses.values())

        self.ax2.bar(cities, amounts)
        self.ax2.set_title("Expenses by City", color="white")

        self.ax2.tick_params(axis='x', colors='white')
        self.ax2.tick_params(axis='y', colors='white')

        self.fig.tight_layout()
        self.canvas.draw()

    # ================= SAVE =================

    def save_report(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt")
        if not path:
            return

        with open(path, "w", encoding="utf-8") as f:
            f.write("=== FINANCE REPORT ===\n\n")

            f.write("CATEGORY BREAKDOWN:\n")
            for item in self.tree_summary.get_children():
                values = self.tree_summary.item(item, "values")
                if values[0] != "":
                    f.write(f"{values[0]:<20} {values[1]}\n")

            f.write("\n")
            # f.write(f"Income: {self.income:.2f}\n")
            f.write(f"Total Expenses: {sum(self.expenses.values()):.2f}\n")
            f.write(f"NET: {self.net:.2f}\n")

        messagebox.showinfo("Saved", "Report saved!")


# ================= RUN =================

if __name__ == "__main__":
    app = tb.Window(themename="darkly")
    app.title("Finance Dashboard")
    app.geometry("1100x650")

    ExpenseApp(app)

    app.mainloop()