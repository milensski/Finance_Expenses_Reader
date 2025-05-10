import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI

# Ensure your OpenAI API key is set in the environment
load_dotenv()


class ExpenseApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.status = None
        self.save_ai_btn = None
        self.ai_text = None
        self.ai_frame = None
        self.progress = None
        self.ai_button = None
        self.tab_cities = None
        self.tab_summary = None
        self.tab_data = None
        self.tabs = None
        self.category_tabs = None
        self.master = master
        self.master.title("Modern Finance Expenses App")
        self.master.geometry("1000x700")

        # Data containers
        self.df = None
        self.expenses = {}
        self.detail_entries = {}
        self.city_expenses = {}
        self.categories = {
            "Monthly Taxes": ['SOFIYSKA VODA', 'OVERGAS', 'PB PERSONAL', 'YETTEL', 'ELEKTROHOLD'],
            "Revolut": ['REVOLUT'],
            "ATM Withdrawals": [],  # matched via method
            "Fuel": ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', 'SHELL'],
            "Food": ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET'],
            "Other": []
        }

        self.create_widgets()
        self.pack(fill="both", expand=True)

    def create_widgets(self):
        # Menu bar
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

        # Tabs
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

        # Category-specific Tabs
        self.category_tabs = {}
        for cat in self.categories:
            frame = ttk.Frame(self.tabs)
            self.tabs.add(frame, text=cat)
            self._build_category_tab(frame, cat)

        # AI Insights controls
        self.ai_button = ttk.Button(self, text="Generate AI Insights", command=self.generate_ai_insights)
        self.ai_button.pack(side="top", pady=5)

        self.progress = ttk.Progressbar(self, mode='determinate', value=0)
        self.progress.pack(fill='x', pady=2)

        self.ai_frame = ttk.Frame(self)
        self.ai_frame.pack(fill="both", expand=True)

        self.ai_text = tk.Text(self.ai_frame, wrap='word', state='disabled')
        self.ai_text.pack(fill='both', pady=5)

        self.save_ai_btn = ttk.Button(self, text="Save AI Insights as TXT", command=self.save_ai_insights)
        self.save_ai_btn.pack(pady=5)
        self.save_ai_btn.config(state='disabled')

        # Status Bar
        self.status = ttk.Label(self, text="Welcome! Open an .xls file to begin.",
                                relief="sunken", anchor="w")
        self.status.pack(fill="x", side="bottom")

    def _build_raw_tab(self):
        btn = ttk.Button(self.tab_data, text="Browse…", command=self.load_file)
        btn.pack(anchor="nw", pady=5)
        columns = ("Date", "Amount", "Method", "Description")
        self.tree_data = ttk.Treeview(self.tab_data, columns=columns, show="headings")
        for col in columns:
            self.tree_data.heading(col, text=col)
            self.tree_data.column(col, anchor="e" if col == "Amount" else "w", width=150)
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
            date, amt, method, desc = row.iloc[1], row.iloc[3], row.iloc[5], row.iloc[7]
            if pd.notna(amt):
                self.tree_data.insert("", "end", values=(date, f"{amt:.2f}", method, desc))

    def _analyze(self):
        self.expenses = {cat: 0.0 for cat in self.categories}
        self.detail_entries = {cat: [] for cat in self.categories}
        self.city_expenses = {"SOFIA": 0.0}
        city_variants = {
            'SOFIYA': 'SOFIA', 'SOFIA': 'SOFIA', 'PLEVEN': 'PLEVEN', 'VARNA': 'VARNA',
            'BURGAS': 'BURGAS', 'PLOVDIV': 'PLOVDIV', 'RUSE': 'RUSE',
            'STARA ZAGORA': 'STARA ZAGORA', 'SEVLIEVO': 'SEVLIEVO'
        }
        for idx in range(9, len(self.df)):
            row = self.df.iloc[idx]
            amt = row.iloc[3]
            if pd.isna(amt): continue
            amt = float(amt)
            method = str(row.iloc[5]).upper() if pd.notna(row.iloc[5]) else ""
            desc = str(row.iloc[7]).upper() if pd.notna(row.iloc[7]) else ""
            city = next((n for v, n in city_variants.items() if v in desc), None)
            chosen = "Other"
            for cat, kws in self.categories.items():
                if kws and any(kw in desc for kw in kws):
                    chosen = cat
                    break
                if cat == "ATM Withdrawals" and "ATM" in method:
                    chosen = cat
                    break
            self.expenses[chosen] += amt
            self.detail_entries[chosen].append((amt, desc))
            if city:
                self.city_expenses.setdefault(city, 0.0)
                self.city_expenses[city] += amt
            elif chosen == "ATM Withdrawals":
                self.city_expenses['SOFIA'] += amt

        self._populate_summary()
        self._populate_cities()
        self._populate_categories()

    def _populate_summary(self):
        for i in self.tree_summary.get_children(): self.tree_summary.delete(i)
        for cat, total in self.expenses.items():
            self.tree_summary.insert("", "end", values=(cat, f"{total:.2f}"))
        self.tree_summary.insert("", "end", values=("", ""))
        self.tree_summary.insert("", "end", values=("Grand Total", f"{sum(self.expenses.values()):.2f}"))

    def _populate_cities(self):
        for i in self.tree_cities.get_children(): self.tree_cities.delete(i)
        for city, total in sorted(self.city_expenses.items()):
            self.tree_cities.insert("", "end", values=(city, f"{total:.2f}"))

    def _populate_categories(self):
        for cat, tree in self.category_tabs.items():
            for i in tree.get_children(): tree.delete(i)
            for amt, desc in self.detail_entries[cat]:
                tree.insert("", "end", values=(f"{amt:.2f}", desc))

    def save_report(self):
        if not self.expenses:
            messagebox.showwarning("No data", "Load data first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Text files", "*.txt")])
        if not path: return
        txt = self._compose_report_text()
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(txt)
            messagebox.showinfo("Saved", f"Report saved to:\n{path}")
            self.status.config(text=f"Report saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save report:\n{e}")

    def _compose_report_text(self) -> str:
        lines = ["== Summary by Category ==", "Category,Total (BGN)"]
        for cat, total in self.expenses.items():
            lines.append(f"{cat},{total:.2f}")
        lines += ["", f"Grand Total,{sum(self.expenses.values()):.2f}", "",
                  "== Detailed Expenses by Category =="]
        for cat, entries in self.detail_entries.items():
            lines.append(f"\n[{cat}]")
            lines.append("Amount (BGN),Description")
            for amt, desc in entries:
                lines.append(f"{amt:.2f},{desc}")
        lines += ["", "== Expenses by City ==", "City,Total (BGN)"]
        for city, total in sorted(self.city_expenses.items()):
            lines.append(f"{city},{total:.2f}")
        return "\n".join(lines)

    def save_ai_insights(self):
        content = self.ai_text.get("1.0", "end").strip()
        if not content:
            messagebox.showwarning("No AI content", "Generate AI insights first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Text files", "*.txt")])
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Saved", f"AI insights saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save AI insights:\n{e}")

    def generate_ai_insights(self):
        if not self.expenses:
            messagebox.showwarning("No data", "Load data first.")
            return

        # clear UI
        self.ai_text.config(state="normal")
        self.ai_text.delete("1.0", "end")
        self.ai_text.config(state="disabled")
        self.save_ai_btn.config(state="disabled")

        self.status.config(text="Generating AI insights...")
        self.progress.start(10)
        self.ai_button.config(state="disabled")

        def worker():
            report = self._compose_report_text()
            try:
                client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

                resp = client.responses.create(
                    model="gpt-4o",
                    instructions="You are a financial assistant.",
                    input=f"Provide breakdown of the report , key insights, and recommendations:\n{report}"
                )
                analysis = resp.output_text
                error = None
            except Exception as e:
                analysis = ""
                error = str(e)

            def on_complete():
                self.progress.stop()
                self.ai_button.config(state="normal")
                if error:
                    messagebox.showerror("AI Error", f"Failed to generate insights:\n{error}")
                    self.status.config(text="Ready")
                else:
                    self.ai_text.config(state="normal")
                    self.ai_text.insert("1.0", analysis)
                    self.ai_text.config(state="disabled")
                    self.save_ai_btn.config(state="normal")
                    self.status.config(text="AI insights generated.")

            self.master.after(0, on_complete)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    ExpenseApp(root)
    root.mainloop()
