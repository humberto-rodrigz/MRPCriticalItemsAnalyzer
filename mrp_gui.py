import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time
import webbrowser
import pandas as pd
from mrp_analyzer import analyze_mrp
from ttkbootstrap import Style
from ttkbootstrap.tooltip import ToolTip


class MRPGUI:
    def __init__(self, root):
        self.root = root
        self.style = Style("flatly")
        self.root.title("MRP Critical Items Analyzer")
        self.root.geometry("1200x800")
        self.root.minsize(900, 600)
        self.theme = "flatly"

        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Cálculo MRP")
        self.df_table = pd.DataFrame()
        self.current_page = 0
        self.page_size = 50

        self.compare_before = None
        self.compare_after = None

        self._build_ui()

    def _toggle_theme(self):
        self.theme = "darkly" if self.theme == "flatly" else "flatly"
        self.style.theme_use(self.theme)
        self._log(f"Theme changed to: {self.theme}")

    def _build_ui(self):
        topbar = ttk.Frame(self.root)
        topbar.pack(fill=tk.X, pady=2)
        theme_btn = ttk.Button(topbar, text="Toggle Theme", command=self._toggle_theme)
        theme_btn.pack(side=tk.RIGHT, padx=10)
        ToolTip(theme_btn, text="Switch between light and dark mode (Ctrl+T)")
        about_btn = ttk.Button(topbar, text="About", command=self._show_about)
        about_btn.pack(side=tk.RIGHT, padx=10)
        self.root.bind('<Control-t>', lambda e: self._toggle_theme())

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_analysis = ttk.Frame(self.notebook, padding=15)
        self.tab_table = ttk.Frame(self.notebook, padding=15)
        self.tab_compare = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.tab_analysis, text="Analysis")
        self.notebook.add(self.tab_table, text="Table")
        self.notebook.add(self.tab_compare, text="Comparison")

        self._build_analysis_tab()
        self._build_table_tab()
        self._build_compare_tab()

    def _build_analysis_tab(self):
        form = ttk.Labelframe(self.tab_analysis, text="Run Analysis", padding=10)
        form.pack(pady=10, fill=tk.X)

        ttk.Label(form, text="Excel File:").grid(row=0, column=0, sticky=tk.E)
        entry_file = ttk.Entry(form, textvariable=self.selected_file, width=60)
        entry_file.grid(row=0, column=1, padx=5)
        btn_browse = ttk.Button(form, text="Browse", command=self._browse_file)
        btn_browse.grid(row=0, column=2)
        ToolTip(btn_browse, text="Select the Excel file to analyze")

        ttk.Label(form, text="Sheet Name:").grid(row=1, column=0, sticky=tk.E, pady=5)
        entry_sheet = ttk.Entry(form, textvariable=self.sheet_name, width=30)
        entry_sheet.grid(row=1, column=1, sticky=tk.W, pady=5)
        ToolTip(entry_sheet, text="Enter the worksheet name (e.g., Cálculo MRP)")

        btn_run = ttk.Button(form, text="Run Analysis", command=self._run_analysis, bootstyle="success")
        btn_run.grid(row=2, column=0, columnspan=3, pady=10)
        ToolTip(btn_run, text="Start the MRP analysis")

        self.progress = ttk.Progressbar(form, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky=tk.EW)

        self.status_label = ttk.Label(form, text="", font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Label(self.tab_analysis, text="Log:").pack(anchor=tk.W, padx=10)
        log_frame = ttk.Frame(self.tab_analysis)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, bg="#f8f9fa", fg="#222")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Spreadsheets", "*.xlsx *.xls")])
        if file:
            self.selected_file.set(file)
            self._log(f"Selected file: {os.path.basename(file)}", "info")

    def _log(self, msg, level="info"):
        color = {"info": "#222", "success": "#155724", "error": "#721c24"}.get(level, "#222")
        self.log_text.insert(tk.END, f"{msg}\n", (level,))
        self.log_text.tag_config(level, foreground=color)
        self.log_text.see(tk.END)

    def _run_analysis(self):
        file = self.selected_file.get()
        sheet = self.sheet_name.get()

        if not os.path.exists(file):
            messagebox.showerror("Error", "File not found.")
            self._log("File not found.", "error")
            return

        if not self._validate_sheet(file, sheet):
            messagebox.showerror("Error", f"Sheet '{sheet}' not found.")
            self._log(f"Sheet '{sheet}' not found.", "error")
            return

        self.progress.start()
        self.status_label.config(text="Analyzing...", foreground="#007bff")
        self.root.update_idletasks()

        start = time.time()
        output_file = os.path.join(os.path.dirname(file), "itens_criticos.xlsx")
        count, error, _ = analyze_mrp(file, sheet, output_file)
        self.progress.stop()

        if error:
            self._log(f"Error: {error}", "error")
            self.status_label.config(text="Analysis failed.", foreground="#dc3545")
            messagebox.showerror("Analysis Error", error)
        else:
            elapsed = round(time.time() - start, 2)
            self._log(f"Completed in {elapsed}s. {count} critical items.", "success")
            self.status_label.config(text=f"Completed in {elapsed}s", foreground="#28a745")
            self._load_table(output_file)
            self.notebook.select(self.tab_table)
            if messagebox.askyesno("Success", "Do you want to open the generated file?"):
                webbrowser.open(output_file)

    def _validate_sheet(self, file, sheet):
        try:
            return sheet in pd.ExcelFile(file).sheet_names
        except Exception as e:
            self._log(f"Error validating sheet: {e}", "error")
            return False

    def _build_table_tab(self):
        filter_frame = ttk.Labelframe(self.tab_table, text="Filter & Export", padding=10)
        filter_frame.pack(fill=tk.X, pady=5)

        self.filter_column = tk.StringVar()
        self.filter_value = tk.StringVar()
        self.qtd_min = tk.StringVar()
        self.qtd_max = tk.StringVar()

        self.column_box = ttk.Combobox(filter_frame, textvariable=self.filter_column, state="readonly", width=30)
        self.column_box.pack(side=tk.LEFT, padx=5)
        ToolTip(self.column_box, text="Select column to filter")
        entry_filter = ttk.Entry(filter_frame, textvariable=self.filter_value, width=30)
        entry_filter.pack(side=tk.LEFT)
        ToolTip(entry_filter, text="Enter value to filter")

        ttk.Label(filter_frame, text="Min Qty:").pack(side=tk.LEFT, padx=2)
        entry_min = ttk.Entry(filter_frame, textvariable=self.qtd_min, width=6)
        entry_min.pack(side=tk.LEFT)
        ToolTip(entry_min, text="Minimum quantity to request")

        ttk.Label(filter_frame, text="Max Qty:").pack(side=tk.LEFT, padx=2)
        entry_max = ttk.Entry(filter_frame, textvariable=self.qtd_max, width=6)
        entry_max.pack(side=tk.LEFT)
        ToolTip(entry_max, text="Maximum quantity to request")

        btn_filter = ttk.Button(filter_frame, text="Apply Filter", command=self._apply_filter)
        btn_filter.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_filter, text="Apply filter to table")
        btn_reload = ttk.Button(filter_frame, text="Reload", command=self._load_table)
        btn_reload.pack(side=tk.LEFT)
        ToolTip(btn_reload, text="Reload table from file")

        btn_export_excel = ttk.Button(filter_frame, text="Export Excel", command=self._export_excel)
        btn_export_excel.pack(side=tk.RIGHT, padx=5)
        ToolTip(btn_export_excel, text="Export table to Excel file")
        btn_export_csv = ttk.Button(filter_frame, text="Export CSV", command=self._export_csv)
        btn_export_csv.pack(side=tk.RIGHT)
        ToolTip(btn_export_csv, text="Export table to CSV file")

        self.tree = ttk.Treeview(self.tab_table, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        nav_frame = ttk.Frame(self.tab_table)
        nav_frame.pack(fill=tk.X, pady=10)

        self.stats_label = ttk.Label(nav_frame, text="")
        self.stats_label.pack(side=tk.LEFT, padx=10)

        btn_frame = ttk.Frame(nav_frame)
        btn_frame.pack(side=tk.RIGHT)
        btn_prev = ttk.Button(btn_frame, text="Previous", command=self._prev_page)
        btn_prev.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_prev, text="Previous page")
        btn_next = ttk.Button(btn_frame, text="Next", command=self._next_page)
        btn_next.pack(side=tk.LEFT)
        ToolTip(btn_next, text="Next page")

    def _load_table(self, path=None):
        try:
            path = path or os.path.join(os.path.dirname(self.selected_file.get()), "itens_criticos.xlsx")
            self.df_table = pd.read_excel(path)
            self.column_box['values'] = list(self.df_table.columns)
            self.current_page = 0
            self._render_table()
        except Exception as e:
            self._log(f"Error loading table: {e}", "error")

    def _render_table(self):
        self.tree.delete(*self.tree.get_children())
        df = self.df_table.copy()
        start = self.current_page * self.page_size
        end = start + self.page_size
        page = df.iloc[start:end]

        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self._sort_column(c))
            self.tree.column(col, width=120, anchor="center")

        for _, row in page.iterrows():
            self.tree.insert("", tk.END, values=list(row))

        total = len(df)
        soma = df["QUANTIDADE A SOLICITAR"].sum() if "QUANTIDADE A SOLICITAR" in df.columns else 0
        media = round(df["QUANTIDADE A SOLICITAR"].mean(), 2) if "QUANTIDADE A SOLICITAR" in df.columns else 0
        top_forn = df["FORNECEDOR PRINCIPAL"].value_counts().idxmax() if "FORNECEDOR PRINCIPAL" in df.columns else "-"
        self.stats_label.config(text=f"Total: {total} | Sum: {soma} | Avg: {media} | Top Supplier: {top_forn}")

    def _apply_filter(self):
        df = self.df_table.copy()
        col = self.filter_column.get()
        val = self.filter_value.get().strip().lower()
        min_qtd = self.qtd_min.get()
        max_qtd = self.qtd_max.get()

        if col and val:
            df = df[df[col].astype(str).str.lower().str.contains(val)]

        if "QUANTIDADE A SOLICITAR" in df.columns:
            if min_qtd.isdigit():
                df = df[df["QUANTIDADE A SOLICITAR"] >= int(min_qtd)]
            if max_qtd.isdigit():
                df = df[df["QUANTIDADE A SOLICITAR"] <= int(max_qtd)]

        self.df_table = df
        self.current_page = 0
        self._render_table()

    def _sort_column(self, col):
        self.df_table.sort_values(by=col, ascending=True, inplace=True, ignore_index=True)
        self.current_page = 0
        self._render_table()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_table()

    def _next_page(self):
        if (self.current_page + 1) * self.page_size < len(self.df_table):
            self.current_page += 1
            self._render_table()

    def _export_csv(self):
        file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if file:
            self.df_table.to_csv(file, index=False)
            self._log(f"CSV saved: {file}", "success")
            messagebox.showinfo("Export", f"CSV file saved: {file}")

    def _export_excel(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if file:
            self.df_table.to_excel(file, index=False)
            self._log(f"Excel saved: {file}", "success")
            messagebox.showinfo("Export", f"Excel file saved: {file}")

    def _build_compare_tab(self):
        frame = ttk.Labelframe(self.tab_compare, text="Compare Analyses", padding=10)
        frame.pack(pady=10, fill=tk.X)

        btn_before = ttk.Button(frame, text="Select Previous Analysis", command=self._load_before)
        btn_before.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_before, text="Select the previous analysis Excel file")
        btn_after = ttk.Button(frame, text="Select Current Analysis", command=self._load_after)
        btn_after.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_after, text="Select the current analysis Excel file")
        btn_compare = ttk.Button(frame, text="Compare", command=self._compare_files, bootstyle="info")
        btn_compare.pack(side=tk.LEFT, padx=10)
        ToolTip(btn_compare, text="Compare the two analyses")

        self.compare_tree = ttk.Treeview(self.tab_compare, show="headings")
        self.compare_tree.pack(fill=tk.BOTH, expand=True)

    def _load_before(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_before = pd.read_excel(file)
            self._log(f"Previous analysis loaded: {os.path.basename(file)}", "info")

    def _load_after(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_after = pd.read_excel(file)
            self._log(f"Current analysis loaded: {os.path.basename(file)}", "info")

    def _compare_files(self):
        if self.compare_before is None or self.compare_after is None:
            messagebox.showwarning("Missing File", "Load both analyses to compare.")
            self._log("Both analyses must be loaded for comparison.", "error")
            return

        before = self.compare_before.set_index("CÓD")
        after = self.compare_after.set_index("CÓD")

        all_codes = sorted(set(before.index) | set(after.index))
        result = []

        for code in all_codes:
            row = {"CÓD": code}
            row["DESCRIÇÃO"] = after.at[code, "DESCRIÇÃOPROMOB"] if code in after.index else before.at[code, "DESCRIÇÃOPROMOB"]
            row["FORNECEDOR"] = after.at[code, "FORNECEDOR PRINCIPAL"] if code in after.index else before.at[code, "FORNECEDOR PRINCIPAL"]
            q_ant = before.at[code, "QUANTIDADE A SOLICITAR"] if code in before.index else 0
            q_atu = after.at[code, "QUANTIDADE A SOLICITAR"] if code in after.index else 0
            row["ANTERIOR"] = q_ant
            row["ATUAL"] = q_atu
            row["DIFERENÇA"] = q_atu - q_ant

            if code not in before.index:
                row["STATUS"] = "New"
            elif code not in after.index:
                row["STATUS"] = "Removed"
            elif q_ant != q_atu:
                row["STATUS"] = "Changed"
            else:
                row["STATUS"] = "Unchanged"

            result.append(row)

        df = pd.DataFrame(result)
        self.compare_tree.delete(*self.compare_tree.get_children())
        self.compare_tree["columns"] = list(df.columns)

        # Melhorias: colorir linhas por status e ajustar largura automática
        status_colors = {
            "New": "#d4edda",
            "Removed": "#f8d7da",
            "Changed": "#fff3cd",
            "Unchanged": "#f9f9f9"
        }
        for col in df.columns:
            self.compare_tree.heading(col, text=col)
            self.compare_tree.column(col, width=120, anchor="center")
        for _, row in df.iterrows():
            tag = row["STATUS"]
            self.compare_tree.insert("", tk.END, values=list(row), tags=(tag,))
        for status, color in status_colors.items():
            self.compare_tree.tag_configure(status, background=color)
        # Ajuste automático de largura
        for col in df.columns:
            max_len = max([len(str(x)) for x in df[col].values] + [len(col)])
            self.compare_tree.column(col, width=min(200, max(80, max_len * 10)))

    def _show_about(self):
        messagebox.showinfo(
            "About",
            "MRP Critical Items Analyzer\n\nDeveloped by Humberto Rodrigues.\nModern UI, color feedback, and Excel/CSV export.\n2025"
        )

def main():
    root = tk.Tk()
    app = MRPGUI(root)
    root.mainloop()

def set_style():
    style = Style("flatly")
    style.configure("TButton", padding=5, relief="flat")
    style.configure("TLabel", padding=5)
    style.configure("TEntry", padding=5)
    style.configure("TFrame", padding=10)
    return style

if __name__ == "__main__":
    main()
