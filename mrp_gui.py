import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time
import webbrowser
import pandas as pd
from mrp_analyzer import analyze_mrp
from ttkbootstrap import Style

class MRPGUI:
    def __init__(self, root):
        self.root = root
        self.style = Style("flatly")
        self.root.title("Analisador de Itens Críticos MRP")
        self.root.geometry("1200x800")

        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Cálculo MRP")
        self.df_table = pd.DataFrame()
        self.current_page = 0
        self.page_size = 50

        self.compare_before = None
        self.compare_after = None

        self._build_ui()

    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_analysis = ttk.Frame(self.notebook, padding=15)
        self.tab_table = ttk.Frame(self.notebook, padding=15)
        self.tab_compare = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.tab_analysis, text="Análise")
        self.notebook.add(self.tab_table, text="Tabela")
        self.notebook.add(self.tab_compare, text="Comparação")

        self._build_analysis_tab()
        self._build_table_tab()
        self._build_compare_tab()

    # --- ABA ANÁLISE ---
    def _build_analysis_tab(self):
        form = ttk.Frame(self.tab_analysis)
        form.pack(pady=10)

        ttk.Label(form, text="Arquivo Excel:").grid(row=0, column=0, sticky=tk.E)
        ttk.Entry(form, textvariable=self.selected_file, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(form, text="Selecionar", command=self._browse_file).grid(row=0, column=2)

        ttk.Label(form, text="Nome da Aba:").grid(row=1, column=0, sticky=tk.E, pady=5)
        ttk.Entry(form, textvariable=self.sheet_name, width=30).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Button(form, text="Executar Análise", command=self._run_analysis).grid(row=2, column=0, columnspan=3, pady=10)

        self.progress = ttk.Progressbar(form, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky=tk.EW)

        self.status_label = ttk.Label(form, text="")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Label(self.tab_analysis, text="Log:").pack(anchor=tk.W, padx=10)
        log_frame = ttk.Frame(self.tab_analysis)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
        if file:
            self.selected_file.set(file)
            self._log(f"Arquivo selecionado: {os.path.basename(file)}")

    def _log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def _run_analysis(self):
        file = self.selected_file.get()
        sheet = self.sheet_name.get()

        if not os.path.exists(file):
            messagebox.showerror("Erro", "Arquivo não encontrado.")
            return

        if not self._validate_sheet(file, sheet):
            messagebox.showerror("Erro", f"A aba '{sheet}' não foi encontrada.")
            return

        self.progress.start()
        self.status_label.config(text="Analisando...")
        self.root.update_idletasks()

        start = time.time()
        output_file = os.path.join(os.path.dirname(file), "itens_criticos.xlsx")
        count, error, _ = analyze_mrp(file, sheet, output_file)
        self.progress.stop()

        if error:
            self._log(f"Erro: {error}")
            messagebox.showerror("Erro de Análise", error)
        else:
            elapsed = round(time.time() - start, 2)
            self._log(f"Concluído em {elapsed}s. {count} itens críticos.")
            self.status_label.config(text=f"Concluído em {elapsed}s")
            self._load_table(output_file)
            if messagebox.askyesno("Sucesso", "Deseja abrir o arquivo gerado?"):
                webbrowser.open(output_file)

    def _validate_sheet(self, file, sheet):
        try:
            return sheet in pd.ExcelFile(file).sheet_names
        except Exception as e:
            self._log(f"Erro ao validar aba: {e}")
            return False

    # --- ABA TABELA ---
    def _build_table_tab(self):
        filter_frame = ttk.Frame(self.tab_table)
        filter_frame.pack(fill=tk.X, pady=5)

        self.filter_column = tk.StringVar()
        self.filter_value = tk.StringVar()
        self.qtd_min = tk.StringVar()
        self.qtd_max = tk.StringVar()

        self.column_box = ttk.Combobox(filter_frame, textvariable=self.filter_column, state="readonly", width=30)
        self.column_box.pack(side=tk.LEFT, padx=5)
        ttk.Entry(filter_frame, textvariable=self.filter_value, width=30).pack(side=tk.LEFT)

        ttk.Label(filter_frame, text="Qtd Mín:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(filter_frame, textvariable=self.qtd_min, width=6).pack(side=tk.LEFT)

        ttk.Label(filter_frame, text="Qtd Máx:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(filter_frame, textvariable=self.qtd_max, width=6).pack(side=tk.LEFT)

        ttk.Button(filter_frame, text="Filtrar", command=self._apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="Recarregar", command=self._load_table).pack(side=tk.LEFT)

        ttk.Button(filter_frame, text="Exportar Excel", command=self._export_excel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(filter_frame, text="Exportar CSV", command=self._export_csv).pack(side=tk.RIGHT)

        self.tree = ttk.Treeview(self.tab_table, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        nav_frame = ttk.Frame(self.tab_table)
        nav_frame.pack(fill=tk.X, pady=10)

        self.stats_label = ttk.Label(nav_frame, text="")
        self.stats_label.pack(side=tk.LEFT, padx=10)

        btn_frame = ttk.Frame(nav_frame)
        btn_frame.pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Anterior", command=self._prev_page).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Próximo", command=self._next_page).pack(side=tk.LEFT)

    def _load_table(self, path=None):
        try:
            path = path or os.path.join(os.path.dirname(self.selected_file.get()), "itens_criticos.xlsx")
            self.df_table = pd.read_excel(path)
            self.column_box['values'] = list(self.df_table.columns)
            self.current_page = 0
            self._render_table()
        except Exception as e:
            self._log(f"Erro ao carregar tabela: {e}")

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
        self.stats_label.config(text=f"Total: {total} | Soma: {soma} | Média: {media} | Fornecedor: {top_forn}")

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
            self._log(f"CSV salvo: {file}")

    def _export_excel(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if file:
            self.df_table.to_excel(file, index=False)
            self._log(f"Excel salvo: {file}")

    # --- ABA COMPARAÇÃO ---
    def _build_compare_tab(self):
        frame = ttk.Frame(self.tab_compare)
        frame.pack(pady=10, fill=tk.X)

        ttk.Button(frame, text="Selecionar Análise Anterior", command=self._load_before).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Selecionar Análise Atual", command=self._load_after).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Comparar", command=self._compare_files).pack(side=tk.LEFT, padx=10)

        self.compare_tree = ttk.Treeview(self.tab_compare, show="headings")
        self.compare_tree.pack(fill=tk.BOTH, expand=True)

    def _load_before(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_before = pd.read_excel(file)
            self._log(f"Análise anterior carregada: {os.path.basename(file)}")

    def _load_after(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_after = pd.read_excel(file)
            self._log(f"Análise atual carregada: {os.path.basename(file)}")

    def _compare_files(self):
        if self.compare_before is None or self.compare_after is None:
            messagebox.showwarning("Faltando Arquivo", "Carregue as duas análises para comparar.")
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
                row["STATUS"] = "Novo"
            elif code not in after.index:
                row["STATUS"] = "Removido"
            elif q_ant != q_atu:
                row["STATUS"] = "Alterado"
            else:
                row["STATUS"] = "Inalterado"

            result.append(row)

        df = pd.DataFrame(result)
        self.compare_tree.delete(*self.compare_tree.get_children())
        self.compare_tree["columns"] = list(df.columns)

        for col in df.columns:
            self.compare_tree.heading(col, text=col)
            self.compare_tree.column(col, width=120, anchor="center")

        for _, row in df.iterrows():
            self.compare_tree.insert("", tk.END, values=list(row))

def main():
    root = tk.Tk()
    app = MRPGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
