import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time
import webbrowser
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mrp_analyzer import analyze_mrp
from ttkbootstrap import Style

class MRPAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.style = Style("darkly")
        self.style.master = root

        self.root.title("Analisador de Itens Críticos MRP")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)

        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Cálculo MRP")
        self.df_tabela = pd.DataFrame()
        self.current_page = 0
        self.page_size = 50

        self.create_widgets()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.main_frame = ttk.Frame(self.notebook, padding=15)
        self.graph_frame = ttk.Frame(self.notebook, padding=15)
        self.table_frame = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.main_frame, text="Análise MRP")
        self.notebook.add(self.graph_frame, text="Gráfico")
        self.notebook.add(self.table_frame, text="Tabela")

        self.create_main_frame()
        self.create_table_frame()

    def create_main_frame(self):
        form_frame = ttk.Frame(self.main_frame)
        form_frame.pack(pady=20)

        ttk.Label(form_frame, text="Arquivo Excel:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        ttk.Entry(form_frame, textvariable=self.selected_file, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(form_frame, text="Procurar", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(form_frame, text="Nome da Aba:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)
        ttk.Entry(form_frame, textvariable=self.sheet_name, width=30).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Button(form_frame, text="Analisar MRP", command=self.analyze_mrp_file).grid(row=2, column=0, columnspan=3, pady=10)

        self.progress = ttk.Progressbar(form_frame, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.status_label = ttk.Label(form_frame, text="")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=(0, 10))

        ttk.Label(self.main_frame, text="Log:").pack(anchor=tk.W, padx=10)
        log_frame = ttk.Frame(self.main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.root.bind("<Control-t>", self.toggle_theme)
        self.log_message("Bem-vindo ao Analisador de Itens Críticos MRP!")

    def create_table_frame(self):
        top_toolbar = ttk.Frame(self.table_frame)
        top_toolbar.pack(fill=tk.X, pady=5)

        self.filter_column = tk.StringVar()
        self.filter_entry = tk.StringVar()
        self.filter_box = ttk.Combobox(top_toolbar, textvariable=self.filter_column, state="readonly", width=25)
        self.filter_box.pack(side=tk.LEFT, padx=5)

        ttk.Entry(top_toolbar, textvariable=self.filter_entry, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_toolbar, text="Filtrar", command=self.apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_toolbar, text="Recarregar", command=self.load_table_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_toolbar, text="Salvar Como Excel", command=self.salvar_como_excel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_toolbar, text="Exportar CSV", command=self.exportar_csv).pack(side=tk.RIGHT, padx=5)

        self.tree = ttk.Treeview(self.table_frame, show="headings")
        self.tree.bind("<Double-1>", self.show_details_popup)
        self.tree.pack(fill=tk.BOTH, expand=True)

        bottom = ttk.Frame(self.table_frame)
        bottom.pack(fill=tk.X, pady=10)

        self.stats_label = ttk.Label(bottom, text="")
        self.stats_label.pack(side=tk.LEFT, padx=10)

        self.page_buttons = ttk.Frame(bottom)
        self.page_buttons.pack(side=tk.RIGHT)

        ttk.Button(self.page_buttons, text="Anterior", command=self.previous_page).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.page_buttons, text="Próximo", command=self.next_page).pack(side=tk.LEFT, padx=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if filename:
            self.selected_file.set(filename)
            self.log_message(f"Arquivo selecionado: {os.path.basename(filename)}")

    def log_message(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)

    def validate_sheet(self, file, sheet):
        try:
            return sheet in pd.ExcelFile(file).sheet_names
        except Exception as e:
            self.log_message(f"Erro ao validar aba: {e}")
            return False

    def analyze_mrp_file(self):
        file = self.selected_file.get()
        sheet = self.sheet_name.get()
        if not file or not os.path.exists(file):
            messagebox.showerror("Erro", "Arquivo não encontrado.")
            return
        if not self.validate_sheet(file, sheet):
            messagebox.showerror("Erro", f"A aba '{sheet}' não existe no arquivo.")
            return

        self.progress.start()
        self.status_label.config(text="Analisando...")
        self.root.update_idletasks()

        start = time.time()
        output_file = os.path.join(os.path.dirname(file), "itens_criticos.xlsx")
        num_items, error, _ = analyze_mrp(file, sheet, output_file)
        self.progress.stop()

        if error:
            self.log_message(f"Erro: {error}")
            messagebox.showerror("Erro", error)
        else:
            tempo = round(time.time() - start, 2)
            self.log_message(f"Concluído em {tempo}s. {num_items} itens críticos.")
            self.status_label.config(text=f"Concluído em {tempo}s")
            self.plot_graph(output_file)
            self.load_table_data(output_file)
            if messagebox.askyesno("Sucesso", "Deseja abrir o arquivo gerado?"):
                webbrowser.open(output_file)

    def plot_graph(self, excel_path):
        try:
            df = pd.read_excel(excel_path)
            df = df[df["QUANTIDADE A SOLICITAR"] > 0]
            fig, ax = plt.subplots(figsize=(9, 6))
            bars = ax.barh(df["CÓD"].astype(str), df["QUANTIDADE A SOLICITAR"])
            ax.set_xlabel("Qtd a Solicitar")
            ax.set_title("Itens Críticos - Quantidade a Solicitar")
            for bar in bars:
                width = bar.get_width()
                ax.text(width + 1, bar.get_y() + bar.get_height()/2, str(int(width)), va='center')
            fig.tight_layout()

            for widget in self.graph_frame.winfo_children():
                widget.destroy()
            canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            self.log_message(f"Erro ao gerar gráfico: {e}")

    def load_table_data(self, excel_path=None):
        try:
            file = self.selected_file.get()
            excel_path = excel_path or os.path.join(os.path.dirname(file), "itens_criticos.xlsx")
            self.df_tabela = pd.read_excel(excel_path)
            self.filter_box['values'] = list(self.df_tabela.columns)
            self.current_page = 0
            self.show_table_page()
        except Exception as e:
            self.log_message(f"Erro ao carregar tabela: {e}")

    def show_table_page(self):
        self.tree.delete(*self.tree.get_children())
        df = self.df_tabela.copy()
        start = self.current_page * self.page_size
        end = start + self.page_size
        page = df.iloc[start:end]

        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=100, anchor="center")

        for _, row in page.iterrows():
            self.tree.insert("", tk.END, values=list(row))

        total = len(df)
        soma = df["QUANTIDADE A SOLICITAR"].sum() if "QUANTIDADE A SOLICITAR" in df.columns else 0
        top_forn = df["FORNECEDOR PRINCIPAL"].value_counts().idxmax() if "FORNECEDOR PRINCIPAL" in df.columns else "-"
        self.stats_label.config(text=f"Total: {total} | Soma: {soma} | Fornecedor Top: {top_forn}")

    def next_page(self):
        if (self.current_page + 1) * self.page_size < len(self.df_tabela):
            self.current_page += 1
            self.show_table_page()

    def previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.show_table_page()

    def apply_filter(self):
        col = self.filter_column.get()
        val = self.filter_entry.get().strip().lower()
        if col and val:
            self.df_tabela = self.df_tabela[self.df_tabela[col].astype(str).str.lower().str.contains(val)]
            self.current_page = 0
            self.show_table_page()

    def sort_column(self, col):
        self.df_tabela.sort_values(by=col, ascending=True, inplace=True, ignore_index=True)
        self.current_page = 0
        self.show_table_page()

    def show_details_popup(self, event):
        item = self.tree.selection()
        if item:
            values = self.tree.item(item[0], 'values')
            colnames = self.df_tabela.columns.tolist()
            msg = "\n".join([f"{k}: {v}" for k, v in zip(colnames, values)])
            messagebox.showinfo("Detalhes do Item", msg)

    def exportar_csv(self):
        file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file:
            self.df_tabela.to_csv(file, index=False)
            self.log_message(f"CSV exportado para: {file}")

    def salvar_como_excel(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.df_tabela.to_excel(file, index=False)
            self.log_message(f"Excel salvo em: {file}")

    def toggle_theme(self, event=None):
        atual = self.style.theme.name
        novo = "darkly" if atual == "flatly" else "flatly"
        self.style.theme_use(novo)
        self.log_message(f"Tema alterado para: {novo}")

def main():
    root = tk.Tk()
    app = MRPAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
