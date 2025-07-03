# === mrp_gui.py atualizado com gráfico, validação, barra de progresso e tema escuro ===
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
        self.style = Style("darkly")  # ou "cosmo", "flatly", "darkly", etc.
        self.style.master = root

        self.root.title("Analisador de Itens Críticos MRP")
        self.root.geometry("900x600")
        self.root.resizable(True, True)

        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Cálculo MRP")

        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.main_frame = ttk.Frame(notebook, padding=10)
        self.graph_frame = ttk.Frame(notebook, padding=10)

        notebook.add(self.main_frame, text="Análise MRP")
        notebook.add(self.graph_frame, text="Gráfico")

        ttk.Label(self.main_frame, text="Arquivo Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.main_frame, textvariable=self.selected_file, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(self.main_frame, text="Procurar", command=self.browse_file).grid(row=0, column=2, padx=5)

        ttk.Label(self.main_frame, text="Nome da Aba:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.main_frame, textvariable=self.sheet_name, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)

        ttk.Button(self.main_frame, text="Analisar MRP", command=self.analyze_mrp_file).grid(row=2, column=0, columnspan=3, pady=10)

        self.progress = ttk.Progressbar(self.main_frame, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.status_label = ttk.Label(self.main_frame, text="")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=(0, 10))

        ttk.Label(self.main_frame, text="Log de Execução:").grid(row=5, column=0, sticky=tk.W, pady=(10, 5))
        log_frame = ttk.Frame(self.main_frame)
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.root.bind("<Control-t>", self.toggle_theme)

        self.log_message("Bem-vindo ao Analisador de Itens Críticos MRP!")
        self.log_message("Selecione um arquivo Excel e clique em 'Analisar MRP'.")

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
            sheets = pd.ExcelFile(file).sheet_names
            return sheet in sheets
        except Exception as e:
            self.log_message(f"Erro ao ler abas: {e}")
            return False

    def analyze_mrp_file(self):
        file = self.selected_file.get()
        sheet = self.sheet_name.get()

        if not file or not os.path.exists(file):
            messagebox.showerror("Erro", "Arquivo inválido.")
            return

        if not self.validate_sheet(file, sheet):
            messagebox.showerror("Erro", f"A aba '{sheet}' não foi encontrada no arquivo.")
            return

        self.progress.start()
        self.status_label.config(text="Analisando...")
        self.root.update_idletasks()

        start = time.time()
        output_file = os.path.join(os.path.dirname(file), "itens_criticos.xlsx")
        num_items, error = analyze_mrp(file, sheet, output_file)
        tempo = round(time.time() - start, 2)
        self.progress.stop()

        if error:
            self.log_message(f"Erro: {error}")
            messagebox.showerror("Erro", error)
        else:
            self.log_message(f"Concluído em {tempo}s. {num_items} itens críticos.")
            self.status_label.config(text=f"Concluído em {tempo}s")
            self.plot_graph(output_file)
            abrir = messagebox.askyesno("Sucesso", "Deseja abrir o arquivo gerado?")
            if abrir:
                webbrowser.open(output_file)

    def plot_graph(self, excel_path):
        try:
            df = pd.read_excel(excel_path)
            df = df[df["QUANTIDADE A SOLICITAR"] > 0]
            fig, ax = plt.subplots(figsize=(8, 5))
            ax.barh(df["CÓD"].astype(str), df["QUANTIDADE A SOLICITAR"])
            ax.set_xlabel("Qtd a Solicitar")
            ax.set_title("Itens Críticos - Quantidade a Solicitar")
            fig.tight_layout()

            for widget in self.graph_frame.winfo_children():
                widget.destroy()

            canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            self.log_message(f"Erro ao gerar gráfico: {e}")

    def toggle_theme(self, event=None):
        theme = "darkly" if self.style.theme.name == "flatly" else "flatly"
        self.style.theme_use(theme)
        self.log_message(f"Tema alterado para: {theme}")

def main():
    root = tk.Tk()
    app = MRPAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
