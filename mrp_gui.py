import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import pandas as pd
from mrp_analyzer import analyze_mrp

class MRPAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador de Itens Críticos MRP")
        self.root.geometry("800x550")
        self.root.resizable(True, True)

        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.sheet_options = []
        self.excel_path = None
        self.pdf_path = None

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Analisador de Itens Críticos MRP", font=("Arial", 16, "bold"))\
            .grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Seleção de arquivo
        ttk.Label(main_frame, text="Arquivo Excel:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.selected_file, width=60)\
            .grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Procurar", command=self.browse_file).grid(row=1, column=2)

        # Aba do Excel
        ttk.Label(main_frame, text="Selecionar Aba:").grid(row=2, column=0, sticky=tk.W, pady=(10, 5))
        self.sheet_dropdown = ttk.Combobox(main_frame, textvariable=self.sheet_name, values=self.sheet_options, state="readonly")
        self.sheet_dropdown.grid(row=2, column=1, sticky=tk.W, padx=5, pady=(10, 5))

        # Botão Analisar
        ttk.Button(main_frame, text="Analisar MRP", command=self.analyze_mrp_file)\
            .grid(row=3, column=0, columnspan=3, pady=20)

        # Botões abrir arquivos
        self.open_excel_btn = ttk.Button(main_frame, text="Abrir Excel Gerado", command=self.open_excel, state=tk.DISABLED)
        self.open_excel_btn.grid(row=4, column=0, columnspan=1)

        self.open_pdf_btn = ttk.Button(main_frame, text="Abrir PDF Gerado", command=self.open_pdf, state=tk.DISABLED)
        self.open_pdf_btn.grid(row=4, column=1, columnspan=2, sticky=tk.W)

        # Log
        ttk.Label(main_frame, text="Log de Execução:").grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.N, tk.S, tk.E, tk.W))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        main_frame.rowconfigure(6, weight=1)

        self.log_message("Bem-vindo ao Analisador de Itens Críticos MRP!")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Selecionar Planilha MRP",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.selected_file.set(filename)
            self.log_message(f"Arquivo selecionado: {os.path.basename(filename)}")
            self.load_sheet_names(filename)

    def load_sheet_names(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            self.sheet_options = xls.sheet_names
            self.sheet_dropdown['values'] = self.sheet_options
            if self.sheet_options:
                self.sheet_name.set(self.sheet_options[0])
                self.log_message("Abas disponíveis: " + ", ".join(self.sheet_options))
        except Exception as e:
            self.log_message(f"Erro ao ler abas: {e}")
            messagebox.showerror("Erro", str(e))

    def analyze_mrp_file(self):
        if not self.selected_file.get():
            messagebox.showerror("Erro", "Selecione um arquivo Excel.")
            return
        if not self.sheet_name.get():
            messagebox.showerror("Erro", "Selecione uma aba da planilha.")
            return

        self.log_message("\n⏳ Iniciando análise...")
        try:
            result, error = analyze_mrp(self.selected_file.get(), self.sheet_name.get())

            if error:
                self.log_message(f"❌ Erro: {error}")
                messagebox.showerror("Erro na análise", error)
                self.open_excel_btn.config(state=tk.DISABLED)
                self.open_pdf_btn.config(state=tk.DISABLED)
            else:
                self.log_message(f"✅ {result} itens críticos encontrados.")
                output_dir = os.path.join(os.getcwd(), "output")
                files = sorted(os.listdir(output_dir), reverse=True)
                excel_file = next((f for f in files if f.endswith(".xlsx")), None)
                pdf_file = next((f for f in files if f.endswith(".pdf")), None)

                if excel_file:
                    self.excel_path = os.path.join(output_dir, excel_file)
                    self.log_message(f"📄 Excel gerado: {excel_file}")
                    self.open_excel_btn.config(state=tk.NORMAL)
                if pdf_file:
                    self.pdf_path = os.path.join(output_dir, pdf_file)
                    self.log_message(f"📝 PDF gerado: {pdf_file}")
                    self.open_pdf_btn.config(state=tk.NORMAL)

                messagebox.showinfo("Sucesso", f"Análise finalizada com {result} itens críticos.")

        except Exception as e:
            self.log_message(f"❌ Erro inesperado: {str(e)}")
            messagebox.showerror("Erro inesperado", str(e))

    def log_message(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def open_excel(self):
        if self.excel_path and os.path.exists(self.excel_path):
            os.startfile(self.excel_path)

    def open_pdf(self):
        if self.pdf_path and os.path.exists(self.pdf_path):
            os.startfile(self.pdf_path)

def main():
    root = tk.Tk()
    app = MRPAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
