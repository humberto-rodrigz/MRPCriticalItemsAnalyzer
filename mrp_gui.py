import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from mrp_analyzer import analyze_mrp

class MRPAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador de Itens Críticos MRP")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Variáveis
        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Cálculo MRP")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Analisador de Itens Críticos MRP", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Seleção de arquivo
        ttk.Label(main_frame, text="Arquivo Excel:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        file_entry = ttk.Entry(main_frame, textvariable=self.selected_file, width=50)
        file_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        
        browse_button = ttk.Button(main_frame, text="Procurar", command=self.browse_file)
        browse_button.grid(row=1, column=2, pady=5)
        
        # Nome da aba
        ttk.Label(main_frame, text="Nome da Aba:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        sheet_entry = ttk.Entry(main_frame, textvariable=self.sheet_name, width=30)
        sheet_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Botão de análise
        analyze_button = ttk.Button(main_frame, text="Analisar MRP", 
                                   command=self.analyze_mrp_file, style="Accent.TButton")
        analyze_button.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Área de log
        ttk.Label(main_frame, text="Log de Execução:").grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        
        # Frame para o log com scrollbar
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=10, width=70, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar expansão do grid
        main_frame.rowconfigure(5, weight=1)
        
        # Adicionar mensagem inicial
        self.log_message("Bem-vindo ao Analisador de Itens Críticos MRP!")
        self.log_message("Selecione um arquivo Excel e clique em \'Analisar MRP\' para começar.")
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Selecionar Planilha MRP",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.selected_file.set(filename)
            self.log_message(f"Arquivo selecionado: {os.path.basename(filename)}")
    
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def analyze_mrp_file(self):
        if not self.selected_file.get():
            messagebox.showerror("Erro", "Por favor, selecione um arquivo Excel.")
            return
        
        if not self.sheet_name.get():
            messagebox.showerror("Erro", "Por favor, informe o nome da aba.")
            return
        
        try:
            self.log_message("Iniciando análise...")
            
            input_file = self.selected_file.get()
            output_dir = os.path.dirname(input_file)
            output_file = os.path.join(output_dir, "itens_criticos.xlsx")
            
            # Garantir que o diretório de saída exista
            os.makedirs(output_dir, exist_ok=True)

            # Executar análise e capturar o resultado e a mensagem de erro
            num_critical_items, error_message = analyze_mrp(input_file, self.sheet_name.get(), output_file)
            
            if error_message:
                self.log_message(f"Análise falhou: {error_message}")
                messagebox.showerror("Erro", f"Análise falhou.\n\nDetalhes: {error_message}")
            elif num_critical_items is not None:
                self.log_message(f"Análise concluída com sucesso! {num_critical_items} itens críticos encontrados.")
                self.log_message(f"Arquivo de saída: {output_file}")
                messagebox.showinfo("Sucesso", 
                                  f"Análise concluída!\n\n{num_critical_items} itens críticos encontrados.\nArquivo de saída salvo em:\n{output_file}")
            else:
                self.log_message("Análise falhou. Motivo desconhecido. Verifique o log para mais detalhes.")
                messagebox.showerror("Erro", "Análise falhou. Motivo desconhecido. Verifique o log para mais detalhes.")
            
        except Exception as e:
            error_msg = f"Erro inesperado durante a análise: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Erro", error_msg)

def main():
    root = tk.Tk()
    app = MRPAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

