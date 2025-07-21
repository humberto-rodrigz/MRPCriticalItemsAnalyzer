import pandas as pd
from datetime import datetime
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def analyze_mrp_from_excel(input_file, sheet_name, output_file='itens_criticos.xlsx'):
    """
    Realiza a análise MRP a partir de um arquivo Excel, salva resultados e histórico, retorna quantidade de itens críticos.
    """
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        # Normaliza nomes de colunas para evitar erros de digitação
        df.columns = [col.strip().upper() for col in df.columns]
        required_cols = [
            "CÓD", "DESCRIÇÃOPROMOB", "ESTQ10", "ESTQ20", "DEMANDAMRP", "ESTOQSEG", "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
        ]
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"Coluna obrigatória ausente: {col}")
        df["ESTOQUE DISPONÍVEL"] = df["ESTQ10"] + (df["ESTQ20"] / 3)
        criticos = df[(df["ESTOQUE DISPONÍVEL"] - df["DEMANDAMRP"]) < df["ESTOQSEG"]].copy()
        criticos["QUANTIDADE A SOLICITAR"] = (
            criticos["DEMANDAMRP"] - criticos["ESTOQUE DISPONÍVEL"] + criticos["ESTOQSEG"] - criticos["PEDIDOS"]
        ).clip(lower=0).round().astype(int)
        criticos["FORNECEDOR PRINCIPAL"] = criticos["FORNECEDORPRINCIPAL"]
        criticos["ESTOQUE DISPONÍVEL"] = criticos["ESTOQUE DISPONÍVEL"].round().astype(int)
        final_columns = [
            "CÓD", "FORNECEDOR PRINCIPAL", "DESCRIÇÃOPROMOB", "ESTQ10", "ESTQ20",
            "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL",
            "QUANTIDADE A SOLICITAR", "OBS"
        ]
        criticos = criticos[final_columns].fillna("")
        salvar_excel_formatado(criticos, output_file)
        hist_dir = os.path.join(os.path.dirname(output_file), "historico_mrp")
        os.makedirs(hist_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        hist_path = os.path.join(hist_dir, f"itens_criticos_{timestamp}.xlsx")
        salvar_excel_formatado(criticos, hist_path)
        return len(criticos), None, criticos
    except Exception as e:
        logging.error(f"Erro na análise: {e}")
        return None, f"Erro na análise: {e}", None

def format_excel(writer, df):
    """Formata a planilha Excel gerada com estilos e destaques."""
    workbook = writer.book
    worksheet = writer.sheets['Itens Críticos']

    header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                                      'fg_color': '#D7E4BC', 'border': 1})
    int_fmt = workbook.add_format({'num_format': '0', 'border': 1})
    text_fmt = workbook.add_format({'border': 1})
    highlight_fmt = workbook.add_format({'bg_color': '#F4CCCC', 'border': 1})
    alt_row_fmt = workbook.add_format({'bg_color': '#F9F9F9'})

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    for i, col in enumerate(df.columns):
        fmt = int_fmt if pd.api.types.is_numeric_dtype(df[col]) else text_fmt
        worksheet.set_column(i, i, 20, fmt)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            fmt = highlight_fmt if df.columns[col_idx] == "QUANTIDADE A SOLICITAR" and isinstance(value, (int, float)) and value > 0 else alt_row_fmt if row_idx % 2 == 0 else None
            worksheet.write(row_idx, col_idx, value, fmt)

def salvar_excel_formatado(df, output_file):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Itens Críticos", index=False)
    format_excel(writer, df)
    writer.close()

if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3:
        input_file = sys.argv[1]
        sheet_name = sys.argv[2]
        count, err, df = analyze_mrp_from_excel(input_file, sheet_name)
        if err:
            print("Erro:", err)
        else:
            print(f"{count} itens críticos identificados.")
    else:
        print("Uso: python mrp_analyzer.py <arquivo.xlsx> <nome_da_aba>")
