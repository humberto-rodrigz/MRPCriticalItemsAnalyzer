import pandas as pd
import os
from datetime import datetime

def format_excel(writer, df):
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

    col_widths = {
        "FORNECEDOR PRINCIPAL": 20,
        "DESCRIÇÃOPROMOB": 30,
        "ESTOQUE DISPONÍVEL": 18,
        "QUANTIDADE A SOLICITAR": 20
    }

    for i, col in enumerate(df.columns):
        width = col_widths.get(col, max(10, len(col) + 2))
        fmt = int_fmt if pd.api.types.is_numeric_dtype(df[col]) else text_fmt
        worksheet.set_column(i, i, width, fmt)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            fmt = None
            if df.columns[col_idx] == "QUANTIDADE A SOLICITAR" and isinstance(value, (int, float)) and value > 0:
                fmt = highlight_fmt
            elif row_idx % 2 == 0:
                fmt = alt_row_fmt
            worksheet.write(row_idx, col_idx, value, fmt)

def analyze_mrp(input_file, sheet_name, output_file='itens_criticos.xlsx'):
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except Exception as e:
        return None, f"Error ao ler Excel: {e}", None

    df.columns = df.columns.str.strip().str.upper().str.replace(" ", "").str.replace(".", "", regex=False)

    required = [
        "CÓD", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20",
        "DEMANDAMRP", "ESTOQSEG", "STATUS",
        "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        return None, f"Colunas ausentes: {missing}", None

    df = df[df["STATUS"].str.lower() != "inativo"].copy()
    df["ESTOQUE DISPONÍVEL"] = df["ESTOQ10"] + (df["ESTOQ20"] / 3)

    criticos = df[(df["ESTOQUE DISPONÍVEL"] - df["DEMANDAMRP"]) < df["ESTOQSEG"]].copy()
    criticos["QUANTIDADE A SOLICITAR"] = (
        criticos["DEMANDAMRP"] - criticos["ESTOQUE DISPONÍVEL"] + criticos["ESTOQSEG"] - criticos["PEDIDOS"]
    ).clip(lower=0).round().astype(int)

    criticos["FORNECEDOR PRINCIPAL"] = criticos["FORNECEDORPRINCIPAL"]
    criticos["ESTOQUE DISPONÍVEL"] = criticos["ESTOQUE DISPONÍVEL"].round().astype(int)

    final_columns = [
        "CÓD", "FORNECEDOR PRINCIPAL", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20",
        "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL",
        "QUANTIDADE A SOLICITAR", "OBS"
    ]
    criticos = criticos[final_columns].fillna("").replace([float("inf"), float("-inf")], pd.NA)
    criticos.sort_values(by="QUANTIDADE A SOLICITAR", ascending=False, inplace=True)

    try:
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        criticos.to_excel(writer, sheet_name="Itens Críticos", index=False)
        format_excel(writer, criticos)
        writer.close()

        hist_dir = os.path.join(os.path.dirname(output_file), "historico_mrp")
        os.makedirs(hist_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        hist_path = os.path.join(hist_dir, f"itens_criticos_{timestamp}.xlsx")

        hist_writer = pd.ExcelWriter(hist_path, engine='xlsxwriter')
        criticos.to_excel(hist_writer, sheet_name="Itens Críticos", index=False)
        format_excel(hist_writer, criticos)
        hist_writer.close()

        return len(criticos), None, criticos
    except Exception as e:
        return None, f"Error ao salvar Excel: {e}", None

if __name__ == "__main__":
    print("Use: analyze_mrp(input_file, sheet_name)")
