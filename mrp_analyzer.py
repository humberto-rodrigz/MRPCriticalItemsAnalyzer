import pandas as pd
import os
from datetime import datetime

def classify_criticidade(row):
    falta = (row["DEMANDAMRP"] - row["ESTOQUE DISPONÍVEL"])
    if falta >= row["ESTOQSEG"] * 0.5:
        return "Alta"
    elif falta >= row["ESTOQSEG"] * 0.1:
        return "Média"
    else:
        return "Baixa"

def analyze_mrp(input_file, sheet_name, output_folder='output'):
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except Exception as e:
        return None, f"Erro ao ler o Excel: {e}"

    df.columns = df.columns.str.strip().str.replace(' ', '').str.replace('.', '', regex=False).str.upper()

    required = [
        "CÓD", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20", "DEMANDAMRP",
        "ESTOQSEG", "STATUS", "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        return None, f"Colunas faltando: {missing}"

    df = df[df["STATUS"].str.lower() != "inativo"].copy()
    df["ESTOQUE DISPONÍVEL"] = df["ESTOQ10"].fillna(0) + (df["ESTOQ20"].fillna(0) / 3)
    df["QUANTIDADE A SOLICITAR"] = (
        df["DEMANDAMRP"].fillna(0) - df["ESTOQUE DISPONÍVEL"] + df["ESTOQSEG"].fillna(0) - df["PEDIDOS"].fillna(0)
    ).apply(lambda x: max(0, x))

    df["ESTOQUE DISPONÍVEL"] = df["ESTOQUE DISPONÍVEL"].fillna(0).round().astype(int)
    df["QUANTIDADE A SOLICITAR"] = df["QUANTIDADE A SOLICITAR"].fillna(0).round().astype(int)

    criticos = df[(df["ESTOQUE DISPONÍVEL"] - df["DEMANDAMRP"]) < df["ESTOQSEG"]].copy()

    if criticos.empty:
        return 0, "Nenhum item crítico encontrado."

    criticos["FORNECEDOR"] = criticos["FORNECEDORPRINCIPAL"]
    criticos["CRITICIDADE"] = criticos.apply(classify_criticidade, axis=1)
    criticos.sort_values(by="QUANTIDADE A SOLICITAR", ascending=False, inplace=True)

    colunas = {
        "CÓD": "CÓD.",
        "FORNECEDOR": "FORNEC.",
        "DESCRIÇÃOPROMOB": "DESCRIÇÃO",
        "ESTOQ10": "EST. 10",
        "ESTOQ20": "EST. 20",
        "DEMANDAMRP": "DEM. MRP",
        "ESTOQSEG": "EST. SEG.",
        "PEDIDOS": "PED.",
        "ESTOQUE DISPONÍVEL": "EST. DISP.",
        "QUANTIDADE A SOLICITAR": "QTD. A SOLIC.",
        "CRITICIDADE": "CRITICIDADE",
        "OBS": "OBS."
    }

    output_df = criticos[list(colunas.keys())].rename(columns=colunas)

    os.makedirs(output_folder, exist_ok=True)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    excel_path = os.path.join(output_folder, f"itens_criticos_{timestamp}.xlsx")

    # Exportar para Excel com aba de totais
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, sheet_name='Itens Críticos', index=False)
        resumo = output_df.groupby("FORNEC.")["QTD. A SOLIC."].sum().reset_index()
        resumo.to_excel(writer, sheet_name='Resumo por Fornecedor', index=False)

        worksheet = writer.sheets['Itens Críticos']
        workbook = writer.book

        red_format = workbook.add_format({'bg_color': '#F4CCCC', 'border': 1})
        col_index = output_df.columns.get_loc("QTD. A SOLIC.")
        col_letter = chr(65 + col_index)
        worksheet.conditional_format(f"{col_letter}2:{col_letter}{len(output_df)+1}",
                                     {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': red_format})
        worksheet.freeze_panes(1, 0)

        # Gráfico
        chart = workbook.add_chart({'type': 'column'})
        max_rows = min(10, len(output_df))
        chart.add_series({
            'categories': ['Itens Críticos', 1, 2, max_rows, 2],  # DESCRIÇÃO
            'values':     ['Itens Críticos', 1, col_index, max_rows, col_index],
            'name':       'QTD. A SOLIC.',
        })
        chart.set_title({'name': 'Top 10 - Quantidade a Solicitar'})
        chart.set_x_axis({'name': 'Item'})
        chart.set_y_axis({'name': 'Qtd'})
        worksheet.insert_chart('L2', chart)

    return len(output_df), None
