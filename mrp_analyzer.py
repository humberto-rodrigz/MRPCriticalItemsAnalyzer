import pandas as pd

def analyze_mrp(input_file, sheet_name, output_file='itens_criticos.xlsx'):
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except FileNotFoundError:
        return None, f"Erro: O arquivo '{input_file}' não foi encontrado."
    except KeyError:
        return None, f"Erro: A aba '{sheet_name}' não foi encontrada no arquivo Excel."
    except Exception as e:
        return None, f"Erro ao ler o arquivo Excel: {e}"

    df.columns = df.columns.str.strip().str.replace(' ', '').str.replace('.', '', regex=False).str.upper()
    
    required_columns_upper = [
        "CÓD", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20",
        "DEMANDAMRP", "ESTOQSEG", "STATUS",
        "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
    ]
    missing_columns = [col for col in required_columns_upper if col not in df.columns]
    if missing_columns:
        return None, f"Erro: As seguintes colunas não foram encontradas: {missing_columns}."

    df_ativos = df[df["STATUS"].str.lower() != "inativo"].copy()
    df_ativos["ESTOQUE DISPONÍVEL"] = df_ativos["ESTOQ10"] + (df_ativos["ESTOQ20"] / 3)

    df_criticos = df_ativos[
        (df_ativos["ESTOQUE DISPONÍVEL"] - df_ativos["DEMANDAMRP"]) < df_ativos["ESTOQSEG"]
    ].copy()

    df_criticos["QUANTIDADE A SOLICITAR"] = (
        df_criticos["DEMANDAMRP"] - df_criticos["ESTOQUE DISPONÍVEL"] + df_criticos["ESTOQSEG"] - df_criticos["PEDIDOS"]
    )
    df_criticos["QUANTIDADE A SOLICITAR"] = df_criticos["QUANTIDADE A SOLICITAR"].apply(lambda x: max(0, x))

    df_criticos["ESTOQUE DISPONÍVEL"] = df_criticos["ESTOQUE DISPONÍVEL"].round().astype(int)
    df_criticos["QUANTIDADE A SOLICITAR"] = df_criticos["QUANTIDADE A SOLICITAR"].round().astype(int)
    df_criticos["FORNECEDOR PRINCIPAL"] = df_criticos["FORNECEDORPRINCIPAL"]

    output_columns = [
        "CÓD", "FORNECEDOR PRINCIPAL", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20",
        "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL", "QUANTIDADE A SOLICITAR", "OBS"
    ]
    df_criticos = df_criticos[output_columns]
    df_criticos.sort_values(by="QUANTIDADE A SOLICITAR", ascending=False, inplace=True)

    df_criticos.replace([float('inf'), float('-inf')], pd.NA, inplace=True)
    df_criticos.fillna('', inplace=True)

    try:
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df_criticos.to_excel(writer, sheet_name='Itens Críticos', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Itens Críticos']

        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })
        int_format = workbook.add_format({'num_format': '0', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        red_fill = workbook.add_format({'bg_color': '#F4CCCC', 'border': 1})
        alt_gray = workbook.add_format({'bg_color': '#F9F9F9'})

        for col_num, value in enumerate(df_criticos.columns.values):
            worksheet.write(0, col_num, value, header_format)

        col_widths = {
            "FORNECEDOR PRINCIPAL": 20,
            "DESCRIÇÃOPROMOB": 30,
            "ESTOQUE DISPONÍVEL": 18,
            "QUANTIDADE A SOLICITAR": 20
        }

        for col_num, col_name in enumerate(output_columns):
            width = col_widths.get(col_name, max(10, len(col_name) + 2))
            fmt = int_format if df_criticos[col_name].dtype.kind in 'iufc' else text_format
            worksheet.set_column(col_num, col_num, width, fmt)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df_criticos), len(output_columns)-1)

        for row_num, row in enumerate(df_criticos.itertuples(index=False), start=1):
            for col_num, value in enumerate(row):
                col_is_qtd = col_num == output_columns.index("QUANTIDADE A SOLICITAR")
                format_ = red_fill if col_is_qtd and str(value).isdigit() and int(value) != 0 else None
                if not format_:
                    format_ = alt_gray if row_num % 2 == 0 else None
                worksheet.write(row_num, col_num, value, format_)

        writer.close()
        return len(df_criticos), None
    except Exception as e:
        return None, f"Erro ao salvar o arquivo Excel com formatação: {e}"

if __name__ == "__main__":
    print("Para usar, chame a função analyze_mrp(input_file, sheet_name, output_file)")
