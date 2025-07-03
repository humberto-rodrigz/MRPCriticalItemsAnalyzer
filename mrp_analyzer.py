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

    # Padronizar nomes de colunas para facilitar o acesso
    df.columns = df.columns.str.strip().str.replace(' ', '').str.replace('.', '', regex=False).str.upper()

    # Verificar se as colunas necessárias existem após padronização
    required_columns_upper = ["CÓD", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20", "DEMANDAMRP", "ESTOQSEG", "STATUS", "FORNECEDORPRINCIPAL", "PEDIDOS"]
    missing_columns = [col for col in required_columns_upper if col not in df.columns]
    
    if missing_columns:
        return None, f"Erro: As seguintes colunas não foram encontradas: {missing_columns}. Colunas disponíveis: {list(df.columns)}"

    # Filtrar itens inativos (usando STATUS em maiúsculo)
    df_ativos = df[df["STATUS"].str.lower() != "inativo"].copy()

    # Calcular estoque disponível
    df_ativos["ESTOQUE DISPONÍVEL"] = df_ativos["ESTOQ10"] + (df_ativos["ESTOQ20"] / 3)

    # Identificar itens críticos
    # (Estoque Disponível - Demanda MRP) < Estoq. Seg.
    df_criticos = df_ativos[
        (df_ativos["ESTOQUE DISPONÍVEL"] - df_ativos["DEMANDAMRP"]) < df_ativos["ESTOQSEG"]
    ].copy()

    # Adicionar a coluna FORNECEDOR PRINCIPAL
    df_criticos["FORNECEDOR PRINCIPAL"] = df_criticos["FORNECEDORPRINCIPAL"]

    # Calcular Quantidade a Solicitar
    # Subtrair a coluna PEDIDOS do cálculo
    df_criticos["QUANTIDADE A SOLICITAR"] = df_criticos["DEMANDAMRP"] - df_criticos["ESTOQUE DISPONÍVEL"] + df_criticos["ESTOQSEG"] - df_criticos["PEDIDOS"]
    df_criticos["QUANTIDADE A SOLICITAR"] = df_criticos["QUANTIDADE A SOLICITAR"].apply(lambda x: max(0, x))

    # Arredondar valores para inteiros
    df_criticos["ESTOQUE DISPONÍVEL"] = df_criticos["ESTOQUE DISPONÍVEL"].round().astype(int)
    df_criticos["QUANTIDADE A SOLICITAR"] = df_criticos["QUANTIDADE A SOLICITAR"].round().astype(int)

    # Selecionar colunas relevantes para a saída
    output_columns = ["CÓD", "DESCRIÇÃOPROMOB", "ESTOQ10", "ESTOQ20", "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL", "QUANTIDADE A SOLICITAR", "FORNECEDOR PRINCIPAL"]
    df_criticos = df_criticos[output_columns]

    try:
        # Criar um objeto ExcelWriter para aplicar formatação
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df_criticos.to_excel(writer, sheet_name='Itens Críticos', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Itens Críticos']

        # Formato para cabeçalhos
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # Formato para números inteiros
        int_format = workbook.add_format({'num_format': '0', 'border': 1})

        # Formato para texto
        text_format = workbook.add_format({'border': 1})

        # Escrever os cabeçalhos com formatação
        for col_num, value in enumerate(df_criticos.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Aplicar formatação às colunas
        for col_num, col_name in enumerate(output_columns):
            if col_name in ["ESTOQ10", "ESTOQ20", "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL", "QUANTIDADE A SOLICITAR"]:
                worksheet.set_column(col_num, col_num, 15, int_format) # Largura 15, formato inteiro
            elif col_name == "FORNECEDOR PRINCIPAL":
                worksheet.set_column(col_num, col_num, 30, text_format) # Largura 30 para fornecedor
            else:
                worksheet.set_column(col_num, col_num, 25, text_format) # Largura 25, formato texto

        # Ajustar largura da coluna de descrição
        desc_col_idx = output_columns.index('DESCRIÇÃOPROMOB')
        worksheet.set_column(desc_col_idx, desc_col_idx, 40, text_format)

        writer.close()
        return len(df_criticos), None
    except Exception as e:
        return None, f"Erro ao salvar o arquivo Excel com formatação: {e}"

if __name__ == "__main__":
    print("Para usar, chame a função analyze_mrp(input_file, sheet_name, output_file) com o nome da sua planilha e da aba.")
    print("Ex: analyze_mrp('minha_planilha.xlsx', 'Cálculo MRP')")

