import pandas as pd
import ibm_db
from datetime import datetime
import os
import logging
import sys

# Configuração básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def connect_db():
    """Estabelece conexão com o banco de dados DB2 usando variáveis de ambiente."""
    conn_str = (
        f"DATABASE=ENGATCAR;"
        f"HOSTNAME={os.getenv('DB2_HOST', 'SEU_SERVIDOR')};"
        f"PORT=50000;"
        f"PROTOCOL=TCPIP;"
        f"UID={os.getenv('DB2_USER', 'seu_usuario')};"
        f"PWD={os.getenv('DB2_PASS', 'sua_senha')};"
    )
    try:
        conn = ibm_db.connect(conn_str, "", "")
        logging.info("Conexão com o banco de dados estabelecida com sucesso.")
        return conn
    except Exception as e:
        logging.error(f"Erro ao conectar ao banco de dados: {e}")
        return None

def fetch_mrp_data(conn):
    """Busca os dados MRP do banco e retorna um DataFrame pandas."""
    sql = """
    SELECT
        CODIGO AS "CÓD",
        DESCRICAO AS "DESCRIÇÃOPROMOB",
        ESTQ10,
        ESTQ20,
        DEMANDA AS "DEMANDAMRP",
        ESTQSEG AS "ESTOQSEG",
        STATUS,
        FORNECEDOR AS "FORNECEDORPRINCIPAL",
        PEDIDOS,
        OBSERVACAO AS "OBS"
    FROM TABELA_MRP
    WHERE STATUS <> 'Inativo'
    """
    stmt = ibm_db.exec_immediate(conn, sql)
    col_count = ibm_db.num_fields(stmt)
    col_names = [ibm_db.field_name(stmt, i) for i in range(col_count)]
    rows = []
    row = ibm_db.fetch_assoc(stmt)
    while row:
        rows.append(row)
        row = ibm_db.fetch_assoc(stmt)
    df = pd.DataFrame.from_records(rows, columns=col_names)
    return df

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
    """Salva o DataFrame em Excel formatado."""
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Itens Críticos", index=False)
    format_excel(writer, df)
    writer.close()

def analyze_mrp_from_db(output_file='itens_criticos.xlsx'):
    """Executa a análise MRP, salva resultados e histórico, retorna quantidade de itens críticos."""
    conn = connect_db()
    if not conn:
        return None, "Erro ao conectar ao banco", None

    try:
        df = fetch_mrp_data(conn)
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

        # Histórico
        hist_dir = os.path.join(os.path.dirname(output_file), "historico_mrp")
        os.makedirs(hist_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        hist_path = os.path.join(hist_dir, f"itens_criticos_{timestamp}.xlsx")
        salvar_excel_formatado(criticos, hist_path)

        return len(criticos), None, criticos

    except Exception as e:
        logging.error(f"Erro na análise: {e}")
        return None, f"Erro na análise: {e}", None
    finally:
        try:
            if conn:
                ibm_db.close(conn)
                logging.info("Conexão com o banco de dados fechada.")
        except Exception as e:
            logging.warning(f"Erro ao fechar conexão: {e}")

if __name__ == "__main__":
    count, err, df = analyze_mrp_from_db()
    if err:
        print("Erro:", err)
    else:
        print(f"{count} itens críticos identificados.")
