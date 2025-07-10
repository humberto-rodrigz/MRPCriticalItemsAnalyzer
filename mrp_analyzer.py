import pandas as pd
import ibm_db
from datetime import datetime
import os

def connect_db():
    conn_str = (
        "DATABASE=ENGATCAR;"
        "HOSTNAME=SEU_SERVIDOR;"  # ex: 192.168.0.10
        "PORT=50000;"
        "PROTOCOL=TCPIP;"
        "UID=seu_usuario;"
        "PWD=sua_senha;"
    )
    try:
        conn = ibm_db.connect(conn_str, "", "")
        return conn
    except Exception as e:
        print("Erro ao conectar:", e)
        return None

def fetch_mrp_data(conn):
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
    rows = []
    col_count = ibm_db.num_fields(stmt)
    col_names = [ibm_db.field_name(stmt, i) for i in range(col_count)]
    row = ibm_db.fetch_assoc(stmt)
    while row:
        rows.append(row)
        row = ibm_db.fetch_assoc(stmt)
    return pd.DataFrame(rows, columns=col_names)

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

    for i, col in enumerate(df.columns):
        fmt = int_fmt if pd.api.types.is_numeric_dtype(df[col]) else text_fmt
        worksheet.set_column(i, i, 20, fmt)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            fmt = highlight_fmt if df.columns[col_idx] == "QUANTIDADE A SOLICITAR" and isinstance(value, (int, float)) and value > 0 else alt_row_fmt if row_idx % 2 == 0 else None
            worksheet.write(row_idx, col_idx, value, fmt)

def analyze_mrp_from_db(output_file='itens_criticos.xlsx'):
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

        ibm_db.close(conn)
        return len(criticos), None, criticos

    except Exception as e:
        ibm_db.close(conn)
        return None, f"Erro na análise: {e}", None

if __name__ == "__main__":
    count, err, df = analyze_mrp_from_db()
    if err:
        print("Erro:", err)
    else:
        print(f"{count} itens críticos identificados.")
