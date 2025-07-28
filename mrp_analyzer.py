import pandas as pd
from datetime import datetime
import os
import logging
from pathlib import Path
from typing import Tuple, Optional, Dict, Any
import numpy as np
from functools import lru_cache

# Configuração de logging mais detalhada
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('mrp_analyzer.log')
    ]
)

class ValidationError(Exception):
    """Exceção personalizada para erros de validação de dados."""
    pass

def validate_numeric_columns(df: pd.DataFrame, columns: list) -> None:
    """
    Valida se as colunas numéricas contêm apenas números válidos.
    
    Args:
        df: DataFrame a ser validado
        columns: Lista de colunas que devem ser numéricas
    
    Raises:
        ValidationError: Se encontrar valores não numéricos
    """
    for col in columns:
        if not pd.to_numeric(df[col], errors='coerce').notna().all():
            invalid_rows = df[pd.to_numeric(df[col], errors='coerce').isna()]
            raise ValidationError(f"Valores não numéricos encontrados na coluna {col}. Linhas: {invalid_rows.index.tolist()}")

def validate_positive_values(df: pd.DataFrame, columns: list) -> None:
    """Valida se as colunas contêm apenas valores positivos."""
    for col in columns:
        if (df[col] < 0).any():
            negative_rows = df[df[col] < 0]
            raise ValidationError(f"Valores negativos encontrados na coluna {col}. Linhas: {negative_rows.index.tolist()}")

@lru_cache(maxsize=32)
def analyze_mrp(input_file: str, sheet_name: str, output_file: str = 'itens_criticos.xlsx') -> Tuple[Optional[int], Optional[str], Optional[pd.DataFrame]]:
    """
    Realiza a análise MRP a partir de um arquivo Excel, salva resultados e histórico, retorna quantidade de itens críticos.
    """
    try:
        logging.info(f"Iniciando análise do arquivo: {input_file}")
        
        # Validação do arquivo
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"Arquivo não encontrado: {input_file}")
        
        # Leitura otimizada do Excel
        df = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            dtype={
                'ESTQ10': 'float64',
                'ESTQ20': 'float64',
                'DEMANDAMRP': 'float64',
                'ESTOQSEG': 'float64',
                'PEDIDOS': 'float64'
            }
        )
        
        # Normalização e validação das colunas
        df.columns = [col.strip().upper() for col in df.columns]
        required_cols = [
            "CÓD", "DESCRIÇÃOPROMOB", "ESTQ10", "ESTQ20", "DEMANDAMRP",
            "ESTOQSEG", "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
        ]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValidationError(f"Colunas obrigatórias ausentes: {', '.join(missing_cols)}")
            
        # Validação de dados
        numeric_cols = ["ESTQ10", "ESTQ20", "DEMANDAMRP", "ESTOQSEG", "PEDIDOS"]
        validate_numeric_columns(df, numeric_cols)
        validate_positive_values(df, numeric_cols)
        
        # Cálculos otimizados com numpy
        df["ESTOQUE DISPONÍVEL"] = np.add(df["ESTQ10"], np.divide(df["ESTQ20"], 3))
        mask = (df["ESTOQUE DISPONÍVEL"] - df["DEMANDAMRP"]) < df["ESTOQSEG"]
        criticos = df[mask].copy()
        
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
    except ValidationError as e:
        logging.error(f"Erro de validação: {str(e)}")
        return None, f"Erro de validação: {str(e)}", None
    except Exception as e:
        logging.error(f"Erro durante a análise: {str(e)}", exc_info=True)
        return None, f"Erro durante a análise: {str(e)}", None

def format_excel(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """
    Formata a planilha Excel com estilos e destaques.
    
    Args:
        writer: Excel writer object
        df: DataFrame a ser formatado
    """
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

def save_history(df: pd.DataFrame, output_file: str) -> None:
    """Salva uma cópia do arquivo no histórico com timestamp."""
    hist_dir = Path(output_file).parent / "historico_mrp"
    hist_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    hist_path = hist_dir / f"itens_criticos_{timestamp}.xlsx"
    salvar_excel_formatado(df, str(hist_path))
    logging.info(f"Histórico salvo em: {hist_path}")

def salvar_excel_formatado(df: pd.DataFrame, output_file: str) -> None:
    """Salva o DataFrame em Excel com formatação."""
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Itens Críticos", index=False)
    format_excel(writer, df)
    writer.close()
    logging.info(f"Arquivo Excel salvo em: {output_file}")

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
