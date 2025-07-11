import calendar
import pandas as pd
import ibm_db
from datetime import datetime   
import os
import logging      
import xlsxwriter
from mrp_analyzer import connect_db, fetch_mrp_data, format_excel

def salvar_excel_formatado(df, output_file):
    """Salva o DataFrame em Excel formatado."""
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Itens Críticos", index=False)
    format_excel(writer, df)
    writer.close()  


