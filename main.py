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


def format_excel(writer, df):   

    ()  

    directory = os.path.dirname(writer.path)
    if not os.path.exists(directory):   
        os.makedirs(directory)  

def salvar_excel_formatado(df, output_file):
    """Salva o DataFrame em Excel formatado."""
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Itens Críticos", index=False)
    format_excel(writer, df)
    writer.close()
    """Formata o Excel com cores e bordas."""
    workbook = writer.book  
    worksheet = writer.sheets["Itens Críticos"]
    cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9EAD3'})

