import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

st.title("Generador Formulario Renta Sociedades")

uploaded_file = st.file_uploader("Sube el FORMULARIO RENTA SOCIEDADES-1", type=["xls", "xlsx"])

if uploaded_file is not None:
    
    # Leer el archivo
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # Hoja 1: Formulario original
        df.to_excel(writer, sheet_name="FORMULARIO RENTA SOCIEDADES-1", index=False)
        
        workbook = writer.book
        worksheet1 = writer.sheets["FORMULARIO RENTA SOCIEDADES-1"]
        
        # Crear hoja F101 nueva
        worksheet2 = workbook.add_worksheet("F101_GENERADO")
        
        # Encabezados ejemplo
        headers = ["Código", "Valor Formulario"]
        
        for col_num, header in enumerate(headers):
            worksheet2.write(0, col_num, header)
        
        # Formato encabezado
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        for col_num in range(len(headers)):
            worksheet2.write(0, col_num, headers[col_num], header_format)
        
        # Ejemplo de códigos (puedes adaptar a tus casilleros reales)
        codigos = [101, 102, 103, 104, 105]
        
        for row_num, codigo in enumerate(codigos, start=1):
            worksheet2.write(row_num, 0, codigo)
            
            # Fórmula BUSCARV dinámica
            formula = f'=BUSCARV(A{row_num+1},\'FORMULARIO RENTA SOCIEDADES-1\'!A:Q,17,FALSO)'
            worksheet2.write_formula(row_num, 1, formula)
        
        # Ajustar ancho columnas
        worksheet2.set_column(0, 1, 25)
        
        # Fijar encabezado
        worksheet2.freeze_panes(1, 0)
    
    output.seek(0)
    
    st.success("Archivo generado correctamente")
    
    st.download_button(
        label="Descargar Excel Generado",
        data=output,
        file_name="FORMULARIO_PROCESADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
