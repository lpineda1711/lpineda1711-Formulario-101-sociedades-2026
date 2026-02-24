import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Generador Final - Reemplazo F101")

uploaded_f101 = st.file_uploader(
    "Sube el nuevo F101 (.xlsx)",
    type=["xlsx"]
)

if uploaded_f101:

    # Abrir plantilla base (debe estar en el repo)
    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"
    wb = load_workbook(plantilla_path)

    # Eliminar F101 original
    if "F101" in wb.sheetnames:
        std = wb["F101"]
        wb.remove(std)

    # Cargar el archivo que subes
    wb_nuevo = load_workbook(uploaded_f101)
    hoja_nueva = wb_nuevo.active

    # Crear nueva hoja F101 en la plantilla
    hoja_f101 = wb.create_sheet("F101")

    # Copiar contenido exactamente
    for row in hoja_nueva.iter_rows():
        for cell in row:
            hoja_f101[cell.coordinate].value = cell.value

    # Reordenar hojas para que quede:
    # BC primero, F101 después
    wb._sheets.sort(key=lambda ws: ws.title != "BC")

    # Guardar resultado
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("F101 reemplazado correctamente. BC vinculado automáticamente.")

    st.download_button(
        "Descargar archivo final",
        data=output,
        file_name="ARCHIVO_FINAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
