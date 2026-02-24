import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Generador F101 Oficial - Renta Sociedades")

uploaded_formulario = st.file_uploader(
    "Sube el FORMULARIO RENTA SOCIEDADES-1 (.xlsx)",
    type=["xlsx"]
)

if uploaded_formulario:

    # Cargar plantilla base (debe estar en el repo)
    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"
    wb = load_workbook(plantilla_path)

    # Cargar formulario subido
    wb_form = load_workbook(uploaded_formulario)
    hoja_form_original = wb_form.active

    # Crear hoja nueva dentro de la plantilla
    hoja_form = wb.create_sheet("FORMULARIO_SUBIDO")

    # Copiar contenido del formulario a la nueva hoja
    for row in hoja_form_original.iter_rows():
        for cell in row:
            hoja_form[cell.coordinate].value = cell.value

    # ===== MODIFICAR FORMULAS EN F101 =====
    hoja_f101 = wb["F101"]

    for row in hoja_f101.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                if "Casilleros!" in cell.value:
                    cell.value = cell.value.replace(
                        "Casilleros!",
                        "FORMULARIO_SUBIDO!"
                    )

    # Guardar archivo final
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("Archivo generado con formato original intacto.")

    st.download_button(
        "Descargar F101 generado",
        data=output,
        file_name="F101_GENERADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
