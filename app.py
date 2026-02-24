import streamlit as st
from openpyxl import load_workbook
import tempfile
import shutil
import os

st.title("Generador F101 con Plantilla Oficial")

uploaded_formulario = st.file_uploader(
    "Sube el FORMULARIO RENTA SOCIEDADES-1 (.xlsx)",
    type=["xlsx"]
)

if uploaded_formulario is not None:

    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"

    with tempfile.TemporaryDirectory() as tmpdir:

        plantilla_temp = os.path.join(tmpdir, "plantilla.xlsx")
        formulario_temp = os.path.join(tmpdir, "formulario.xlsx")

        # Guardar archivos temporales
        with open(formulario_temp, "wb") as f:
            f.write(uploaded_formulario.read())

        shutil.copy(plantilla_path, plantilla_temp)

        # Cargar plantilla
        wb = load_workbook(plantilla_temp)

        # Cargar formulario subido
        wb_form = load_workbook(formulario_temp)

        hoja_form = wb_form.active
        wb._add_sheet(hoja_form)

        hoja_form.title = "FORMULARIO_SUBIDO"

        # Ahora actualizar fórmulas en F101
        hoja_f101 = wb["F101"]  # asegúrate que se llame así en tu plantilla

        for row in hoja_f101.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if "Casilleros!" in cell.value:
                        cell.value = cell.value.replace(
                            "Casilleros!",
                            "FORMULARIO_SUBIDO!"
                        )

        archivo_final = os.path.join(tmpdir, "F101_GENERADO.xlsx")
        wb.save(archivo_final)

        with open(archivo_final, "rb") as f:
            st.download_button(
                "Descargar Excel Generado",
                f,
                file_name="F101_GENERADO.xlsx"
            )

        st.success("Archivo generado con formato original intacto.")
