import streamlit as st
from openpyxl import load_workbook
import tempfile
import shutil
import os

st.set_page_config(page_title="Generador F101 Oficial", layout="wide")

st.title("Generador F101 desde Formulario Renta Sociedades")

uploaded_formulario = st.file_uploader(
    "Sube el FORMULARIO RENTA SOCIEDADES-1 (.xlsx)",
    type=["xlsx"]
)

if uploaded_formulario is not None:

    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"

    with tempfile.TemporaryDirectory() as tmpdir:

        plantilla_temp = os.path.join(tmpdir, "plantilla.xlsx")
        formulario_temp = os.path.join(tmpdir, "formulario.xlsx")

        # Guardar formulario subido
        with open(formulario_temp, "wb") as f:
            f.write(uploaded_formulario.read())

        # Copiar plantilla
        shutil.copy(plantilla_path, plantilla_temp)

        # Cargar plantilla
        wb = load_workbook(plantilla_temp)

        # Cargar formulario
        wb_form = load_workbook(formulario_temp)
        hoja_form = wb_form.active

        hoja_form.title = "FORMULARIO_SUBIDO"

        wb._add_sheet(hoja_form)

        # Modificar fórmulas en F101
        hoja_f101 = wb["F101"]

        for row in hoja_f101.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if "Casilleros!" in cell.value:
                        cell.value = cell.value.replace(
                            "Casilleros!",
                            "FORMULARIO_SUBIDO!"
                        )

        archivo_final = os.path.join(tmpdir, "F101_GENERADO.xlsx")
        wb.save(archivo_final)

        with open(archivo_final, "rb") as f:
            st.download_button(
                "Descargar Excel con F101 generado",
                f,
                file_name="F101_GENERADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("Archivo generado correctamente con formato original intacto.")
