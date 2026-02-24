import streamlit as st
from openpyxl import load_workbook
from copy import copy
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Generador FV101 vinculado con BCD")

uploaded_fv101 = st.file_uploader(
    "Sube el FV101 (.xlsx)",
    type=["xlsx"]
)

if uploaded_fv101:

    # ===============================
    # CARGAR PLANTILLA BASE
    # ===============================
    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"
    wb_base = load_workbook(plantilla_path)

    # Eliminar FV101 anterior
    if "FV101" in wb_base.sheetnames:
        wb_base.remove(wb_base["FV101"])

    # ===============================
    # CARGAR FV101 SUBIDO
    # ===============================
    wb_nuevo = load_workbook(uploaded_fv101)
    hoja_origen = wb_nuevo.active

    hoja_destino = wb_base.create_sheet("FV101")

    # ===============================
    # COPIAR FV101 CON FORMATOS
    # ===============================
    for row in hoja_origen.iter_rows():
        for cell in row:
            new_cell = hoja_destino[cell.coordinate]
            new_cell.value = cell.value

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    # Copiar celdas combinadas
    for merged in hoja_origen.merged_cells.ranges:
        hoja_destino.merge_cells(str(merged))

    # Copiar ancho columnas
    for col in hoja_origen.column_dimensions:
        hoja_destino.column_dimensions[col].width = hoja_origen.column_dimensions[col].width

    # ===============================
    # VINCULAR FV101 CON BCD
    # ===============================
    if "BCD" in wb_base.sheetnames:

        merged_ranges = hoja_destino.merged_cells.ranges
        fila_inicio = 1
        fila_fin = hoja_destino.max_row

        for fila in range(fila_inicio, fila_fin + 1):
            codigo_cell = hoja_destino[f"N{fila}"]
            resultado_cell = hoja_destino[f"O{fila}"]

            # Si no hay código, no hacer nada
            if codigo_cell.value in (None, ""):
                continue

            # Verificar si O está combinada
            escribir = True
            for merged in merged_ranges:
                if resultado_cell.coordinate in merged:
                    if resultado_cell.coordinate != merged.start_cell.coordinate:
                        escribir = False
                    break

            if not escribir:
                continue

            # Insertar fórmula de sumatoria por código
            resultado_cell.value = (
                f'=SUMAR.SI(BCD!$E:$E;FV101!N{fila};BCD!$D:$D)'
            )

    # ===============================
    # EXPORTAR ARCHIVO
    # ===============================
    output = BytesIO()
    wb_base.save(output)
    output.seek(0)

    st.success("FV101 vinculado correctamente: sumatoria por código en columna O")

    st.download_button(
        "Descargar archivo final",
        data=output,
        file_name="FV101_VINCULADO_BCD.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
