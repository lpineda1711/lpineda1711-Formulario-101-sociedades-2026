import streamlit as st
from openpyxl import load_workbook
from copy import copy
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Generador Final F101 Vinculado con BC")

uploaded_f101 = st.file_uploader(
    "Sube el F101 (.xlsx)",
    type=["xlsx"]
)

if uploaded_f101:

    plantilla_path = "CT 2024 f101 ARCTURUS NUEVO.xlsx"
    wb_base = load_workbook(plantilla_path)

    # Eliminar F101 original si existe
    if "F101" in wb_base.sheetnames:
        wb_base.remove(wb_base["F101"])

    # Cargar archivo subido
    wb_nuevo = load_workbook(uploaded_f101)
    hoja_origen = wb_nuevo.active

    # Crear nueva hoja F101
    hoja_destino = wb_base.create_sheet("F101")

    # ==========================
    # COPIA COMPLETA CON ESTILOS
    # ==========================
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

    # ==========================
    # INSERTAR FORMULAS AUTOMATICAS
    # ==========================
    hoja_bc = wb_base["BC"]
    hoja_f101 = wb_base["F101"]

    for row in range(1, hoja_f101.max_row + 1):
        codigo_cell = hoja_f101[f"N{row}"]

        if codigo_cell.value is not None:
            hoja_f101[f"O{row}"] = f'=SUMAR.SI(BC!E:E;N{row};BC!D:D)'

    # ==========================
    # DEJAR SOLO BC Y F101
    # ==========================
    for sheet in wb_base.sheetnames:
        if sheet not in ["BC", "F101"]:
            wb_base.remove(wb_base[sheet])

    # Ordenar hojas
    wb_base._sheets.sort(key=lambda ws: ws.title != "BC")

    # Guardar archivo final
    output = BytesIO()
    wb_base.save(output)
    output.seek(0)

    st.success("F101 pegado tal cual y vinculado con BC correctamente")

    st.download_button(
        "Descargar archivo final",
        data=output,
        file_name="ARCHIVO_FINAL_VINCULADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
