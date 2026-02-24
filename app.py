import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
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

    # Eliminar hoja F101 si existe
    if "F101" in wb_base.sheetnames:
        wb_base.remove(wb_base["F101"])

    # Cargar archivo subido directamente desde memoria
    wb_nuevo = load_workbook(uploaded_f101, data_only=False)
    hoja_origen = wb_nuevo.active

    # Crear nueva hoja F101 limpia
    hoja_destino = wb_base.create_sheet("F101")

    # ==========================
    # COPIAR SOLO VALORES (ESTABLE)
    # ==========================
    for row in hoja_origen.iter_rows():
        for cell in row:
            hoja_destino.cell(row=cell.row, column=cell.column).value = cell.value

    # Copiar merges correctamente
    for merged in hoja_origen.merged_cells.ranges:
        hoja_destino.merge_cells(str(merged))

    hoja_f101 = wb_base["F101"]

    # ==========================
    # INSERTAR FORMULAS SEGURAS
    # ==========================
    for row in range(1, hoja_f101.max_row + 1):

        codigo_cell = hoja_f101.cell(row=row, column=14)  # Columna N
        celda_resultado = hoja_f101.cell(row=row, column=15)  # Columna O

        if codigo_cell.value not in (None, ""):

            if not isinstance(celda_resultado, MergedCell):
                celda_resultado.value = f'=SUMAR.SI(BC!E:E;N{row};BC!D:D)'

    # Dejar solo BC y F101
    for sheet in wb_base.sheetnames:
        if sheet not in ["BC", "F101"]:
            wb_base.remove(wb_base[sheet])

    # Ordenar hojas
    wb_base._sheets.sort(key=lambda ws: ws.title != "BC")

    # Guardar archivo final
    output = BytesIO()
    wb_base.save(output)
    output.seek(0)

    st.success("Archivo generado sin corrupción y vinculado correctamente")

    st.download_button(
        "Descargar archivo final",
        data=output,
        file_name="ARCHIVO_FINAL_VINCULADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
