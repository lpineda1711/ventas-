import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("📊 Generador de Ventas SRI desde PDFs")

uploaded_files = st.file_uploader("Sube tus facturas en PDF", type="pdf", accept_multiple_files=True)

def extraer_datos(pdf):
    texto = ""
    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            texto += page.extract_text() + "\n"

    # EXTRAER DATOS CON REGEX
    cliente = re.search(r"Raz[oó]n Social\s*/?\s*Nombres.*:\s*(.*)", texto)
    ruc = re.search(r"R\.U\.C\.:\s*(\d+)", texto)
    autorizacion = re.search(r"N[ÚU]MERO DE AUTORIZACI[ÓO]N\s*:\s*(\d+)", texto)

    base_0 = re.search(r"Base 0%.*?(\d+\.\d+)", texto)
    base_15 = re.search(r"Base 15%.*?(\d+\.\d+)", texto)
    iva = re.search(r"IVA.*?(\d+\.\d+)", texto)
    total = re.search(r"TOTAL.*?(\d+\.\d+)", texto)

    return {
        "CLIENTE": cliente.group(1) if cliente else "",
        "RUC": ruc.group(1) if ruc else "",
        "AUTORIZACION": autorizacion.group(1) if autorizacion else "",
        "BASE 0%": float(base_0.group(1)) if base_0 else 0,
        "BASE 15%": float(base_15.group(1)) if base_15 else 0,
        "IVA": float(iva.group(1)) if iva else 0,
        "TOTAL": float(total.group(1)) if total else 0,
        "N° RETENCION": "NA"
    }

if uploaded_files:
    data = []

    for file in uploaded_files:
        datos = extraer_datos(file)
        data.append(datos)

    df = pd.DataFrame(data)

    st.dataframe(df)

    # CREAR EXCEL CON COLORES
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"

    headers = list(df.columns)
    ws.append(headers)

    # Colores
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")

    # Aplicar encabezados
    for col_num, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = amarillo
        cell.font = Font(bold=True)

        if col_name in ["N° RETENCION"]:
            cell.fill = azul

    # Agregar datos
    for row in df.itertuples(index=False):
        ws.append(row)

    # Guardar en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="ventas_sri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
