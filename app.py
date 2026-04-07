import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("📊 Generador de Ventas SRI desde PDFs")

uploaded_files = st.file_uploader(
    "Sube tus facturas en PDF",
    type="pdf",
    accept_multiple_files=True
)

def extraer_datos(pdf):
    texto = ""
    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            if page.extract_text():
                texto += page.extract_text() + "\n"

    # REGEX (más flexibles)
    cliente = re.search(r"Raz[oó]n Social.*?:\s*(.*)", texto, re.IGNORECASE)
    ruc = re.search(r"R\.?U\.?C\.?:\s*(\d+)", texto)
    autorizacion = re.search(r"Autorizaci[oó]n.*?:\s*(\d+)", texto, re.IGNORECASE)

    base_0 = re.search(r"0%.*?(\d+\.\d+)", texto)
    base_15 = re.search(r"(12%|15%).*?(\d+\.\d+)", texto)
    iva = re.search(r"IVA.*?(\d+\.\d+)", texto)
    total = re.search(r"TOTAL.*?(\d+\.\d+)", texto)

    return {
        "CLIENTE": cliente.group(1).strip() if cliente else "",
        "RUC": ruc.group(1) if ruc else "",
        "AUTORIZACION": autorizacion.group(1) if autorizacion else "",
        "BASE 0%": float(base_0.group(1)) if base_0 else 0,
        "BASE 15%": float(base_15.group(2)) if base_15 else 0,
        "IVA": float(iva.group(1)) if iva else 0,
        "TOTAL": float(total.group(1)) if total else 0,
        "N° RETENCION": "NA"
    }

if uploaded_files:
    data = []

    for file in uploaded_files:
        try:
            datos = extraer_datos(file)
            data.append(datos)
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")

    df = pd.DataFrame(data)

    st.subheader("📋 Datos extraídos")
    st.dataframe(df)

    # -------- CREAR EXCEL ----------
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas Enero"

    headers = list(df.columns)
    ws.append(headers)

    # Colores tipo tu imagen
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")

    # Encabezados con color
    for col_num, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = amarillo
        cell.font = Font(bold=True)

        if col_name == "N° RETENCION":
            cell.fill = azul

    # Agregar datos
    for row in df.itertuples(index=False):
        ws.append(row)

    # -------- TOTALES ----------
    fila_total = len(df) + 2
    ws.cell(row=fila_total, column=4, value="TOTAL")

    for col in ["BASE 0%", "BASE 15%", "IVA", "TOTAL"]:
        col_index = headers.index(col) + 1
        suma = df[col].sum()
        ws.cell(row=fila_total, column=col_index, value=suma)

    # -------- DESCARGA ----------
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="ventas_sri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
