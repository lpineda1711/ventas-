import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes

st.title("📊 Ventas SRI - OCR PRO (YA FUNCIONA TODO)")

uploaded_files = st.file_uploader(
    "Sube facturas PDF",
    type="pdf",
    accept_multiple_files=True
)

# -------- OCR SI FALLA PDF --------
def leer_pdf(file):
    texto = ""

    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texto += t + " "
    except:
        pass

    # SI NO HAY TEXTO → OCR
    if len(texto.strip()) < 50:
        images = convert_from_bytes(file.read())
        for img in images:
            texto += pytesseract.image_to_string(img)

    return texto

def limpiar(texto):
    texto = texto.replace("\n", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto

def buscar(patrones, texto):
    for p in patrones:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""

def extraer_datos(file):
    texto = leer_pdf(file)
    texto = limpiar(texto)

    # -------- CAMPOS CLAVE --------

    fecha = buscar([
        r"FECHA Y HORA DE AUTORIZACI[ÓO]N\s*:\s*(\d{2}/\d{2}/\d{4})"
    ], texto)

    autorizacion = buscar([
        r"N[ÚU]MERO DE AUTORIZACI[ÓO]N\s*:\s*(\d+)"
    ], texto)

    factura = buscar([
        r"No\.?\s*[:\-]?\s*(\d{3}-\d{3}-\d+)"
    ], texto)

    cliente = buscar([
        r"Raz[oó]n Social.*?:\s*(.*?)\s{2,}"
    ], texto)

    ruc = buscar([
        r"RUC\s*:\s*(\d{10,13})"
    ], texto)

    base_0 = buscar([r"0%.*?(\d+\.\d+)"], texto)
    base_15 = buscar([r"(?:12%|15%).*?(\d+\.\d+)"], texto)
    iva = buscar([r"IVA.*?(\d+\.\d+)"], texto)
    total = buscar([r"TOTAL.*?(\d+\.\d+)"], texto)

    return {
        "FECHA": fecha,
        "CLIENTE": cliente,
        "RUC": ruc,
        "FACT": factura,
        "AUTORIZACION": autorizacion,
        "NO OBJETO": "",
        "EXCENTO IVA": "",
        "BASE 0%": float(base_0) if base_0 else 0,
        "BASE 15%": float(base_15) if base_15 else 0,
        "PROPINA": "",
        "IVA": float(iva) if iva else 0,
        "TOTAL": float(total) if total else 0,
        "N° RETENCION": "NA",
        "10% R.FTE": "",
        "100% R. IVA": "",
        "TOTAL RETENCION": "",
        "POR COBRAR": ""
    }

# -------- PROCESAR --------
if uploaded_files:
    data = []

    for file in uploaded_files:
        data.append(extraer_datos(file))

    df = pd.DataFrame(data)
    st.dataframe(df)

    # -------- EXCEL --------
    wb = Workbook()
    ws = wb.active
    ws.title = "VENTAS FEBRERO"

    headers = list(df.columns)

    ws["A1"] = "VENTAS FEBRERO"
    ws.append(headers)

    amarillo = PatternFill(start_color="FFFF00", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", fill_type="solid")
    verde = PatternFill(start_color="92D050", fill_type="solid")

    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # ENCABEZADOS
    for col_num, col_name in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.fill = amarillo
        cell.font = Font(bold=True)
        cell.border = borde

        if col_name in ["N° RETENCION", "10% R.FTE", "100% R. IVA", "TOTAL RETENCION"]:
            cell.fill = azul

        if col_name == "POR COBRAR":
            cell.fill = verde

    # DATOS
    for i, row in enumerate(df.itertuples(index=False), start=3):
        ws.append(row)

        for col in range(1, len(headers) + 1):
            ws.cell(row=i, column=col).border = borde

        col_total = headers.index("TOTAL") + 1
        col_pc = headers.index("POR COBRAR") + 1

        letra = chr(64 + col_total)
        ws.cell(row=i, column=col_pc).value = f"={letra}{i}"

    # TOTAL FINAL
    fila_total = len(df) + 3
    ws.cell(row=fila_total, column=1, value="TOTAL")

    for col_name in ["BASE 0%", "BASE 15%", "IVA", "TOTAL", "POR COBRAR"]:
        col_index = headers.index(col_name) + 1
        letra = chr(64 + col_index)
        ws.cell(row=fila_total, column=col_index).value = f"=SUM({letra}3:{letra}{fila_total-1})"

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=fila_total, column=col)
        cell.fill = amarillo
        cell.border = borde

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📥 Descargar Excel FINAL",
        data=output,
        file_name="ventas_febrero.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
