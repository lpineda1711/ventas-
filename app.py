import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

st.title("📊 Ventas SRI - Formato FINAL")

uploaded_files = st.file_uploader(
    "Sube facturas PDF",
    type="pdf",
    accept_multiple_files=True
)

# -------- LIMPIAR TEXTO --------
def limpiar_texto(texto):
    texto = texto.replace("\n", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto

# -------- BUSCADOR FLEXIBLE --------
def buscar(patrones, texto):
    for patron in patrones:
        match = re.search(patron, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""

def extraer_datos(pdf):
    texto = ""

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            t = page.extract_text()
            if t:
                texto += t + " "

    texto = limpiar_texto(texto)

    # -------- CAMPOS CLAVE --------

    # FECHA (SOLO FECHA)
    fecha_raw = buscar([
        r"FECHA Y HORA DE AUTORIZACI[ÓO]N\s*:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})",
        r"FECHA DE AUTORIZACI[ÓO]N\s*:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})"
    ], texto)

    fecha = fecha_raw

    # AUTORIZACION
    autorizacion = buscar([
        r"N[ÚU]MERO DE AUTORIZACI[ÓO]N\s*:\s*(\d+)",
        r"AUTORIZACI[ÓO]N\s*:\s*(\d+)"
    ], texto)

    # FACTURA (No.)
    factura = buscar([
        r"No\.?\s*:\s*(\d{3}-\d{3}-\d+)",
        r"FACTURA\s*:\s*(\d{3}-\d{3}-\d+)",
        r"COMPROBANTE\s*:\s*(\d{3}-\d{3}-\d+)"
    ], texto)

    # CLIENTE
    cliente = buscar([
        r"Raz[oó]n Social.*?:\s*(.*?)\s{2,}",
    ], texto)

    # RUC
    ruc = buscar([
        r"R\.?U\.?C\.?\s*:\s*(\d{10,13})"
    ], texto)

    # BASES
    base_0 = buscar([r"0%.*?(\d+\.\d+)"], texto)
    base_15 = buscar([r"(?:12%|15%).*?(\d+\.\d+)"], texto)

    # IVA Y TOTAL
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

# -------- PROCESO --------
if uploaded_files:
    data = []

    for file in uploaded_files:
        try:
            data.append(extraer_datos(file))
        except Exception as e:
            st.error(f"Error en {file.name}: {e}")

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

        # FORMULA POR COBRAR
        col_total = headers.index("TOTAL") + 1
        col_pc = headers.index("POR COBRAR") + 1

        letra = chr(64 + col_total)
        ws.cell(row=i, column=col_pc).value = f"={letra}{i}"

    # FILA TOTAL
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

    # DESCARGA
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📥 Descargar Excel FINAL",
        data=output,
        file_name="ventas_febrero.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
