import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

st.title("📊 Ventas SRI - Formato PRO")

uploaded_files = st.file_uploader(
    "Sube facturas PDF",
    type="pdf",
    accept_multiple_files=True
)

# -------- FUNCIONES --------
def limpiar_numero(valor):
    if not valor:
        return 0
    valor = valor.replace(",", ".")
    try:
        return float(valor)
    except:
        return 0

def buscar(patrones, texto):
    for patron in patrones:
        match = re.search(patron, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""

# -------- EXTRACCION --------
def extraer_datos(pdf):
    texto = ""

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    # -------- CLIENTE --------
    cliente = buscar([
        r"Raz[oó]n Social\s*:\s*(.+)",
        r"Cliente\s*:\s*(.+)"
    ], texto)

    # -------- RUC --------
    ruc = buscar([
        r"R\.?U\.?C\.?\s*:\s*(\d{10,13})",
        r"Identificaci[oó]n\s*:\s*(\d{10,13})"
    ], texto)

    # -------- AUTORIZACION --------
    autorizacion = buscar([
        r"Autorizaci[oó]n\s*:\s*(\d+)",
        r"N[úu]mero\s+de\s+Autorizaci[oó]n\s*:\s*(\d+)"
    ], texto)

    # -------- FECHA --------
    fecha = buscar([
        r"FECHA\s*Y\s*HORA\s*DE\s*AUTORIZACI[oó]N\s*:\s*([0-9/\-]+)",
        r"Fecha\s*:\s*([0-9/\-]+)"
    ], texto)

    # -------- FACTURA --------
    factura = buscar([
        r"Factura\s*No\.?\s*:\s*([\d\-]+)",
        r"No\.?\s*Factura\s*:\s*([\d\-]+)",
        r"Comprobante\s*:\s*([\d\-]+)"
    ], texto)

    # -------- BASES --------
    base_0 = limpiar_numero(buscar([
        r"0%\s*\$?\s*([\d\.,]+)"
    ], texto))

    base_15 = limpiar_numero(buscar([
        r"(?:12%|15%)\s*\$?\s*([\d\.,]+)"
    ], texto))

    # -------- IVA --------
    if base_15 > 0:
        iva = round(base_15 * 0.15, 2)
    else:
        iva = 0

    # -------- TOTAL --------
    total = limpiar_numero(buscar([
        r"TOTAL\s*\$?\s*([\d\.,]+)"
    ], texto))

    # fallback si no encuentra total
    if total == 0:
        total = base_0 + base_15 + iva

    return {
        "FECHA": fecha,
        "CLIENTE": cliente,
        "RUC": ruc,
        "FACT": factura,
        "AUTORIZACION": autorizacion,
        "NO OBJETO": "",
        "EXCENTO IVA": "",
        "BASE 0%": base_0,
        "BASE 15%": base_15,
        "PROPINA": "",
        "IVA": iva,
        "TOTAL": total,
        "N° RETENCION": "NA",
        "10% R.FTE": "",
        "100% R. IVA": "",
        "TOTAL RETENCION": "",
        "POR COBRAR": total  # ← IGUAL AL TOTAL
    }

# -------- PROCESO --------
if uploaded_files:
    data = []

    for file in uploaded_files:
        try:
            datos = extraer_datos(file)
            data.append(datos)
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

    # COLORES
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    verde = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")

    # BORDES
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

    # TOTAL
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
