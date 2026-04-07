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

# -------- FUNCION MEJORADA --------
def buscar(patron, texto):
    match = re.search(patron, texto, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else ""

def extraer_datos(pdf):
    texto = ""

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    # -------- EXTRACCIONES ROBUSTAS --------

    # CLIENTE
    cliente = buscar(r"Raz[oó]n Social.*?:\s*(.+)", texto)

    # RUC
    ruc = buscar(r"R\.?U\.?C\.?\s*:\s*(\d{10,13})", texto)

    # AUTORIZACION (MUY IMPORTANTE - VARIAS FORMAS)
    autorizacion = buscar(r"(?:N[úu]mero\s+de\s+Autorizaci[oó]n|Autorizaci[oó]n)\s*:\s*(\d{10,})", texto)

    # FECHA (solo fecha, ignora hora)
    fecha_raw = buscar(r"FECHA\s+Y\s+HORA\s+DE\s+AUTORIZACI[oó]N\s*:\s*([0-9/\-]+)", texto)

    # LIMPIAR FECHA
    fecha = ""
    if fecha_raw:
        fecha = fecha_raw.split()[0]

    # FACTURA (varios formatos posibles)
    factura = buscar(r"(?:Factura|No\.?\s*Factura|Comprobante)\s*[:#]?\s*(\d{3}-\d{3}-\d+)", texto)

    # BASES
    base_0 = buscar(r"0%\s*.*?(\d+\.\d+)", texto)
    base_15 = buscar(r"(?:12%|15%)\s*.*?(\d+\.\d+)", texto)

    # IVA y TOTAL
    iva = buscar(r"IVA\s*.*?(\d+\.\d+)", texto)
    total = buscar(r"TOTAL\s*.*?(\d+\.\d+)", texto)

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

        # FORMULA POR COBRAR
        col_total = headers.index("TOTAL") + 1
        col_pc = headers.index("POR COBRAR") + 1

        letra_total = chr(64 + col_total)
        ws.cell(row=i, column=col_pc).value = f"={letra_total}{i}"

    # FILA TOTAL
    fila_total = len(df) + 3
    ws.cell(row=fila_total, column=1, value="TOTAL")

    for col_name in ["BASE 0%", "BASE 15%", "IVA", "TOTAL", "POR COBRAR"]:
        col_index = headers.index(col_name) + 1
        letra = chr(64 + col_index)

        ws.cell(row=fila_total, column=col_index).value = f"=SUM({letra}3:{letra}{fila_total-1})"

    # COLOR TOTAL
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
