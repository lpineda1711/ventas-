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

def extraer_datos(pdf):
    texto = ""

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            contenido = page.extract_text()
            if contenido:
                texto += contenido + "\n"

    # -------- EXTRACCIÓN --------
    cliente = re.search(r"Raz[oó]n Social.*?:\s*(.*)", texto, re.IGNORECASE)
    ruc = re.search(r"R\.?U\.?C\.?:\s*(\d+)", texto)
    autorizacion = re.search(r"N[ÚU]MERO DE AUTORIZACI[ÓO]N\s*:\s*(\d+)", texto)

    # FECHA (solo fecha sin hora)
    fecha = re.search(r"FECHA Y HORA DE AUTORIZACI[ÓO]N\s*:\s*([\d/]+)", texto)

    # FACTURA
    factura = re.search(r"Factura.*?(\d{3}-\d{3}-\d+)", texto)

    base_0 = re.search(r"0%.*?(\d+\.\d+)", texto)
    base_15 = re.search(r"(12%|15%).*?(\d+\.\d+)", texto)
    iva = re.search(r"IVA.*?(\d+\.\d+)", texto)
    total = re.search(r"TOTAL.*?(\d+\.\d+)", texto)

    return {
        "FECHA": fecha.group(1) if fecha else "",
        "CLIENTE": cliente.group(1).strip() if cliente else "",
        "RUC": ruc.group(1) if ruc else "",
        "FACT": factura.group(1) if factura else "",
        "AUTORIZACION": autorizacion.group(1) if autorizacion else "",
        "NO OBJETO": "",
        "EXCENTO IVA": "",
        "BASE 0%": float(base_0.group(1)) if base_0 else 0,
        "BASE 15%": float(base_15.group(2)) if base_15 else 0,
        "PROPINA": "",
        "IVA": float(iva.group(1)) if iva else 0,
        "TOTAL": float(total.group(1)) if total else 0,
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
            data.append(extraer_datos(file))
        except Exception as e:
            st.error(f"Error en {file.name}: {e}")

    df = pd.DataFrame(data)
    st.dataframe(df)

    # -------- CREAR EXCEL --------
    wb = Workbook()
    ws = wb.active
    ws.title = "VENTAS FEBRERO"

    headers = list(df.columns)

    # TÍTULO
    ws["A1"] = "VENTAS FEBRERO"

    # ENCABEZADOS
    ws.append(headers)

    # -------- COLORES --------
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    verde = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")

    # -------- BORDES --------
    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # ENCABEZADOS CON COLOR Y BORDE
    for col_num, col_name in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.fill = amarillo
        cell.font = Font(bold=True)
        cell.border = borde

        if col_name in ["N° RETENCION", "10% R.FTE", "100% R. IVA", "TOTAL RETENCION"]:
            cell.fill = azul

        if col_name == "POR COBRAR":
            cell.fill = verde

    # -------- DATOS --------
    for i, row in enumerate(df.itertuples(index=False), start=3):
        ws.append(row)

        for col in range(1, len(headers) + 1):
            ws.cell(row=i, column=col).border = borde

        # FORMULA POR COBRAR
        col_total = headers.index("TOTAL") + 1
        col_pc = headers.index("POR COBRAR") + 1

        letra_total = chr(64 + col_total)
        ws.cell(row=i, column=col_pc).value = f"={letra_total}{i}"

    # -------- FILA TOTAL --------
    fila_total = len(df) + 3
    ws.cell(row=fila_total, column=1, value="TOTAL")

    for col_name in ["BASE 0%", "BASE 15%", "IVA", "TOTAL", "POR COBRAR"]:
        col_index = headers.index(col_name) + 1
        letra = chr(64 + col_index)

        ws.cell(row=fila_total, column=col_index).value = f"=SUM({letra}3:{letra}{fila_total-1})"

    # COLOR + BORDE FILA TOTAL
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=fila_total, column=col)
        cell.fill = amarillo
        cell.border = borde

    # -------- DESCARGA --------
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📥 Descargar Excel PRO",
        data=output,
        file_name="ventas_febrero.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
