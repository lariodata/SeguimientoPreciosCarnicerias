import os
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import xlwings as xw
import smtplib
from email.message import EmailMessage

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

EXCEL_NAME = "WFW_SPC.xlsm"
EXCEL_PATH = PROJECT_DIR / EXCEL_NAME

TEMP_DIR = PROJECT_DIR / "temp_ventas"
TEMP_DIR.mkdir(exist_ok=True)

SHEET_CLUSTERS = "Clusters"
SHEET_PARAM = "Parametros"
SHEET_VENTAS = "Ventas"

SMTP_SERVER = "10.10.11.240"
SMTP_PORT = 20025

RUBROS_VALIDOS = ["Carnes", "Fiambres", "Dira"]

FECHA_ARCH = datetime.now().strftime("%d %B %Y")
FECHA_ARCH = FECHA_ARCH.replace("January","Enero").replace("February","Febrero") \
    .replace("March","Marzo").replace("April","Abril").replace("May","Mayo") \
    .replace("June","Junio").replace("July","Julio").replace("August","Agosto") \
    .replace("September","Septiembre").replace("October","Octubre") \
    .replace("November","Noviembre").replace("December","Diciembre")

# =====================================================
# HELPERS
# =====================================================
def clean_str(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s

def split_emails(s):
    s = clean_str(s)
    if not s:
        return []
    partes = re.split(r"[;,]", s)
    return [e.strip() for e in partes if e.strip()]

def safe_filename(text):
    text = str(text).strip()
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    return text

def get_or_open_workbook(app, name, path):
    try:
        return app.books[name]
    except Exception:
        return app.books.open(str(path))

# =====================================================
# CONECTAR A EXCEL
# =====================================================
try:
    app = xw.apps.active
except Exception:
    app = xw.App(visible=False)

wb = get_or_open_workbook(app, EXCEL_NAME, EXCEL_PATH)

sht_clusters = wb.sheets[SHEET_CLUSTERS]
sht_param = wb.sheets[SHEET_PARAM]
sht_ventas = wb.sheets[SHEET_VENTAS]

# =====================================================
# LEER PARÁMETROS MAIL (B14–B18)
# =====================================================
MAIL_FROM = clean_str(sht_param.range("B14").value)
MAIL_TO   = split_emails(sht_param.range("B15").value)
MAIL_CC   = split_emails(sht_param.range("B16").value)
ASUNTO    = clean_str(sht_param.range("B17").value)
CUERPO    = clean_str(sht_param.range("B18").value)

# =====================================================
# TEXTO A1 POR CLUSTER (HOJA VENTAS)
# =====================================================
A1_POR_CLUSTER = {
    "Casilda": sht_ventas.range("B19").value,
    "Rafaela": sht_ventas.range("B20").value,
    "Maria Luisa": sht_ventas.range("B21").value,
    "Horeca": sht_ventas.range("B22").value,
}

# =====================================================
# DETECTAR CLUSTERS (FILA 2)
# =====================================================
fila_cluster = 2
ultima_col = sht_clusters.used_range.last_cell.column

clusters = []
col = 1
while col <= ultima_col:
    val = sht_clusters.range((fila_cluster, col)).value
    if val:
        col_inicio = col
        col_fin = col
        c = col + 1
        while c <= ultima_col and not sht_clusters.range((fila_cluster, c)).value:
            col_fin = c
            c += 1
        clusters.append({
            "nombre": str(val).replace("CLUSTER","").strip(),
            "col_inicio": col_inicio,
            "col_fin": col_fin
        })
        col = col_fin + 1
    else:
        col += 1

# =====================================================
# FUNCIÓN EXPORTAR EXCEL VENTAS
# =====================================================
def exportar_excel_ventas(df, path_excel, texto_a1):
    app_tmp = xw.App(visible=False)
    app_tmp.api.DisplayAlerts = False
    wb_tmp = app_tmp.books.add()
    sht = wb_tmp.sheets[0]

    # A1
    sht.range("A1").value = texto_a1

    # Datos desde fila 2
    sht.range("A2").options(index=False, header=False).value = df

    last_row = 1 + len(df)

    # Formato columna B (precio)
    sht.range((2,2),(last_row,2)).api.NumberFormat = "#.##0,00"

    # Formato columna C (0,00)
    sht.range((2,3),(last_row,3)).api.NumberFormat = "0,00"

    # Fecha fija columna D
    sht.range((2,4),(last_row,4)).api.NumberFormat = "dd/mm/yyyy"

    sht.autofit()
    wb_tmp.save(str(path_excel))
    wb_tmp.close()
    app_tmp.quit()

# =====================================================
# PROCESO PRINCIPAL
# =====================================================
adjuntos = []

for c in clusters:
    # Leer tabla completa del cluster
    rango = sht_clusters.range(
        (fila_cluster + 1, c["col_inicio"]),
        (sht_clusters.used_range.last_cell.row, c["col_fin"])
    )

    df = rango.options(pd.DataFrame, header=False, index=False).value
    df = df.dropna(how="all")

    if df.empty:
        continue

    # Columnas esperadas:
    # 0 Rubro | 1 Código | 2 Precio
    df.columns = ["Rubro", "Codigo", "Precio"]

    # Limpiar rubro
    df["Rubro"] = df["Rubro"].astype(str).str.strip()
    df = df[df["Rubro"].isin(RUBROS_VALIDOS)]

    if df.empty:
        continue

    for rubro in RUBROS_VALIDOS:
        df_r = df[df["Rubro"] == rubro].copy()
        if df_r.empty:
            continue

        df_out = pd.DataFrame({
            "Codigo": df_r["Codigo"],
            "Precio": pd.to_numeric(df_r["Precio"], errors="coerce"),
            "Cero": 0.00,
            "Fecha": pd.to_datetime("2019-01-01")
        })

        nombre_arch = f"{FECHA_ARCH} - {c['nombre']} {rubro}.xlsx"
        path_excel = TEMP_DIR / safe_filename(nombre_arch)

        texto_a1 = A1_POR_CLUSTER.get(c["nombre"], "")
        exportar_excel_ventas(df_out, path_excel, texto_a1)

        adjuntos.append(path_excel)

# =====================================================
# ENVÍO DE CORREO (UN SOLO MAIL)
# =====================================================
if MAIL_TO and adjuntos:
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(MAIL_TO)
    if MAIL_CC:
        msg["Cc"] = ", ".join(MAIL_CC)

    msg["Subject"] = ASUNTO or "Precios de Venta"
    msg.set_content(CUERPO)

    for path in adjuntos:
        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=path.name
            )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15) as server:
        server.send_message(msg)
