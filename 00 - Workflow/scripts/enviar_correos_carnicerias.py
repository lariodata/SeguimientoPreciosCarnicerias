import os
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import xlwings as xw
import smtplib
from email.message import EmailMessage
from pathlib import Path
from dotenv import load_dotenv

import sys


# =====================================================
# RECIBIR FECHA DESDE VBA (YYYY-MM-DD) / FALLBACK MANUAL
# =====================================================
if len(sys.argv) >= 2 and sys.argv[1]:
    try:
        FECHA_PROCESO = datetime.strptime(sys.argv[1], "%Y-%m-%d")
    except ValueError:
        raise ValueError("❌ Formato de fecha inválido. Se esperaba YYYY-MM-DD (ej: 2026-02-14).")
else:
    FECHA_PROCESO = datetime.now()
    print("⚠️ No se recibió fecha desde VBA. Se usa fecha actual.")

fecha_arch = FECHA_PROCESO.strftime("%d-%m-%Y")
fecha_excel_str = FECHA_PROCESO.strftime("%d-%b").lower()


# =====================================================
# CONFIGURACIÓN GENERAL / RUTAS
# =====================================================

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

EXCEL_NAME = "WFW_SPC.xlsm"
EXCEL_PATH = PROJECT_DIR / EXCEL_NAME

TEMP_DIR = PROJECT_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)

SHEET_CLUSTERS = "Clusters"
SHEET_PARAMETROS = "Parametros"


SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")          # Aplicaciones@rafalim.com
SMTP_PASS = os.getenv("SMTP_PASS")          # (password)

if not SMTP_USER or not SMTP_PASS:
    raise RuntimeError("❌ SMTP_USER / SMTP_PASS no configurados (revisar .env)")

# =====================================================
# HELPERS
# =====================================================

def enviar_mail_o365(msg, from_addr, to_addrs):
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg, from_addr=from_addr, to_addrs=to_addrs)

def clean_str(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s

def split_emails(s):
    s = clean_str(s)
    if not s:
        return []
    # permite ; o ,
    partes = re.split(r"[;,]", s)
    return [e.strip() for e in partes if e.strip()]


def get_or_open_workbook(app, name, path):
    try:
        return app.books[name]
    except Exception:
        return app.books.open(str(path))

def safe_filename(text):
    text = str(text).strip()
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text)
    return text

def norm_key(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower() if s else ""

def find_col(df, candidates):
    lookup = {norm_key(c): c for c in df.columns}
    for cand in candidates:
        if norm_key(cand) in lookup:
            return lookup[norm_key(cand)]
    return None

# =====================================================
# CONECTAR A EXCEL
# =====================================================
try:
    app = xw.apps.active
except Exception:
    app = xw.App(visible=False)

wb = get_or_open_workbook(app, EXCEL_NAME, EXCEL_PATH)
sht_clusters = wb.sheets[SHEET_CLUSTERS]
sht_param = wb.sheets[SHEET_PARAMETROS]

# =====================================================
# LEER PARÁMETROS MAIL
# =====================================================
MAIL_FROM = SMTP_USER
MAIL_TO   = split_emails(sht_param.range("B7").value)
MAIL_CC   = split_emails(sht_param.range("B8").value)
ASUNTO    = clean_str(sht_param.range("B9").value)
CUERPO    = clean_str(sht_param.range("B10").value)

# =====================================================
# DETECTAR CLUSTERS
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
            "nombre": str(val).strip(),
            "col_inicio": col_inicio,
            "col_fin": col_fin
        })
        col = col_fin + 1
    else:
        col += 1

# =====================================================
# LECTURA TABLAS
# =====================================================
def buscar_fila_encabezado(sht, col_inicio, max_filas=300):
    for fila in range(1, max_filas):
        val = sht.range((fila, col_inicio)).value
        if val and norm_key(val) == "rubro":
            return fila
    return None

def leer_tabla_cluster(sht, col_inicio, col_fin, fila_header):
    rango = sht.range(
        (fila_header, col_inicio),
        (sht.used_range.last_cell.row, col_fin)
    )
    df = rango.options(pd.DataFrame, header=1, index=False).value
    return df.dropna(how="all")

# =====================================================
# EXPORTAR EXCEL
# =====================================================
def exportar_excel_precios_xlwings(df, path_excel, nombre_cluster):
    app_tmp = xw.App(visible=False)
    app_tmp.api.DisplayAlerts = False
    wb_tmp = app_tmp.books.add()

    tabs = {"Carnes": "Carnes", "Fiambres": "Fiambres", "Reventa": "Dira"}
    fecha_hoy = fecha_excel_str

    COLOR_ROJO_FONT = 255
    FORMATO_MONEDA = "$ #.##0"

    col_rubro = find_col(df, ["Rubro"])
    col_tipo = find_col(df, ["Tipo"])
    col_cod = find_col(df, ["Cód.", "Cod", "Código"])
    col_desc = find_col(df, ["Descripción"])
    col_sin_iva = find_col(df, ["sin iva"])
    col_precio = find_col(df, ["$"])
    col_unidad = find_col(df, ["unidad_me"])
    col_vs_lista = find_col(df, ["VS LISTA ANTERIOR", "Vs lista anterior", "VS LISTA", "VS LISTA ANT"])

    cols_out = [c for c in [col_rubro, col_tipo, col_cod, col_desc,
                            col_sin_iva, col_precio, col_vs_lista, col_unidad] if c]

    # 👇 TOMAMOS LA HOJA INICIAL QUE CREA EXCEL
    sht_base = wb_tmp.sheets[0]

    for i, (nombre_tab, rubro) in enumerate(tabs.items()):

        # 👇 Primera hoja: reutilizamos la inicial
        if i == 0:
            sht = sht_base
            sht.name = nombre_tab
        else:
            sht = wb_tmp.sheets.add(nombre_tab, after=wb_tmp.sheets[-1])

        df_tab = df[df[col_rubro] == rubro].copy() if col_rubro else df.copy()
        df_tab = df_tab[cols_out].copy()

        # SIN IVA obligatorio
        if col_sin_iva:
            ser = pd.to_numeric(df_tab[col_sin_iva], errors="coerce").round(0)
            df_tab = df_tab[ser.notna() & (ser != 0)]
            df_tab[col_sin_iva] = ser.loc[df_tab.index].astype("Int64")

        # $ opcional
        if col_precio:
            df_tab[col_precio] = pd.to_numeric(df_tab[col_precio], errors="coerce").round(0).astype("Int64")

        # VS LISTA ANTERIOR
        if col_vs_lista and col_vs_lista in df_tab.columns:
            vs = df_tab[col_vs_lista]

            vs_num = (
                vs.astype(str)
                  .str.replace("%", "", regex=False)
                  .str.replace(",", ".", regex=False)
            )

            vs_num = pd.to_numeric(vs_num, errors="coerce")
            vs_num = vs_num.where(vs_num.abs() <= 1, vs_num / 100)

            df_tab[col_vs_lista] = vs_num

        # ==============================
        # ESCRIBIR EN EXCEL
        # ==============================
        sht.range("C2").value = f"LISTA {nombre_cluster}"
        sht.range("C2").api.Font.Bold = True

        sht.range("E2").value = fecha_hoy
        sht.range("E2").api.Font.Bold = True

        headers = df_tab.columns.tolist()
        sht.range("A4").value = headers
        sht.range("A4").expand("right").api.Font.Bold = True
        sht.range("A5").options(index=False, header=False).value = df_tab

        last_row = 4 + len(df_tab)

        # ==============================
        # FORMATO MONEDA
        # ==============================
        for col_name in [col_sin_iva, col_precio]:
            if col_name in headers:
                pos = headers.index(col_name) + 1
                sht.range((5, pos), (last_row, pos)).api.NumberFormat = FORMATO_MONEDA

                for r in range(5, last_row + 1):
                    val = sht.range((r, pos)).value
                    if val is not None and str(int(val)).endswith("99"):
                        sht.range((r, pos)).api.Font.Color = COLOR_ROJO_FONT

        # ==============================
        # FORMATO VS LISTA ANTERIOR
        # ==============================
        if col_vs_lista in headers:
            pos_vs = headers.index(col_vs_lista) + 1
            sht.range((5, pos_vs), (last_row, pos_vs)).api.NumberFormat = "0,00%"

            COLOR_ROJO = 255
            COLOR_VERDE = 5287936
            COLOR_NEGRO = 0

            for j, val in enumerate(df_tab[col_vs_lista]):
                try:
                    if pd.notna(val):
                        fila_excel = 5 + j
                        celda = sht.range((fila_excel, pos_vs)).api

                        celda.Font.Bold = False

                        valor = float(val)

                    if valor < 0:
                        celda.Font.Color = COLOR_ROJO
                        celda.Font.Bold = True

                    elif valor > 0:
                        celda.Font.Color = COLOR_VERDE
                        celda.Font.Bold = True

                    else:
                        celda.Font.Color = COLOR_NEGRO
                        celda.Font.Bold = False

                except Exception:
                    pass

        sht.autofit()

    wb_tmp.save(str(path_excel))
    wb_tmp.close()
    app_tmp.quit()

# =====================================================
# PROCESO PRINCIPAL
# =====================================================

adjuntos = []

for c in clusters:
    fila_header = buscar_fila_encabezado(sht_clusters, c["col_inicio"])
    if not fila_header:
        continue

    df_cluster = leer_tabla_cluster(
        sht_clusters,
        c["col_inicio"],
        c["col_fin"],
        fila_header
    )

    if df_cluster.empty:
        continue

    nombre_limpio = safe_filename(c["nombre"].replace("CLUSTER", "").strip())
    path_excel = TEMP_DIR / f"Precios_{nombre_limpio}_{fecha_arch}.xlsx"

    exportar_excel_precios_xlwings(df_cluster, path_excel, nombre_limpio)

    if path_excel.exists():
        adjuntos.append(path_excel)


# ==========================================
# ENVÍO DE UN SOLO MAIL (FUERA DEL FOR)
# ==========================================

if MAIL_TO and adjuntos:
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(MAIL_TO)

    if MAIL_CC:
        msg["Cc"] = ", ".join(MAIL_CC)

    msg["Subject"] = f"{ASUNTO} - {FECHA_PROCESO.strftime('%d/%m/%Y')}"

    msg.set_content(CUERPO)

    destinatarios = MAIL_TO + MAIL_CC

    for path_excel in adjuntos:
        with open(path_excel, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=path_excel.name
            )

    try:
        enviar_mail_o365(msg, MAIL_FROM, destinatarios)
        print(f"✅ Mail único enviado con {len(adjuntos)} adjuntos")
    except Exception as e:
        print(f"❌ Error enviando mail: {e}")
