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

LOGO_PATH = BASE_DIR.parent / "logo_ra.png"
CM_TO_POINTS = 28.3465
LOGO_HEIGHT_PT = 1.67 * CM_TO_POINTS   # 47.3 pt
LOGO_WIDTH_PT  = 2.96 * CM_TO_POINTS   # 83.9 pt

DATA_START_ROW = 6
HEADER_ROW     = 5
LOGO_ROW_COUNT = 3   # filas 1-3 suman el alto del logo

# xlConstants alineación
XL_ALIGN_CENTER = -4108
XL_ALIGN_LEFT   = -4131

# xlConstants para bordes
XL_EDGE_LEFT       = 7
XL_EDGE_TOP        = 8
XL_EDGE_BOTTOM     = 9
XL_EDGE_RIGHT      = 10
XL_INSIDE_VERTICAL = 11
XL_INSIDE_HORIZ    = 12
XL_CONTINUOUS      = 1
XL_THIN            = 2

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

    sht_base = wb_tmp.sheets[0]

    for i, (nombre_tab, rubro) in enumerate(tabs.items()):

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

        headers = df_tab.columns.tolist()
        last_row = DATA_START_ROW + len(df_tab) - 1

        # ==============================
        # LOGO EN A1 (filas 1-3 suman el alto total)
        # ==============================
        row_h = LOGO_HEIGHT_PT / LOGO_ROW_COUNT
        for r in range(1, LOGO_ROW_COUNT + 1):
            sht.api.Rows(r).RowHeight = row_h
        if LOGO_PATH.exists():
            sht.pictures.add(
                str(LOGO_PATH),
                left=sht.range("A1").left,
                top=sht.range("A1").top,
                width=LOGO_WIDTH_PT,
                height=LOGO_HEIGHT_PT,
            )

        # ==============================
        # TÍTULO Y FECHA
        # ==============================
        sht.range("C2").value = f"LISTA {nombre_cluster}"
        sht.range("C2").api.Font.Bold = True

        sht.range("E2").value = fecha_hoy
        sht.range("E2").api.Font.Bold = True

        # ==============================
        # ENCABEZADO (fondo negro, letra blanca, centrado)
        # ==============================
        sht.range((HEADER_ROW, 1)).value = headers
        header_rng = sht.range((HEADER_ROW, 1), (HEADER_ROW, len(cols_out)))
        header_rng.color = (0, 0, 0)
        header_rng.api.Font.Color = 0xFFFFFF
        header_rng.api.Font.Bold = True
        header_rng.api.HorizontalAlignment = XL_ALIGN_CENTER

        # ==============================
        # DATOS
        # ==============================
        sht.range((DATA_START_ROW, 1)).options(index=False, header=False).value = df_tab

        # ==============================
        # FORMATO MONEDA
        # ==============================
        for col_name in [col_sin_iva, col_precio]:
            if col_name in headers:
                pos = headers.index(col_name) + 1
                sht.range((DATA_START_ROW, pos), (last_row, pos)).api.NumberFormat = FORMATO_MONEDA

                for r in range(DATA_START_ROW, last_row + 1):
                    val = sht.range((r, pos)).value
                    if val is not None and str(int(val)).endswith("99"):
                        sht.range((r, pos)).api.Font.Color = COLOR_ROJO_FONT

        # ==============================
        # FORMATO VS LISTA ANTERIOR
        # ==============================
        if col_vs_lista in headers:
            pos_vs = headers.index(col_vs_lista) + 1
            sht.range((DATA_START_ROW, pos_vs), (last_row, pos_vs)).api.NumberFormat = "0,00%"

            COLOR_ROJO = 255
            COLOR_VERDE = 5287936
            COLOR_NEGRO = 0

            for j, val in enumerate(df_tab[col_vs_lista]):
                try:
                    if pd.notna(val):
                        fila_excel = DATA_START_ROW + j
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

        # ==============================
        # ALINEACIÓN COLUMNAS DE DATOS
        # ==============================
        pos_desc = headers.index(col_desc) + 1 if col_desc and col_desc in headers else None
        for col_idx in range(1, len(cols_out) + 1):
            col_rng = sht.range((DATA_START_ROW, col_idx), (last_row, col_idx))
            if col_idx == pos_desc:
                col_rng.api.HorizontalAlignment = XL_ALIGN_LEFT
            else:
                col_rng.api.HorizontalAlignment = XL_ALIGN_CENTER

        # ==============================
        # FILL GRIS + NEGRITA EN Rubro y Tipo
        # ==============================
        for col_name in [col_rubro, col_tipo]:
            if col_name and col_name in headers:
                pos = headers.index(col_name) + 1
                rng = sht.range((DATA_START_ROW, pos), (last_row, pos))
                rng.color = (217, 217, 217)
                rng.api.Font.Bold = True

        # ==============================
        # BORDES FINOS NEGROS EN TODA LA TABLA
        # ==============================
        if len(df_tab) > 0:
            tabla_rng = sht.range((DATA_START_ROW, 1), (last_row, len(cols_out)))
            for edge in [XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT,
                         XL_INSIDE_VERTICAL, XL_INSIDE_HORIZ]:
                tabla_rng.api.Borders(edge).LineStyle = XL_CONTINUOUS
                tabla_rng.api.Borders(edge).Weight = XL_THIN
                tabla_rng.api.Borders(edge).Color = 0  # negro

        sht.autofit()

    wb_tmp.save(str(path_excel))
    wb_tmp.close()
    app_tmp.quit()

# =====================================================
# PROCESO PRINCIPAL
# =====================================================

# -- Pasada 1: leer todos los clusters y construir tipo_map desde el que tenga "Tipo"
cluster_data = []
tipo_map = {}

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

    cluster_data.append((c, df_cluster))

    # Tomar tipo_map del primer cluster que tenga "Tipo" y "Cód."
    if not tipo_map:
        _col_tipo = find_col(df_cluster, ["Tipo"])
        _col_cod  = find_col(df_cluster, ["Cód.", "Cod", "Código"])
        if _col_tipo and _col_cod:
            tipo_map = (
                df_cluster[[_col_cod, _col_tipo]]
                .dropna(subset=[_col_cod])
                .drop_duplicates(subset=[_col_cod])
                .set_index(_col_cod.strip() if hasattr(_col_cod, "strip") else _col_cod)[_col_tipo]
                .to_dict()
            )

# -- Pasada 2: completar "Tipo" en los que no la tienen y exportar
adjuntos = []

for c, df_cluster in cluster_data:
    _col_tipo = find_col(df_cluster, ["Tipo"])
    _col_cod  = find_col(df_cluster, ["Cód.", "Cod", "Código"])

    if _col_tipo is None and _col_cod and tipo_map:
        df_cluster = df_cluster.copy()
        df_cluster["Tipo"] = df_cluster[_col_cod].map(tipo_map).fillna("")

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
