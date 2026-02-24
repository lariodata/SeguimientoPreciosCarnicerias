import os
import re
import sys
import smtplib
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

import pandas as pd
import xlwings as xw
from dotenv import load_dotenv


# =====================================================
# RECIBIR FECHA DESDE VBA (YYYY-MM-DD)
# =====================================================
if len(sys.argv) < 2:
    raise ValueError("❌ No se recibió la fecha desde Excel (se esperaba YYYY-MM-DD).")
try:
    FECHA_PROCESO = datetime.strptime(sys.argv[1], "%Y-%m-%d")
except ValueError:
    raise ValueError("❌ Formato de fecha inválido. Se esperaba YYYY-MM-DD (ej: 2026-02-14).")

# Formatos derivados
fecha_arch = FECHA_PROCESO.strftime("%d-%m-%Y")   # para nombres de archivo
fecha_excel_str = FECHA_PROCESO.strftime("%d/%m/%Y")  # para asunto/cuerpo si querés


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
SMTP_USER = os.getenv("SMTP_USER")  # Aplicaciones@rafalim.com
SMTP_PASS = os.getenv("SMTP_PASS")  # (password)

if not SMTP_USER or not SMTP_PASS:
    raise RuntimeError("❌ SMTP_USER / SMTP_PASS no configurados (revisar .env)")


# =====================================================
# HELPERS
# =====================================================
def enviar_mail_o365(msg: EmailMessage, from_addr: str, to_addrs: list[str]) -> None:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg, from_addr=from_addr, to_addrs=to_addrs)


def clean_str(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def split_emails(s: str) -> list[str]:
    s = clean_str(s)
    if not s:
        return []
    partes = re.split(r"[;,]", s)
    return [e.strip() for e in partes if e.strip()]


def safe_filename(text: str) -> str:
    text = str(text).strip()
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text)
    return text


def norm_key(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower() if s else ""


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
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
    # fallback por si lo ejecutan manual sin Excel abierto
    app = xw.App(visible=False)

try:
    wb = app.books[EXCEL_NAME]
except Exception:
    # fallback: abrirlo por ruta si no está abierto con ese nombre
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"❌ No se encontró el archivo Excel: {EXCEL_PATH}")
    wb = app.books.open(str(EXCEL_PATH))

sht_clusters = wb.sheets[SHEET_CLUSTERS]
sht_param = wb.sheets[SHEET_PARAMETROS]


# =====================================================
# LEER PARÁMETROS MAIL (VENTAS)
# =====================================================
MAIL_FROM = SMTP_USER
MAIL_TO = split_emails(sht_param.range("B15").value)
MAIL_CC = split_emails(sht_param.range("B16").value)
ASUNTO = clean_str(sht_param.range("B17").value)
CUERPO = clean_str(sht_param.range("B18").value)


# =====================================================
# DETECTAR CLUSTERS
# =====================================================
fila_cluster = 2
ultima_col = sht_clusters.used_range.last_cell.column

clusters: list[dict] = []
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
        clusters.append(
            {
                "nombre": str(val).replace("CLUSTER", "").strip(),
                "col_inicio": col_inicio,
                "col_fin": col_fin,
            }
        )
        col = col_fin + 1
    else:
        col += 1


# =====================================================
# LECTURA TABLAS
# =====================================================
def buscar_fila_encabezado(sht, col_inicio: int, max_filas: int = 300) -> int | None:
    for fila in range(1, max_filas + 1):
        val = sht.range((fila, col_inicio)).value
        if val and "rubro" in norm_key(val):
            return fila
    return None


def leer_tabla_cluster(sht, col_inicio: int, col_fin: int, fila_header: int) -> pd.DataFrame:
    rango = sht.range((fila_header, col_inicio), (sht.used_range.last_cell.row, col_fin))
    df = rango.options(pd.DataFrame, header=1, index=False).value
    return df.dropna(how="all")


# =====================================================
# EXPORTAR EXCEL – VENTAS
# =====================================================
def exportar_excel_ventas_xlwings(df: pd.DataFrame, base_path_excel: Path, nombre_cluster: str) -> list[Path]:
    RUBROS = {
        "Carnes": "CARNES",
        "Fiambres": "FIAMBRES",
        "Dira": "DIRA",
    }

    A1_POR_CLUSTER = {
        "CASILDA": "PubCAS",
        "RAFAELA": "PubRAF",
        "MARIA LUISA": "PUBMAR",
        "MAYORISTAS": "MayRet",
        "HORECA": "MHOR",
    }

    col_rubro = find_col(df, ["Rubro"])
    col_codigo = find_col(df, ["Cód.", "Cód", "Cod", "Código", "COD_ART"])
    col_precio = find_col(df, ["sin iva", "SIN IVA", "SIN_IVA", "Sin Iva", "Sin iva"])

    if not col_rubro or not col_codigo or not col_precio:
        return []

    df[col_rubro] = df[col_rubro].astype(str).str.upper().str.strip()

    archivos_generados: list[Path] = []
    cluster_norm = nombre_cluster.upper().replace("CLUSTER", "").strip()

    for nombre_rubro, rubro in RUBROS.items():
        df_r = df[df[col_rubro] == rubro].copy()
        if df_r.empty:
            continue

        # precio numérico + filtro != 0
        df_r[col_precio] = pd.to_numeric(df_r[col_precio], errors="coerce")
        df_r = df_r[df_r[col_precio].notna() & (df_r[col_precio] != 0)]
        if df_r.empty:
            continue

        df_out = pd.DataFrame(
            {
                "Codigo": df_r[col_codigo],
                "Precio": df_r[col_precio],
                "Cero": 0.00,
                "Fecha": FECHA_PROCESO,  # ✅ FECHA DESDE EXCEL (VBA)
            }
        )

        # ===== NUEVO EXCEL POR RUBRO =====
        app_tmp = xw.App(visible=False)
        app_tmp.api.DisplayAlerts = False

        try:
            wb_tmp = app_tmp.books.add()
            sht = wb_tmp.sheets[0]
            sht.name = nombre_rubro

            sht.range("A1").value = A1_POR_CLUSTER.get(cluster_norm, "")
            sht.range("A2").options(index=False, header=False).value = df_out

            last_row = 1 + len(df_out)
            if last_row >= 2:
                sht.range((2, 2), (last_row, 2)).api.NumberFormat = "#.##0,00"
                sht.range((2, 3), (last_row, 3)).api.NumberFormat = "0,00"
                sht.range((2, 4), (last_row, 4)).api.NumberFormat = "dd/mm/aaaa"

            sht.autofit()

            path_excel = base_path_excel.with_name(
                f"{base_path_excel.stem}_{nombre_rubro}{base_path_excel.suffix}"
            )

            if path_excel.exists():
                path_excel.unlink()

            wb_tmp.save(str(path_excel))
            archivos_generados.append(path_excel)

        finally:
            try:
                wb_tmp.close()
            except Exception:
                pass
            app_tmp.quit()

    return archivos_generados


# =====================================================
# PROCESO PRINCIPAL
# =====================================================

# (opcional) limpiar adjuntos previos de la misma fecha para no mezclar ejecuciones
for p in TEMP_DIR.glob(f"Ventas_*_{fecha_arch}.xlsx"):
    try:
        p.unlink()
    except Exception:
        pass

for p in TEMP_DIR.glob(f"Ventas_*_{fecha_arch}_*.xlsx"):
    # por si cambió la convención de nombres, lo dejamos
    try:
        p.unlink()
    except Exception:
        pass

adjuntos_generados: list[Path] = []

for c in clusters:
    fila_header = buscar_fila_encabezado(sht_clusters, c["col_inicio"])
    if not fila_header:
        continue

    df_cluster = leer_tabla_cluster(sht_clusters, c["col_inicio"], c["col_fin"], fila_header)
    if df_cluster.empty:
        continue

    nombre_limpio = safe_filename(c["nombre"])
    path_excel_base = TEMP_DIR / f"Ventas_{nombre_limpio}_{fecha_arch}.xlsx"

    generados = exportar_excel_ventas_xlwings(df_cluster, path_excel_base, nombre_limpio)
    adjuntos_generados.extend(generados)


# =====================================================
# ENVÍO DE UN SOLO MAIL (ADJUNTOS DE LA FECHA PROCESO)
# =====================================================
adjuntos = [
    p for p in TEMP_DIR.glob("Ventas_*.xlsx")
    if fecha_arch in p.name
]

if MAIL_TO and adjuntos:
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(MAIL_TO)
    if MAIL_CC:
        msg["Cc"] = ", ".join(MAIL_CC)

    # (opcional) agregar fecha al asunto
    msg["Subject"] = f"{ASUNTO or 'Precios de Venta'} - {fecha_excel_str}"
    msg.set_content(CUERPO)

    destinatarios = MAIL_TO + MAIL_CC

    for path in adjuntos:
        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=path.name,
            )

    try:
        enviar_mail_o365(msg, MAIL_FROM, destinatarios)
        print(f"✅ Mail enviado OK → {', '.join(destinatarios)}")
        print(f"📅 Fecha proceso: {fecha_excel_str}")
        print("📎 Adjuntos:")
        for a in adjuntos:
            print(f"   - {a.name}")
    except Exception as e:
        print(f"❌ Error enviando mail: {e}")

else:
    if not MAIL_TO:
        print("❌ No hay destinatarios en MAIL_TO (Parametros!B15).")
    if not adjuntos:
        print(f"❌ No se encontraron adjuntos para la fecha {fecha_arch} en {TEMP_DIR}.")
