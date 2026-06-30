import pyodbc
import pandas as pd
import xlwings as xw
import numpy as np
import xlwings.constants as xwc
import sys
from pathlib import Path

# ================================
#   PARÁMETROS DESDE VBA
# ================================
if len(sys.argv) > 1 and sys.argv[1].isdigit():
    CANTIDAD_SEMANAS = int(sys.argv[1])
else:
    CANTIDAD_SEMANAS = 4

if len(sys.argv) > 2 and sys.argv[2].isdigit():    
    ID_LISTA = int(sys.argv[2])    
else:
    ID_LISTA = 2   # fallback por defecto
    print("⚠️ ID_LISTA no recibido, se usa valor por defecto: 16")


# ================================
#   RUTA DEL ARCHIVO DE ESTADO
# ================================

BASE_DIR = Path(__file__).resolve().parent          # ...\scripts
APP_DIR = BASE_DIR.parent                           # C:\Apps\Seguimiento de Precios

estado_dir = APP_DIR / "logs"
estado_dir.mkdir(parents=True, exist_ok=True)

estado_path = estado_dir / "estado_ListasPrecios.txt"    

# ================================
#   MAPEO ID_LISTA → HOJA EXCEL
# ================================
MAPA_HOJAS = {
    126: "Lista Consumo Interno",
    16: "Listas anteriores R",
    23: "Listas anteriores ML",
    17: "Listas anteriores C",
    28: "Listas anteriores May.",
    29: "Listas anteriores HORECA",
    2: "Listas anteriores ER",
    164: "Lista Base Porcino"
}

if ID_LISTA not in MAPA_HOJAS:
    raise Exception(f"IdLista no reconocido: {ID_LISTA}")

SHEET_NAME = MAPA_HOJAS[ID_LISTA]

try:
    # ================================
    #   1. CONECTAR A SQL SERVER
    # ================================
    conn = pyodbc.connect(
        "Driver={SQL Server};"
        "Server=lariosql70;"
        "Database=DW;"
        "UID=sa;"
        "PWD=sqlmanager;"
    )

    # ================================
    #   2. EJECUTAR STORED PROCEDURE
    # ================================
    df = pd.read_sql(
        "EXEC DW.dbo.SP_CONS_ListarPreciosUltimasSemanas ?, ?",
        conn,
        params=[CANTIDAD_SEMANAS, ID_LISTA]
    )



    if df.empty:
        raise Exception("El SP no devolvió datos")

    if df["id_lista"].nunique() != 1:
        raise Exception("El SP devolvió más de una lista distinta")

    id_lista_real = df["id_lista"].iloc[0]
    nombre_lista = f"IdLista {id_lista_real}"


    # ================================
    #   3. PREPARAR DATOS
    # ================================
    df["FechaSemana_dt"] = pd.to_datetime(df["FechaSemana"], errors="coerce")

    if "UNIDAD_ME" not in df.columns:
        raise Exception("El SP no devuelve UNIDAD_ME")

    info_por_art = (
        df[["Articulo", "DESCRI_AR", "GRAN_RUBRO_CDG", "DESCRIPCION_TIPO_ART", "UNIDAD_ME"]]
        .drop_duplicates("Articulo")
        .set_index("Articulo")
    )


    # ================================
    #   4. PIVOT POR SEMANA
    # ================================

    df_pivot = df.pivot_table(
        index="Articulo",
        columns="FechaSemana_dt",
        values="precio",
        aggfunc="first"
    )


    df_pivot = df_pivot.reindex(
        sorted(df_pivot.columns, reverse=True),
        axis=1
    )

    df_pivot = df_pivot.reset_index()

    df_pivot["DESCRI_AR"] = df_pivot["Articulo"].map(info_por_art["DESCRI_AR"])
    df_pivot["GRAN_RUBRO_CDG"] = df_pivot["Articulo"].map(info_por_art["GRAN_RUBRO_CDG"])
    df_pivot["DESCRIPCION_TIPO_ART"] = df_pivot["Articulo"].map(info_por_art["DESCRIPCION_TIPO_ART"])
    

    nuevas_cols = []
    for c in df_pivot.columns:
        if isinstance(c, pd.Timestamp):
            nuevas_cols.append(c.strftime("%d-%b").lower())
        else:
            nuevas_cols.append(c)
    df_pivot.columns = nuevas_cols

    df_pivot["UNIDAD_ME"] = df_pivot["Articulo"].map(info_por_art["UNIDAD_ME"])

    # ================================
    #   5. VARIACIONES % INTERCALADAS
    # ================================
    cols_base = {"Articulo", "DESCRI_AR", "GRAN_RUBRO_CDG", "DESCRIPCION_TIPO_ART", "UNIDAD_ME"}
    columnas_semanas = [c for c in df_pivot.columns if c not in cols_base]

    df_pivot[columnas_semanas] = (
        df_pivot[columnas_semanas]
        .apply(pd.to_numeric, errors="coerce")
        .round(0)
    )

    nuevo_orden = ["Articulo", "DESCRI_AR", "GRAN_RUBRO_CDG", "DESCRIPCION_TIPO_ART"]

    for i, col_actual in enumerate(columnas_semanas):
        nuevo_orden.append(col_actual)

        if i < len(columnas_semanas) - 1:
            col_anterior = columnas_semanas[i + 1]
            col_var = f"%_{i}"

            df_pivot[col_var] = (
                (df_pivot[col_actual] - df_pivot[col_anterior])
                / df_pivot[col_anterior]
            )

            nuevo_orden.append(col_var)

    nuevo_orden.append("UNIDAD_ME")
    df_pivot = df_pivot[nuevo_orden]

    for col in df_pivot.columns:
        if isinstance(col, str) and col.startswith("%_"):
            df_pivot[col] = df_pivot[col].apply(
                lambda x: "" if pd.isna(x) else f"{x:.3%}"
            )

    df_pivot.columns = [
        "%" if isinstance(c, str) and c.startswith("%_") else c
        for c in df_pivot.columns
    ]

    for col in df_pivot.columns:
        if col != "UNIDAD_ME":
            df_pivot[col] = df_pivot[col].replace({np.nan: ""})
            


    # ================================
    #   6. ESCRIBIR EN EXCEL
    # ================================
    app = xw.apps.active
    wb = app.books["WFW_SPC.xlsm"]
    sht = wb.sheets[SHEET_NAME]

    sht.clear()

    # Encabezado
    sht.range("A1").value = f"Lista: {nombre_lista} (Id {id_lista_real})"
    sht.range("A2").value = f"Últimas {CANTIDAD_SEMANAS} semanas"
    sht.range("A1:A2").api.Font.Bold = True

    # Tabla SIN índice
    sht.range("A3").options(index=False).value = df_pivot

    rng = sht.range("A3").current_region
    last_row = rng.last_cell.row
    last_col = rng.last_cell.column

    header_range = sht.range((3, 1), (3, last_col))
    header_range.api.Font.Bold = True
    header_range.api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter
    header_range.api.VerticalAlignment = xwc.VAlign.xlVAlignCenter

    sht.range((3, 1), (last_row, last_col)).api.EntireColumn.AutoFit()

    # ================================
    #   7. PINTAR UNIDAD_ME = 1-Unit(s) (CELESTE)
    # ================================
    headers = sht.range((3, 1), (3, last_col)).value
    headers_norm = [str(h).strip().upper() if h else "" for h in headers]

    if "UNIDAD_ME" not in headers_norm:
        raise Exception("No se encontró UNIDAD_ME en Excel")

    col_unidad = headers_norm.index("UNIDAD_ME") + 1

    # Limpiar colores SOLO datos
    sht.range((4, 1), (last_row, last_col)).api.Interior.Pattern = xwc.Constants.xlNone

    # Color celeste #C0E6F5
    color_celeste = 192 + (230 << 8) + (245 << 16)

    for r in range(4, last_row + 1):
        cell = sht.range((r, col_unidad)).api
        txt = str(cell.Text).strip().upper()

        if txt == "":
            continue

        # 👉 CONDICIÓN REAL DE LISTAS
        if txt.startswith("1-"):
            # pintar solo columnas descriptivas
            sht.range((r, 1), (r, 4)).api.Interior.Color = color_celeste

    # ================================
    #   8. PINTAR VARIACIONES %
    # ================================
    COLOR_AMARILLO = 255 + (255 << 8) + (153 << 16)
    COLOR_ROJO     = 255 + (199 << 8) + (206 << 16)

    cols_pct = [i + 1 for i, h in enumerate(headers) if h == "%"]

    for c in cols_pct:
        for r in range(4, last_row + 1):
            cell = sht.range((r, c)).api
            txt = str(cell.Text).strip()

            if txt == "":
                continue

            txt = txt.replace("%", "").replace(",", ".")
            try:
                val = float(txt)
            except:
                continue

            if val > 0:
                cell.Interior.Color = COLOR_AMARILLO
            elif val < 0:
                cell.Interior.Color = COLOR_ROJO

    # ================================
    #   OCULTAR COLUMNA UNIDAD_ME (FINAL)
    # ================================
    headers = sht.range((3, 1), (3, last_col)).value
    headers_norm = [str(h).strip().upper() if h else "" for h in headers]

    if "UNIDAD_ME" in headers_norm:
        col_unidad = headers_norm.index("UNIDAD_ME") + 1
        sht.range((3, col_unidad), (last_row, col_unidad)).api.EntireColumn.Hidden = True


    wb.save()

finally:
    try:
        conn.close()
    except:
        pass
