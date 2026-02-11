import pyodbc
import pandas as pd
import xlwings as xw
import os
import xlwings.constants as xwc
import sys
import numpy as np
from datetime import datetime
from pathlib import Path

# ================================
#   PARÁMETRO DESDE VBA
# ================================

if len(sys.argv) > 1 and sys.argv[1].isdigit():
    CANTIDAD_SEMANAS = int(sys.argv[1])
else:
    CANTIDAD_SEMANAS = 4   # fallback seguro


from pathlib import Path

# ================================
#   RUTA DEL ARCHIVO DE ESTADO
# ================================

BASE_DIR = Path(__file__).resolve().parent          # ...\scripts
APP_DIR = BASE_DIR.parent                           # C:\Apps\Seguimiento de Precios

estado_dir = APP_DIR / "logs"
estado_dir.mkdir(parents=True, exist_ok=True)

estado_path = estado_dir / "estado_costos.txt"


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
        "EXEC SP_CONS_ListarCostosUltimasSemanas ?",
        conn,
        params=[CANTIDAD_SEMANAS]
    )

    if df.empty:
        raise Exception("El SP no devolvió datos")

    # ================================
    #   3. PREPARAR DATOS
    # ================================
    df["FechaSemana_dt"] = pd.to_datetime(df["FechaSemana"], errors="coerce")

    print("Filas totales:", len(df))
    print("Artículos únicos:", df["Articulo"].nunique())

    df_ccu = df[
        [
            "Articulo",
            "DESCRI_AR",
            "GRAN_RUBRO_CDG",
            "DESCRIPCION_TIPO_ART",
            "UNIDAD_ME",
            "FechaSemana_dt",
            "CostoUnitario",
        ]
    ].copy()

    # ================================
    #   INFO MAESTRA POR ARTÍCULO
    # ================================
    info_art = (
        df_ccu[["Articulo", "DESCRI_AR", "GRAN_RUBRO_CDG", "DESCRIPCION_TIPO_ART","UNIDAD_ME"]]
        .drop_duplicates(subset=["Articulo"])
        .set_index("Articulo")
    )

    # ================================
    #   4. PIVOT POR SEMANA (CLAVE CORRECTA)
    # ================================
    df_pivot = df_ccu.pivot_table(
        index="Articulo",
        columns="FechaSemana_dt",
        values="CostoUnitario",
        aggfunc="first"
    )

    # Ordenar semanas: más reciente → más antigua
    df_pivot = df_pivot.reindex(
        sorted(df_pivot.columns, reverse=True),
        axis=1
    )

    df_pivot = df_pivot.reset_index()

    # ================================
    #   REINCORPORAR ATRIBUTOS
    # ================================
    df_pivot["DESCRI_AR"] = df_pivot["Articulo"].map(info_art["DESCRI_AR"])
    df_pivot["GRAN_RUBRO_CDG"] = df_pivot["Articulo"].map(info_art["GRAN_RUBRO_CDG"])
    df_pivot["DESCRIPCION_TIPO_ART"] = df_pivot["Articulo"].map(info_art["DESCRIPCION_TIPO_ART"])
    df_pivot["UNIDAD_ME"] = df_pivot["Articulo"].map(info_art["UNIDAD_ME"])

    # ================================
    #   RENOMBRAR COLUMNAS FECHA
    # ================================
    nuevas_cols = []
    for c in df_pivot.columns:
        if isinstance(c, pd.Timestamp):
            nuevas_cols.append(c.strftime("%d-%b").lower())
        else:
            nuevas_cols.append(c)
    df_pivot.columns = nuevas_cols

    # ================================
    #   5. VARIACIONES % INTERCALADAS
    # ================================
    cols_base = {"Articulo", "DESCRI_AR", "GRAN_RUBRO_CDG", "DESCRIPCION_TIPO_ART", "UNIDAD_ME"}
    columnas_semanas = [c for c in df_pivot.columns if c not in cols_base]

    df_pivot[columnas_semanas] = (
        df_pivot[columnas_semanas]
        .replace("", np.nan)
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

    # Formatear columnas %
    for col in df_pivot.columns:
        if isinstance(col, str) and col.startswith("%_"):
            df_pivot[col] = df_pivot[col].apply(
                lambda x: "" if pd.isna(x) else f"{x:.3%}"
            )

    # Renombrar %_* → %
    df_pivot.columns = [
        "%" if isinstance(c, str) and c.startswith("%_") else c
        for c in df_pivot.columns
    ]

    # NaN → vacío (excepto UNIDAD_ME)
    for col in df_pivot.columns:
        if col != "UNIDAD_ME":
            df_pivot[col] = df_pivot[col].replace({np.nan: ""})

    # ================================
    #   6. ESCRIBIR EN EXCEL
    # ================================
    app = xw.apps.active
    wb = app.books["WFW_SPC.xlsm"]

    if "Costos Unitarios" in [s.name for s in wb.sheets]:
        sht = wb.sheets["Costos Unitarios"]
        sht.clear()
    else:
        sht = wb.sheets.add("Costos Unitarios")

    # Encabezado
    sht.range("A1").value = "Costos Unitarios"
    sht.range("A2").value = f"Últimas {CANTIDAD_SEMANAS} semanas"
    sht.range("A1:A2").api.Font.Bold = True

    sht.range("A3").options(index=False).value = df_pivot

    rng = sht.range("A3").current_region
    last_row = rng.last_cell.row
    last_col = rng.last_cell.column

    header_range = sht.range((3, 1), (3, last_col))
    header_range.api.Font.Bold = True
    header_range.api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter
    header_range.api.VerticalAlignment = xwc.VAlign.xlVAlignCenter

    sht.range((1, 1), (last_row, last_col)).api.EntireColumn.AutoFit()

    # ================================
    #   7. PINTAR UNIDAD_ME = 1
    # ================================
    headers = sht.range((3, 1), (3, last_col)).value
    headers_norm = [str(h).strip().upper() if h else "" for h in headers]

    if "UNIDAD_ME" in headers_norm:
        col_unidad = headers_norm.index("UNIDAD_ME") + 1

        sht.range((4, 1), (last_row, last_col)).api.Interior.Pattern = xwc.Constants.xlNone

        color_celeste = 192 + (230 << 8) + (245 << 16)

        for r in range(4, last_row + 1):
            try:
                if int(float(sht.range((r, col_unidad)).value)) == 1:
                    sht.range((r, 1), (r, 4)).api.Interior.Color = color_celeste
            except:
                pass

        # Ocultar columna UNIDAD_ME
        sht.range((3, col_unidad), (last_row, col_unidad)).api.EntireColumn.Hidden = True



    # ================================
    #   8. PINTAR VARIACIONES % (FILL)
    # ================================

    COLOR_AMARILLO = 255 + (255 << 8) + (153 << 16)  # #FFFF99
    COLOR_ROJO     = 255 + (199 << 8) + (206 << 16)  # #FFC7CE

    # Leer headers reales (fila 3)
    headers = sht.range((3, 1), (3, last_col)).value
    headers_norm = [str(h).strip() if h else "" for h in headers]

    # Columnas cuyo encabezado es "%"
    cols_pct = [i + 1 for i, h in enumerate(headers_norm) if h == "%"]

    for c in cols_pct:
        for r in range(4, last_row + 1):
            cell = sht.range((r, c)).api
            txt = str(cell.Text).strip()  # lo que se VE en Excel

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
                         

    wb.save()

    with open(estado_path, "w") as f:
        f.write(
            f"Completado! Costos Unitarios cargados correctamente - "
            f"{datetime.now():%d/%m/%Y %H:%M:%S}"
        )

except Exception as e:
    with open(estado_path, "w") as f:
        f.write(f"ERROR: {str(e)}")
    print("ERROR:", e)

finally:
    try:
        conn.close()
    except:
        pass
