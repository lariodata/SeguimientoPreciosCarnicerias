import pyodbc
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime
import xlwings.constants as xwc
from pathlib import Path

# ================================
#   RUTAS DEL PROYECTO / LOGS
# ================================

SCRIPT_DIR = Path(__file__).resolve().parent          # ...\scripts
PROJECT_DIR = SCRIPT_DIR.parent                      # C:\Apps\Seguimiento de Precios

logs_dir = PROJECT_DIR / "logs"
logs_dir.mkdir(exist_ok=True)

estado_path = logs_dir / "estado_Competencias.txt"

# ================================
# PALETA DE COLORES PASTEL (RGB Excel)
# ================================

PALETA_COLORES = [
    255 + (242 << 8) + (204 << 16),  # Amarillo pastel
    217 + (234 << 8) + (211 << 16),  # Verde pastel
    221 + (235 << 8) + (247 << 16),  # Celeste pastel
    244 + (204 << 8) + (204 << 16),  # Rosado pastel
    234 + (209 << 8) + (220 << 16),  # Lila pastel
    226 + (239 << 8) + (218 << 16),  # Menta pastel
]

# ================================
# 1. CONEXIÓN SQL SERVER
# ================================

conn = pyodbc.connect(
    "Driver={SQL Server};"
    "Server=lariosql70;"
    "Database=DW;"
    "UID=sa;"
    "PWD=sqlmanager;"
)

try:
    # ================================
    # 2. EJECUTAR STORED PROCEDURE
    # ================================
    df = pd.read_sql(
        "EXEC dbo.SP_CONS_ListarPreciosCompetenciaSemanal 1",
        conn
    )

    if df.empty:
        raise Exception("El SP no devolvió datos")

    # ================================
    # 3. COLUMNAS FIJAS
    # ================================
    columnas_fijas = [
        "cod_art",
        "DESCRI_AR",
        "GRAN_RUBRO_CDG",
        "DESCRIPCION_TIPO_ART"
    ]

    df_comp = df[columnas_fijas + ["cluster", "competencia", "precio"]].copy()

    # ================================
    # 4. ORDEN CLUSTER → COMPETENCIA
    # ================================


    orden_personalizado = [
        "Rafaela",
        "Casilda",
        "Maria Luisa",
        "Mayorista",
        "Estancia Rafaela"
    ]

    orden_cols = (
        df_comp[["cluster", "competencia"]]
            .drop_duplicates()
    )

    # Crear índice de orden basado en la lista
    orden_cols["orden_cluster"] = orden_cols["cluster"].apply(
        lambda x: orden_personalizado.index(x)
        if x in orden_personalizado else 999
    )

    orden_cols = orden_cols.sort_values(
        ["orden_cluster", "competencia"]
    )

    clusters = orden_cols["cluster"].tolist()
    competencias = orden_cols["competencia"].tolist()

    clusters = orden_cols["cluster"].tolist()
    competencias = orden_cols["competencia"].tolist()

    # ================================
    # 5. INFO MAESTRA POR ARTÍCULO
    # ================================
    info_art = (
        df_comp[columnas_fijas]
        .drop_duplicates(subset=["cod_art"])
        .set_index("cod_art")
    )

    # ================================
    # 6. PIVOT POR COMPETENCIA
    # ================================
    df_pivot = df_comp.pivot_table(
        index="cod_art",
        columns="competencia",
        values="precio",
        aggfunc="first"
    )

    df_pivot = df_pivot.reindex(columns=competencias).reset_index()

    # ================================
    # 7. REINCORPORAR COLUMNAS FIJAS
    # ================================
    df_pivot["DESCRI_AR"] = df_pivot["cod_art"].map(info_art["DESCRI_AR"])
    df_pivot["GRAN_RUBRO_CDG"] = df_pivot["cod_art"].map(info_art["GRAN_RUBRO_CDG"])
    df_pivot["DESCRIPCION_TIPO_ART"] = (
        df_pivot["cod_art"].map(info_art["DESCRIPCION_TIPO_ART"])
    )

    df_pivot = df_pivot[columnas_fijas + competencias]
    df_pivot = df_pivot.replace({np.nan: ""})

    # ================================
    # 8. EXCEL
    # ================================
    app = xw.apps.active
    app.api.DisplayAlerts = False

    wb = app.books["WFW_SPC.xlsm"]

    if "Competencias" in [s.name for s in wb.sheets]:
        sht = wb.sheets["Competencias"]
        sht.clear()
    else:
        sht = wb.sheets.add("Competencias")

    sht.api.Cells.UnMerge()

    # ================================
    # 9. TÍTULO
    # ================================
    sht.range("A1").value = "Competencias"
    sht.range("A1").api.Font.Bold = True

    # ================================
    # 10. FILA 4 → CLUSTERS
    # ================================
    col_ini = len(columnas_fijas) + 1
    i = 0
    col = col_ini

    from itertools import cycle
    paleta = cycle(PALETA_COLORES)

    while i < len(clusters):
        j = i
        while j < len(clusters) and clusters[j] == clusters[i]:
            j += 1

        color = next(paleta)

        col_inicio = col
        col_fin = col + (j - i) - 1

        sht.range((4, col_inicio), (4, col_fin)).api.Merge()
        sht.range((4, col_inicio)).value = clusters[i]
        sht.range((4, col_inicio)).api.Font.Bold = True
        sht.range((4, col_inicio)).api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter
        sht.range((4, col_inicio), (4, col_fin)).api.Interior.Color = color

        sht.range((5, col_inicio), (5, col_fin)).api.Interior.Color = color

        bloque = sht.range((4, col_inicio), (5, col_fin)).api
        for borde in [
            xwc.BordersIndex.xlEdgeLeft,
            xwc.BordersIndex.xlEdgeRight,
            xwc.BordersIndex.xlEdgeTop,
            xwc.BordersIndex.xlEdgeBottom
        ]:
            bloque.Borders(borde).LineStyle = 1
            bloque.Borders(borde).Weight = 2

        col += (j - i)
        i = j

    # ================================
    # 11. ENCABEZADOS Y DATOS
    # ================================
    sht.range("A5").value = columnas_fijas + competencias
    sht.range("A5").api.Font.Bold = True
    sht.range("A5").api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter

    sht.range("A6").options(index=False, header=False).value = df_pivot


    # ================================
    # 12. FORMATO MONEDA
    # ================================
    ultima_fila = sht.range("A5").current_region.last_cell.row
    ultima_col = sht.range("A5").current_region.last_cell.column

    for c in range(col_ini, ultima_col + 1):
        sht.range((6, c), (ultima_fila, c)).api.NumberFormat = "$ #.##0"

    sht.range("A5").current_region.api.EntireColumn.AutoFit()

    wb.save()

    # ================================
    # ESTADO OK
    # ================================
    with open(estado_path, "w") as f:
        f.write(
            f"OK - Competencias actualizadas correctamente - "
            f"{datetime.now():%d/%m/%Y %H:%M:%S}"
        )

except Exception as e:
    with open(estado_path, "w") as f:
        f.write(
            f"ERROR - {str(e)} - "
            f"{datetime.now():%d/%m/%Y %H:%M:%S}"
        )
    print("ERROR:", e)
    raise

finally:
    try:
        conn.close()
    except:
        pass
