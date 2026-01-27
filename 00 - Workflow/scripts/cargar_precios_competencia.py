import pyodbc
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime
import xlwings.constants as xwc

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
    #    (SIEMPRE ÚLTIMA SEMANA)
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

    df_comp = df[
        columnas_fijas + ["cluster", "competencia", "precio"]
    ].copy()

    # ================================
    # 4. ORDEN CLUSTER → COMPETENCIA
    # ================================
    orden_cols = (
        df_comp[["cluster", "competencia"]]
        .drop_duplicates()
        .sort_values(["cluster", "competencia"])
    )

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
    df_pivot["GRAN_RUBRO_CDG"] = df_pivot["GRAN_RUBRO_CDG"] = (
        df_pivot["cod_art"].map(info_art["GRAN_RUBRO_CDG"])
    )
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
    # 10. FILA 4 → CLUSTERS (E en adelante)
    # ================================
    col_ini = len(columnas_fijas) + 1  # E
    i = 0
    col = col_ini

    while i < len(clusters):
        j = i
        while j < len(clusters) and clusters[j] == clusters[i]:
            j += 1

        sht.range((4, col)).value = clusters[i]
        sht.range((4, col)).api.Font.Bold = True
        sht.range((4, col)).api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter

        if j - i > 1:
            sht.range((4, col), (4, col + (j - i) - 1)).api.Merge()

        col += (j - i)
        i = j

    # ================================
    # 11. FILA 5 → ENCABEZADOS REALES
    # ================================
    encabezados = columnas_fijas + competencias

    sht.range("A5").value = encabezados
    sht.range("A5").api.Font.Bold = True
    sht.range("A5").api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter

    # ================================
    # 12. DATOS DESDE FILA 6
    # ================================
    sht.range("A5").options(index=False).value = df_pivot

    from itertools import cycle

    paleta = cycle(PALETA_COLORES)

    col_ini = len(columnas_fijas) + 1  # columna E
    i = 0
    col = col_ini

    while i < len(clusters):
        j = i
        while j < len(clusters) and clusters[j] == clusters[i]:
            j += 1

        color = next(paleta)

        col_inicio = col
        col_fin = col + (j - i) - 1

        # ---------
        # FILA 4: CLUSTER
        # ---------
        sht.range((4, col_inicio)).value = clusters[i]
        sht.range((4, col_inicio)).api.Font.Bold = True
        sht.range((4, col_inicio)).api.HorizontalAlignment = xwc.HAlign.xlHAlignCenter
        sht.range((4, col_inicio)).api.Interior.Color = color

        if col_fin > col_inicio:
            sht.range((4, col_inicio), (4, col_fin)).api.Merge()
            sht.range((4, col_inicio), (4, col_fin)).api.Interior.Color = color

        # ---------
        # FILA 5: COMPETENCIAS
        # ---------
        sht.range((5, col_inicio), (5, col_fin)).api.Interior.Color = color

        # ---------
        # BORDES (cluster + competencias)
        # ---------
        bloque = sht.range((4, col_inicio), (5, col_fin)).api

        for borde in [
            xwc.BordersIndex.xlEdgeLeft,
            xwc.BordersIndex.xlEdgeRight,
            xwc.BordersIndex.xlEdgeTop,
            xwc.BordersIndex.xlEdgeBottom
        ]:
            bloque.Borders(borde).LineStyle = 1  # xlContinuous
            bloque.Borders(borde).Weight = 2     # xlThin

        col += (j - i)
        i = j

    # ================================
    # COLOR GRIS PARA COLUMNAS FIJAS (ENCABEZADOS)
    # ================================
    GRIS_CLARO = 242 + (242 << 8) + (242 << 16)

    for col in range(1, len(columnas_fijas) + 1):
        celda = sht.range((5, col)).api
        celda.Interior.Color = GRIS_CLARO
        celda.Font.Color = 0  # negro
        celda.Font.Bold = True

        # borde fino
        for borde in [
            xwc.BordersIndex.xlEdgeLeft,
            xwc.BordersIndex.xlEdgeRight,
            xwc.BordersIndex.xlEdgeTop,
            xwc.BordersIndex.xlEdgeBottom
        ]:
            celda.Borders(borde).LineStyle = 1  # xlContinuous
            celda.Borders(borde).Weight = 2     # xlThin



    # ================================
    # 13. AJUSTES FINALES
    # ================================
    sht.range("A5").current_region.api.EntireColumn.AutoFit()

    wb.save()

    print(
        f"Competencias actualizadas correctamente - "
        f"{datetime.now():%d/%m/%Y %H:%M:%S}"
    )

except Exception as e:
    print("ERROR:", e)
    raise

finally:
    try:
        conn.close()
    except:
        pass
