import sys
import pandas as pd
import pyodbc
import shutil
import os
from datetime import datetime
from pathlib import Path
from decimal import Decimal

# =====================================================
# ARCHIVO EXCEL (VBA o Default local)
# =====================================================

if len(sys.argv) >= 2:
    archivo_excel = Path(sys.argv[1]).resolve()
    print("📂 Archivo recibido desde VBA:", archivo_excel)
else:
    archivo_excel = Path.home() / "Downloads" / "Template_Competencias.xlsx"
    print("⚠ No se recibió archivo desde VBA. Usando ruta local:")
    print("📂", archivo_excel)

if not archivo_excel.exists():
    raise FileNotFoundError(f"No existe el archivo Excel: {archivo_excel}")

HOJAS_VALIDAS = [
    "Rafaela",
    "Maria Luisa",
    "Casilda",
    "Mayorista",
    "Estancia Rafaela"
]

# =====================================================
# FECHA
# =====================================================

fecha_hoy = datetime.now()
anio = fecha_hoy.year
semana = fecha_hoy.isocalendar()[1]

# =====================================================
# CONEXIÓN SQL
# =====================================================

conn = pyodbc.connect(
    "Driver={SQL Server};"
    "Server=lariosql70;"
    "Database=DW;"
    "UID=sa;"
    "PWD=sqlmanager;"
)
cursor = conn.cursor()

# =====================================================
# FUNCIÓN ROBUSTA DE PARSEO DE PRECIO
# =====================================================

def parse_precio(valor):
    """
    Convierte distintos formatos de precio a Decimal compatible
    con decimal(18,4)
    """

    if pd.isna(valor):
        return None

    # Si ya es numérico
    if isinstance(valor, (int, float)):
        precio = float(valor)
    else:
        try:
            valor_str = str(valor).strip()

            if valor_str == "":
                return None

            valor_str = valor_str.replace("$", "")
            valor_str = valor_str.replace(" ", "")

            # Formato argentino 1.234,56
            if "," in valor_str and "." in valor_str:
                valor_str = valor_str.replace(".", "")
                valor_str = valor_str.replace(",", ".")
            elif "," in valor_str:
                valor_str = valor_str.replace(",", ".")

            precio = float(valor_str)

        except (ValueError, TypeError):
            return None

    if precio == 0:
        return None

    # Redondear a 4 decimales (según SQL)
    precio = round(precio, 4)

    # Protección overflow decimal(18,4)
    if abs(precio) > 99999999999999.9999:
        print(f"⚠ Precio fuera de rango decimal(18,4): {precio}")
        return None

    # Convertir a Decimal exacto para SQL
    return Decimal(str(precio))


# =====================================================
# PROCESO
# =====================================================

try:
    print(f"Procesando semana {semana} / {anio}")
    conn.autocommit = False

    # -------------------------------------------------
    # BORRAR SEMANA ACTUAL
    # -------------------------------------------------

    cursor.execute("""
        DELETE FROM dbo.Lista_Precios_Competencia
        WHERE anio = ? AND semana = ?
    """, anio, semana)

    print("Semana actual eliminada")

    # -------------------------------------------------
    # COPIAR A TEMP
    # -------------------------------------------------

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    PROJECT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))
    TEMP_DIR = os.path.join(PROJECT_DIR, "temp")
    os.makedirs(TEMP_DIR, exist_ok=True)

    archivo_tmp = os.path.join(
        TEMP_DIR,
        f"Template_Competencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    shutil.copy2(archivo_excel, archivo_tmp)

    xls = pd.ExcelFile(archivo_tmp)
    inserts = []

    # -------------------------------------------------
    # PROCESAR HOJAS
    # -------------------------------------------------

    for hoja in HOJAS_VALIDAS:

        if hoja not in xls.sheet_names:
            continue

        print(f"Procesando hoja: {hoja}")

        df = pd.read_excel(
            archivo_tmp,
            sheet_name=hoja,
            header=1,
            engine="openpyxl"
        )

        df = df.rename(columns={
            "Código": "cod_art",
            "Descripción": "descripcion"
        })

        # -------------------------------
        # LIMPIEZA COD_ART
        # -------------------------------

        df["cod_art"] = pd.to_numeric(df["cod_art"], errors="coerce")
        df = df[df["cod_art"].notna()]
        df["cod_art"] = df["cod_art"].astype(int)

        if df.empty:
            print(f"⚠ Hoja {hoja} sin códigos válidos.")
            continue

        # -------------------------------
        # COLUMNAS COMPETENCIA
        # -------------------------------

        columnas_competencias = df.columns[3:]

        columnas_competencias = [
            c for c in columnas_competencias
            if df[c].notna().any()
        ]

        if not columnas_competencias:
            print(f"⚠ Hoja {hoja} sin precios cargados.")
            continue

        # -------------------------------
        # MELT
        # -------------------------------

        df_melt = df.melt(
            id_vars=["cod_art"],
            value_vars=columnas_competencias,
            var_name="competencia",
            value_name="precio"
        )

        # -------------------------------
        # LIMPIEZA PRECIOS
        # -------------------------------

        for _, row in df_melt.iterrows():

            precio = parse_precio(row["precio"])

            inserts.append((
                int(row["cod_art"]),
                hoja,
                row["competencia"],
                precio,
                anio,
                semana
            ))

    # -------------------------------------------------
    # INSERT MASIVO
    # -------------------------------------------------

    if inserts:
        df_insert = pd.DataFrame(inserts, columns=[
            "cod_art", "cluster", "competencia", "precio", "anio", "semana"
        ])

        # Eliminar duplicados conservando el último valor
        df_insert = df_insert.drop_duplicates(
            subset=["cod_art", "cluster", "competencia", "anio", "semana"],
            keep="last"
        )

        inserts = list(df_insert.itertuples(index=False, name=None))
        
    if inserts:
        cursor.executemany("""
            INSERT INTO dbo.Lista_Precios_Competencia
            (cod_art, cluster, competencia, precio, anio, semana)
            VALUES (?, ?, ?, ?, ?, ?)
        """, inserts)

        conn.commit()
        print(f"✅ Carga finalizada. Registros insertados: {len(inserts)}")
    else:
        print("⚠ No se generaron registros válidos.")
        conn.rollback()

except Exception as e:
    conn.rollback()
    print("❌ Error durante la carga. ROLLBACK ejecutado.")
    raise e

finally:
    try:
        xls.close()
        os.remove(archivo_tmp)
    except:
        pass

    conn.close()