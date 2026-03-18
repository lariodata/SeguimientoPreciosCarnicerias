import sys
import pandas as pd
import pyodbc
import shutil
from datetime import datetime
from pathlib import Path
from decimal import Decimal

# =====================================================
# RUTA BASE
# =====================================================

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent

# =====================================================
# ARCHIVO EXCEL (VBA o Default local)
# =====================================================

if len(sys.argv) >= 2:
    archivo_excel = Path(sys.argv[1]).resolve()
    print("📂 Archivo recibido desde VBA:", archivo_excel)
else:
    archivo_excel = BASE_DIR / "Template_Competencias.xlsx"
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

    PROJECT_DIR = BASE_DIR.parent
    TEMP_DIR = PROJECT_DIR / "temp"
    TEMP_DIR.mkdir(exist_ok=True)

    archivo_tmp = TEMP_DIR / f"Template_Competencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

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

        # Leer fila de fechas (fila 2 Excel = índice 1) y encabezados (fila 3 = índice 2)
        df_raw = pd.read_excel(
            archivo_tmp,
            sheet_name=hoja,
            header=None,
            nrows=3,
            engine="openpyxl"
        )
        date_row    = df_raw.iloc[1]
        header_row  = df_raw.iloc[2]

        # Mapa: nombre_competencia -> fecha_modificacion
        fecha_map = {}
        for col_idx, col_name in enumerate(header_row):
            if pd.notna(col_name) and str(col_name).strip() not in ("Código", "Descripción"):
                fecha_val = date_row.iloc[col_idx]
                if pd.notna(fecha_val):
                    if isinstance(fecha_val, datetime):
                        fecha_map[str(col_name).strip()] = fecha_val.date()
                    else:
                        try:
                            fecha_map[str(col_name).strip()] = pd.to_datetime(fecha_val, dayfirst=True).date()
                        except Exception:
                            fecha_map[str(col_name).strip()] = None

        df = pd.read_excel(
            archivo_tmp,
            sheet_name=hoja,
            header=2,
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
            fecha_mod = fecha_map.get(str(row["competencia"]).strip())
            # Convertir date a string ISO para SQL
            if fecha_mod:
                fecha_mod = fecha_mod.isoformat()

            inserts.append((
                int(row["cod_art"]),
                hoja,
                row["competencia"],
                precio,
                anio,
                semana,
                fecha_mod
            ))

    # -------------------------------------------------
    # INSERT MASIVO
    # -------------------------------------------------

    if inserts:
        df_insert = pd.DataFrame(inserts, columns=[
            "cod_art", "cluster", "competencia", "precio", "anio", "semana", "fecha_modificacion"
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
            (cod_art, cluster, competencia, precio, anio, semana, fecha_modificacion)
            VALUES (?, ?, ?, ?, ?, ?, ?)
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
        archivo_tmp.unlink()
    except:
        pass

    conn.close()