import sys
import pandas as pd
import pyodbc
import shutil
import os
from datetime import datetime
from pathlib import Path
import sys

# ================================
#   ARCHIVO EXCEL RECIBIDO DESDE VBA
# ================================

if len(sys.argv) < 2:
    raise ValueError("No se recibió la ruta del archivo Excel desde VBA")

archivo_excel = Path(sys.argv[1]).resolve()

if not archivo_excel.exists():
    raise FileNotFoundError(f"No existe el archivo Excel: {archivo_excel}")

print("📂 Archivo de competencias:", archivo_excel)

HOJAS_VALIDAS = ["Rafaela", "Maria Luisa", "Casilda", "Mayorista"]


# ================================
# FECHA DE NEGOCIO
# ================================
fecha_hoy = datetime.now()
anio = fecha_hoy.year
semana = fecha_hoy.isocalendar()[1]  # semana ISO

# ================================
# CONEXIÓN SQL SERVER
# ================================
conn = pyodbc.connect(
    "Driver={SQL Server};"
    "Server=lariosql70;"
    "Database=DW;"
    "UID=sa;"
    "PWD=sqlmanager;"
)
cursor = conn.cursor()

try:
    print(f"Procesando semana {semana} / {anio}")

    # ================================
    # TRANSACCIÓN
    # ================================
    conn.autocommit = False

    # 1️⃣ BORRAR SEMANA ACTUAL
    cursor.execute("""
        DELETE FROM dbo.Lista_Precios_Competencia
        WHERE anio = ? AND semana = ?
    """, anio, semana)

    print("Semana actual eliminada")

    # ================================
    # LECTURA EXCEL
    # ================================


    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # Subir 1 nivel: scripts -> 00 - Workflow
    PROJECT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))

    TEMP_DIR = os.path.join(PROJECT_DIR, "temp")
    os.makedirs(TEMP_DIR, exist_ok=True)

    archivo_tmp = os.path.join(
        TEMP_DIR,
        f"Template_Competencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    shutil.copy2(archivo_excel, archivo_tmp)

    print("Archivo copiado a:", archivo_tmp)



    xls = pd.ExcelFile(archivo_tmp)

    inserts = []

    for hoja in HOJAS_VALIDAS:
        if hoja not in xls.sheet_names:
            continue

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

        df = df[df["cod_art"].notna()]

        # Columnas de competencias: desde la columna D en adelante
        columnas_competencias = df.columns[3:]

        # Eliminar columnas totalmente vacías
        columnas_competencias = [
            c for c in columnas_competencias
            if not df[c].isna().all()
        ]


        df_melt = df.melt(
            id_vars=["cod_art"],
            value_vars=columnas_competencias,
            var_name="competencia",
            value_name="precio"
        )

        for _, row in df_melt.iterrows():
            valor = row["precio"]

            if pd.isna(valor):
                precio = None
            else:
                try:
                    precio = float(str(valor).replace(",", "."))
                    if precio == 0:
                        precio = None
                except (ValueError, TypeError):
                    precio = None



            inserts.append((
                int(row["cod_art"]),
                hoja,
                row["competencia"],
                precio,
                anio,
                semana
            ))

    # ================================
    # INSERT EN BLOQUE
    # ================================
    cursor.executemany("""
        INSERT INTO dbo.Lista_Precios_Competencia
        (cod_art, cluster, competencia, precio, anio, semana)
        VALUES (?, ?, ?, ?, ?, ?)
    """, inserts)

    conn.commit()
    print(f"Carga finalizada correctamente. Registros insertados: {len(inserts)}")

except Exception as e:
    conn.rollback()
    print("❌ Error durante la carga, se hace ROLLBACK")
    raise e

finally:
    xls.close()
    os.remove(archivo_tmp)
    conn.close()

