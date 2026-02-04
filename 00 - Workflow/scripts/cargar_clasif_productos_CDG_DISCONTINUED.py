import pandas as pd
import pyodbc

# -------------------------------------------------
# Configuración
# -------------------------------------------------
EXCEL_PATH = r"C:\Users\RodrigoMarozzi\Downloads\GRUPOS_CDG.xlsx"
SHEET_NAME = "GRUPOS"

SQL_SERVER = "lariosql70"
SQL_DATABASE = "DW"

# -------------------------------------------------
# Conexión SQL Server
# -------------------------------------------------
conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={SQL_SERVER};"
    f"DATABASE={SQL_DATABASE};"
    f"UID=sa;"
    f"PWD=sqlmanager;"
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

try:
    # -------------------------------------------------
    # Leer Excel
    # -------------------------------------------------
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    # Selección explícita y orden de columnas
    df = df[[
        "COD_ART",
        "RUBRO_CDG",
        "GRAN_RUBRO_CDG",
        "Rub_comercial",
        "Rub_carnes",
        "ESTRATEGIA"
    ]]

    # Limpieza básica
    df = df.dropna(subset=["COD_ART"])
    df = df.astype(str)

    # -------------------------------------------------
    # Truncate tabla destino
    # -------------------------------------------------
    cursor.execute("TRUNCATE TABLE dbo.DimClasificacionProducto")
    conn.commit()

    # -------------------------------------------------
    # Insert masivo
    # -------------------------------------------------
    insert_sql = """
        INSERT INTO dbo.DimClasificacionProducto (
            COD_ART,
            RUBRO_CDG,
            GRAN_RUBRO_CDG,
            Rub_comercial,
            Rub_carnes,
            Estrategia
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """

    cursor.fast_executemany = True
    cursor.executemany(insert_sql, df.values.tolist())
    conn.commit()

    print(f"Carga OK - Filas insertadas: {len(df)}")

except Exception as e:
    conn.rollback()
    print("ERROR en la carga:", e)

finally:
    cursor.close()
    conn.close()
