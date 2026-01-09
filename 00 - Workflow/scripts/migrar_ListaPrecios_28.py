import pandas as pd
import pyodbc

# =====================================================
# CONFIGURACIÓN
# =====================================================
import os


import logging
from datetime import datetime
import os




BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpeta logs al mismo nivel que scripts
LOG_DIR = os.path.join(BASE_DIR, "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

# Archivo de log fijo
log_path = os.path.join(LOG_DIR, "log_lp_Cas.txt")

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
)

logging.info("===================================================")
logging.info("Inicio proceso carga Lista_Precios_Odoo - LP R")
logging.info("===================================================")


EXCEL_PATH = os.path.join(BASE_DIR, "..", "WFW_SPC.xlsm")
SHEET_NAME = "Listas anteriores May.old"

# Valores fijos de negocio
ORIGEN = "Odoo8"
ID_LISTA = 28
LISTA = "RETAIL MAYORISTA (MayRet)"
LOG_DIR = os.path.join(BASE_DIR, "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

log_path = os.path.join(
    LOG_DIR,
    f"carga_lista_precios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
)

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logging.info("Inicio proceso carga Lista_Precios_Odoo")




# =====================================================
# 1. LEER EXCEL COMPLETO (SIN HEADERS)
# =====================================================
df_raw = pd.read_excel(
    EXCEL_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

# =====================================================
# 2. ÍNDICES DE COLUMNAS DE FECHA (D, G, J, ...)
# =====================================================
col_fechas_idx = list(range(3, df_raw.shape[1], 3))

# =====================================================
# 3. FECHAS (FILA 4)
# =====================================================
fechas = pd.to_datetime(
    df_raw.iloc[3, col_fechas_idx],
    dayfirst=True,
    errors="coerce"
)

# =====================================================
# 4. DATOS BASE (FILA 5+)
# =====================================================
base = df_raw.iloc[4:, :3].copy()
base.columns = ["rubro", "cod_art", "producto"]
base["row_id"] = base.index

# =====================================================
# 5. PRECIOS (SOLO COLUMNAS DE FECHA)
# =====================================================
data = df_raw.iloc[4:, col_fechas_idx].copy()
data.index = base.index
data.columns = fechas



# =====================================================
# 5. PASAR A FORMATO LARGO (FECHA → CAMPO)
# =====================================================
df_long = data.stack().reset_index()
df_long.columns = ["row_id", "fecha", "precio"]


# =====================================================
# 6. UNIR CON DATOS BASE
# =====================================================
df_final = (
    df_long
    .merge(base, on="row_id", how="left")
    .drop(columns=["row_id"])
)

# =====================================================
# 7. LIMPIEZA Y REGLAS DE NEGOCIO
# =====================================================

# Precio:
# - NULL / vacío → -1 (producto no existe)
# - 0 → se mantiene (sin costo)
# - >0 → se mantiene

# =====================================================
# TRATAMIENTO FINAL DE PRECIO (CERRADO)
# =====================================================

# Convertir a numérico:
# - números quedan igual
# - celdas vacías → NaN
# Convertir precio a numérico
df_final["precio"] = pd.to_numeric(df_final["precio"], errors="coerce")

# ❌ Descartar filas SIN precio real
# (producto no existía para esa fecha)
df_final = df_final[
    df_final["precio"].notna() & (df_final["precio"] != "")
]




# Normalizar cod_art
df_final["cod_art"] = df_final["cod_art"].replace("", None)

# Descartar filas sin cod_art NI producto
df_final = df_final[
    ~(
        df_final["cod_art"].isna() &
        df_final["producto"].isna()
    )
]



# cod_art debe ser INT puro (no float, no string)
df_final["cod_art"] = (
    pd.to_numeric(df_final["cod_art"], errors="coerce")
    .fillna(0)
    .astype(int)
)



# Producto: texto NOT NULL
df_final["producto"] = df_final["producto"].fillna("")

# Fecha desde Excel
df_final["fecha"] = pd.to_datetime(
    df_final["fecha"],
    dayfirst=True,
    errors="coerce"
)

df_final["fecha"] = df_final["fecha"].dt.floor("min")
df_final = df_final[df_final["fecha"].notna()]


# ing_hora = misma fecha
df_final["ing_hora"] = df_final["fecha"]

# =====================================================
# 8. CAMPOS FIJOS PARA SQL (NOT NULL)
# =====================================================
df_final["origen"] = ORIGEN
df_final["id_lista"] = ID_LISTA
df_final["lista"] = LISTA

df_final["product_id"] = 0
df_final["uom"] = ""
df_final["unidad_me"] = ""
df_final["invoicing_type"] = ""
df_final["contenido"] = 0
df_final["box_kgs"] = 0
df_final["precioxkg"] = 0

# =====================================================
# 9. ORDEN EXACTO DE COLUMNAS SEGÚN SQL
# =====================================================
df_load = df_final[
    [
        "origen",
        "id_lista",
        "lista",
        "fecha",
        "product_id",
        "cod_art",
        "producto",
        "precio",
        "uom",
        "unidad_me",
        "invoicing_type",
        "contenido",
        "box_kgs",
        "precioxkg",
        "ing_hora",
    ]
].copy()


# =====================================================
# 10. CONEXIÓN A SQL SERVER
# =====================================================
conn = pyodbc.connect(
    "Driver={SQL Server};"
    "Server=lariosql70;"
    "Database=DW;"
    "UID=sa;"
    "PWD=sqlmanager;"
)
cursor = conn.cursor()
cursor.fast_executemany = True

# =====================================================
# 11. INSERT 
# =====================================================
sql_merge = """
MERGE DW.dbo.Lista_Precios_Odoo AS target
USING (
    SELECT
        ?  AS origen,
        ?  AS id_lista,
        ?  AS lista,
        ?  AS fecha,
        ?  AS product_id,
        ?  AS cod_art,
        ?  AS producto,
        ?  AS precio,
        ?  AS uom,
        ?  AS unidad_me,
        ?  AS invoicing_type,
        ?  AS contenido,
        ?  AS box_kgs,
        ?  AS precioxkg,
        ?  AS ing_hora
) AS source
ON (
    target.id_lista = source.id_lista
    AND target.cod_art = source.cod_art
    AND target.fecha   = source.fecha
)
WHEN NOT MATCHED THEN
    INSERT (
        origen, id_lista, lista, fecha, product_id,
        cod_art, producto, precio,
        uom, unidad_me, invoicing_type,
        contenido, box_kgs, precioxkg, ing_hora
    )
    VALUES (
        source.origen, source.id_lista, source.lista, source.fecha, source.product_id,
        source.cod_art, source.producto, source.precio,
        source.uom, source.unidad_me, source.invoicing_type,
        source.contenido, source.box_kgs, source.precioxkg, source.ing_hora
    );
"""




data = []

for _, r in df_load.iterrows():
    data.append((
        r.origen,
        r.id_lista,
        r.lista,
        r.fecha.to_pydatetime() if isinstance(r.fecha, pd.Timestamp) else r.fecha,
        r.product_id,
        int(r.cod_art),
        r.producto,
        float(r.precio),
        r.uom,
        r.unidad_me,
        r.invoicing_type,
        r.contenido,
        r.box_kgs,
        r.precioxkg,
        r.ing_hora.to_pydatetime() if isinstance(r.ing_hora, pd.Timestamp) else r.ing_hora,
    ))


import pyodbc

insertados = 0
omitidos = 0
errores = 0

for i, row in enumerate(data):
    try:
        cursor.execute(sql_merge, row)

        if cursor.rowcount == 1:
            insertados += 1
            logging.info(
                f"Fila {i} | INSERT OK | id_lista={row[1]}, cod_art={row[5]}, fecha={row[3]}"
            )
        else:
            omitidos += 1
            logging.info(
                f"Fila {i} | YA EXISTE | id_lista={row[1]}, cod_art={row[5]}, fecha={row[3]}"
            )

    except pyodbc.Error as e:
        errores += 1
        logging.error(
            f"Fila {i} | ERROR SQL | id_lista={row[1]}, cod_art={row[5]}, fecha={row[3]} | {e}"
        )



conn.commit()




logging.info("FIN proceso carga")
logging.info(f"Total filas procesadas: {len(data)}")
logging.info(f"Insertadas: {insertados}")
logging.info(f"Errores: {errores}")

print("Proceso finalizado")
print("Insertadas:", insertados)
print("Errores:", errores)
print("Log:", log_path)



