#Este archivo lo dejamos aca pero pertenece al proyecto de Procesos ya que dispara ETL automático como tarea programada!
#Correr en proyecto Procesos

import os
import time
from datetime import datetime
from dotenv import load_dotenv
from filelock import FileLock, Timeout

from core.db.ConnectionFactory import ConnectionFactory
from core.config.utils import get_env_config
from core.config.db import DB
from core.utils.logger import setup_logger
from core.utils.correo import notificar_finalizacion_etl


# Cargar configuración
env_file = os.getenv("ENV_FILE", ".env.dev")
load_dotenv(dotenv_path=env_file)

# Configuración general
LOG_DIR = "logs"
LOCK_FILE = "etl_seguimientoprecios.lock"
LOG_FILE = os.path.join(LOG_DIR, f"etl_seguimientoprecios.log")
logger = setup_logger("ETL_SeguimientoPrecios", LOG_FILE)


def ejecutar_sp_gold(db, sp_name):
    try:
        db.execute(f"EXEC {sp_name}", autocommit=True)
        msg = f"{sp_name} ejecutado correctamente"
        logger.info("✔ " + msg)
        return msg
    except Exception:
        logger.exception(f"❌ Error en {sp_name}")
        raise




def Llenar_GOLD_SeguimientoPrecios(db_DW) -> str:
    inicio = time.time()
    resumen = []

    try:
        # ================= PRECIOS =================
        resumen.append(
            ejecutar_sp_gold(
                db=db_DW,
                sp_name="dbo.SP_GOLD_CargarFactListaPreciosSemanal"
            )
        )

        # ================= COSTOS =================
        resumen.append(
            ejecutar_sp_gold(
                db=db_DW,
                sp_name="dbo.SP_GOLD_CargarFactCostosSemanal"
            )
        )

        tiempo = time.strftime("%H:%M:%S", time.gmtime(time.time() - inicio))
        resumen.append(f"Tiempo total: {tiempo}")

        return "\n".join(resumen)

    except Exception as ex:
        raise RuntimeError(f"ETL GOLD Seguimiento Precios falló: {ex}")


# ======================== EJECUCIÓN =============================
if __name__ == "__main__":
    from sys import argv

    resumen = ""
    estado_error = False
    p_iniciar = int(argv[1]) if len(argv) > 1 else 30

    db_DW = db_legacy = None

    try:
        with FileLock(LOCK_FILE, timeout=1):

            if os.getenv("AMBIENTE") == "PROD":
                db_DW = ConnectionFactory.create("sqlserver", get_env_config(DB.DW_SQLSERVER))
            else:
                db_DW = ConnectionFactory.create("sqlserver", get_env_config(DB.DW_SQLSERVER))
                

            resumen = Llenar_GOLD_SeguimientoPrecios(db_DW)
    except Timeout:
        resumen = "⏱ El proceso fue bloqueado por timeout."
        estado_error = True
    except RuntimeError as err:
        resumen = str(err)
        estado_error = True
    except Exception as e:
        resumen = str(e)
        estado_error = True
    finally:
        if db_DW: db_DW.close()
        if db_legacy: db_legacy.close()

        asunto = f"{'❌ Error! ' if estado_error else ''}ETL_SeguimientoPrecios - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        notificar_finalizacion_etl(asunto=asunto, proceso="TEAM_BI_PROCESO_OK", mensaje=resumen, error=estado_error)