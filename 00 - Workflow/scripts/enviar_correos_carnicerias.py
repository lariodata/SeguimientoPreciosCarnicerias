import pandas as pd
import smtplib
from email.message import EmailMessage
import xlwings as xw
from datetime import datetime
import os

# ==============================
# CONFIGURACIÓN
# ==============================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Subir un nivel
EXCEL_PATH = os.path.join(BASE_DIR, "..", "WFW_SPC.xlsm")

SHEET_NAME = "Workflow"

SMTP_SERVER = "10.10.11.240"
SMTP_PORT = 20025
MAIL_FROM = "precioscarnicerias@lario.com.ar"

# ==============================
# CONECTAR A EXCEL ABIERTO
# ==============================
app = xw.apps.active
wb = app.books["WFW_SPC.xlsm"]
sht = wb.sheets[SHEET_NAME]

# ==============================
# LEER DATOS DESDE EXCEL
# ==============================
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)

para = str(df.iloc[5, 1]).strip()
cc = str(df.iloc[6, 1]).strip()
asunto = str(df.iloc[7, 1]).strip()
cuerpo = str(df.iloc[8, 1]).strip()

# ==============================
# ENVIAR MAIL
# ==============================
try:
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = para
    msg["Subject"] = asunto

    if cc and cc.lower() != "nan":
        msg["Cc"] = cc

    msg.set_content(cuerpo)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.send_message(msg)

    estado = f"✔ Correo Carnicerías enviado correctamente ({datetime.now():%d/%m %H:%M})"

except Exception as e:
    estado = f"❌ Error al enviar Correo Carnicerías: {str(e)}"

# ==============================
# MOSTRAR ESTADO EN TEXTO ESTÁTICO 25
# ==============================
sht.api.OLEObjects("Texto estático 25").Object.Caption = estado
