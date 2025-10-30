# ============================================================
# DESCARGAR Y RENOMBRAR PDF DESDE GOOGLE SHEET - v4.8 (método alternativo sin Sheets API)
# ============================================================
import os
import io
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------
CRED_PATH = "control-ans-evidencias-1ef0b1b8d1a8.json"
SHEET_ID = "1bPLGVVz50k6PlNp382isJrqtW_3IsrrhGW0UUlMf-bM"  # ID del Formulario Control ANS
RUTA_ONEDRIVE = r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS\Evidencias_PDF"

# ------------------------------------------------------------
# AUTENTICACIÓN A GOOGLE DRIVE
# ------------------------------------------------------------
def crear_servicio():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

# ------------------------------------------------------------
# LEER GOOGLE SHEET COMO CSV
# ------------------------------------------------------------
def leer_google_sheet(service):
    try:
        # Exportar el sheet como CSV temporal
        request = service.files().export_media(fileId=SHEET_ID, mimeType="text/csv")
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_csv(fh)
        print("✅ Hoja leída correctamente.\n")
        print(df.head())
        return df
    except Exception as e:
        print(f"❌ Error al leer Google Sheet: {e}")
        return None

# ------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ------------------------------------------------------------
if __name__ == "__main__":
    service = crear_servicio()
    df = leer_google_sheet(service)
