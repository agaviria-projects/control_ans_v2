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
# DESCARGAR Y RENOMBRAR PDFS EN ONEDRIVE (v4.9 mejorado)
# ------------------------------------------------------------
def descargar_pdfs(service, df):
    # Normalizar encabezados para evitar errores por tildes o espacios
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("á", "a")
        .str.replace("é", "e")
        .str.replace("í", "i")
        .str.replace("ó", "o")
        .str.replace("ú", "u")
        .str.replace("ñ", "n")
    )
    print("🧭 Encabezados normalizados:", list(df.columns))

    # Buscar las columnas relevantes por nombre aproximado
    col_pedido = next((c for c in df.columns if "pedido" in c), None)
    col_tecnico = next((c for c in df.columns if "tecnic" in c), None)
    col_url = next((c for c in df.columns if "evidenc" in c), None)

    if not all([col_pedido, col_tecnico, col_url]):
        print("❌ No se pudieron identificar las columnas necesarias.")
        print(f"pedido={col_pedido}, tecnico={col_tecnico}, url={col_url}")
        return

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    carpeta_dia = os.path.join(RUTA_ONEDRIVE, fecha_hoy)
    os.makedirs(carpeta_dia, exist_ok=True)
    print(f"\n📁 Carpeta destino: {carpeta_dia}\n")

    for i, fila in df.iterrows():
        pedido = str(fila.get(col_pedido, "")).strip()
        tecnico = str(fila.get(col_tecnico, "")).strip()
        url = str(fila.get(col_url, "")).strip()

        if not (pedido and tecnico and url):
            print(f"⚠️ Fila {i+1} incompleta, se omite.")
            continue

        # Extraer ID de la URL
        if "id=" in url:
            file_id = url.split("id=")[-1]
        else:
            print(f"⚠️ URL inválida en la fila {i+1}: {url}")
            continue

        nombre_archivo = f"{pedido} - {tecnico}.pdf"
        ruta_local = os.path.join(carpeta_dia, nombre_archivo)

        try:
            print(f"⬇️ Descargando {nombre_archivo} ...")
            request = service.files().get_media(fileId=file_id)
            with io.FileIO(ruta_local, "wb") as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    if status:
                        progreso = int(status.progress() * 100)
                        print(f"   Progreso: {progreso}%")

            print(f"✅ Guardado en: {ruta_local}\n")

        except Exception as e:
            print(f"❌ Error al descargar {nombre_archivo}: {e}")

# ------------------------------------------------------------
# EJECUCIÓN COMPLETA
# ------------------------------------------------------------
if __name__ == "__main__":
    service = crear_servicio()
    df = leer_google_sheet(service)
    if df is not None:
        descargar_pdfs(service, df)