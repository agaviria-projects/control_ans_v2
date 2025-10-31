# -*- coding: utf-8 -*-
# DESCARGAR EVIDENCIAS DE GOOGLE DRIVE Y MOVER A PAPELERA_API
# ------------------------------------------------------------
import os
import io
import time
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import sys

# Forzar salida UTF-8 para registros
sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# CONFIGURACIÓN
# ============================================================
CARPETA_LOCAL = r"C:\Users\hector.gaviria\Desktop\Control_ANS\Evidencias_ANS"
FOLDER_ID_FORMULARIO = "1cgtia-u95riQzBiqIV4IOw6STXix39Ibry2wGIAWAiyiawdkyTL3Eoln33i82SNyB4dYt9ss"
FOLDER_ID_PAPELERA = "1t8yIQGQJ_Qi0c4ejDUMcr6H8Qz09-O9b"
CRED_PATH = "control-ans-evidencias-1ef0b1b8d1a8.json"

# ============================================================
# AUTENTICACIÓN
# ============================================================
def crear_servicio():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

# ============================================================
# DESCARGAR Y MOVER ARCHIVOS (versión optimizada)
# ============================================================
def descargar_archivos(service):
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    carpeta_dia = os.path.join(CARPETA_LOCAL, fecha_hoy)
    os.makedirs(carpeta_dia, exist_ok=True)

    print(f"\n[INFO] Descargando evidencias del {fecha_hoy}...\n")

    query = f"'{FOLDER_ID_FORMULARIO}' in parents and mimeType != 'application/vnd.google-apps.folder'"
    results = service.files().list(q=query, fields="files(id, name, parents)").execute()
    files = results.get("files", [])

    if not files:
        print("[WARN] No se encontraron archivos en la carpeta del formulario.")
        return

    descargados = 0
    movidos = 0

    for file in files:
        file_id = file["id"]
        file_name = file["name"]
        file_path = os.path.join(carpeta_dia, file_name)

        try:
            print(f"[📥] Descargando {file_name}...")
            request = service.files().get_media(fileId=file_id)
            fh = io.FileIO(file_path, "wb")

            # ✅ Descarga más rápida con chunks grandes
            downloader = MediaIoBaseDownload(fh, request, chunksize=1024 * 1024)  # 1 MB por bloque
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    progreso = int(status.progress() * 100)
                    print(f"   Progreso: {progreso}%")
            fh.close()

            descargados += 1
            print(f"[OK] Archivo descargado: {file_name}")

            # ✅ Mover archivo al Drive/PAPELERA_API
            file_metadata = service.files().get(fileId=file_id, fields="parents").execute()
            padres = ",".join(file_metadata.get("parents", []))

            service.files().update(
                fileId=file_id,
                addParents=FOLDER_ID_PAPELERA,
                removeParents=padres
            ).execute()

            movidos += 1
            print(f"[MOVIDO] Archivo movido a PAPELERA_API: {file_name}")

            # ✅ Pequeña pausa para evitar límites de la API
            time.sleep(0.3)

        except Exception as e:
            print(f"[ERROR] No se pudo procesar {file_name}: {e}")

    print(f"\n✅ Total de archivos descargados: {descargados}")
    print(f"🗑️ Total de archivos movidos a PAPELERA_API: {movidos}")

    # MENSAJE AUTOMÁTICO FINAL
    print("\n------------------------------------------------------------")
    print("[OK] PROCESO COMPLETADO CON ÉXITO")
    print("[INFO] La carpeta del formulario quedó vacía.")
    print("[INFO] Los archivos se encuentran respaldados en:")
    print(f"       → {carpeta_dia}")
    print("[INFO] Los archivos del Drive fueron movidos a la carpeta: PAPELERA_API")
    print("[TIP]  Cuando desees liberar espacio, entra a Google Drive → PAPELERA_API y elimina definitivamente los archivos.")
    print("------------------------------------------------------------\n")

# ============================================================
# EJECUCIÓN
# ============================================================
if __name__ == "__main__":
    service = crear_servicio()
    descargar_archivos(service)
