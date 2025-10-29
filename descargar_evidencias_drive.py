# ------------------------------------------------------------
# DESCARGAR EVIDENCIAS DE GOOGLE DRIVE Y MOVER A PAPELERA_API
# ------------------------------------------------------------
import os
import io
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

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
# DESCARGAR Y MOVER ARCHIVOS
# ============================================================
def descargar_archivos(service):
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    carpeta_dia = os.path.join(CARPETA_LOCAL, fecha_hoy)
    os.makedirs(carpeta_dia, exist_ok=True)

    print(f"\n📁 Descargando evidencias del {fecha_hoy}...\n")

    query = f"'{FOLDER_ID_FORMULARIO}' in parents and mimeType != 'application/vnd.google-apps.folder'"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])

    if not files:
        print("⚠️ No se encontraron archivos en la carpeta del formulario.")
        return

    descargados = 0
    movidos = 0

    for file in files:
        file_id = file["id"]
        file_name = file["name"]
        file_path = os.path.join(carpeta_dia, file_name)

        print(f"⬇️ Descargando: {file_name}...")
        request = service.files().get_media(fileId=file_id)
        fh = io.FileIO(file_path, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"   Progreso: {int(status.progress() * 100)}%")

        descargados += 1
        print(f"✅ Archivo descargado: {file_name}")

        # Mover archivo a la carpeta PAPELERA_API
        try:
            service.files().update(
                fileId=file_id,
                addParents=FOLDER_ID_PAPELERA,
                removeParents=FOLDER_ID_FORMULARIO
            ).execute()
            movidos += 1
            print(f"🗂️ Archivo movido a PAPELERA_API: {file_name}")
        except Exception as e:
            print(f"⚠️ No se pudo mover {file_name}: {e}")

    print(f"\n✅ Todos los archivos ({descargados}) se guardaron en: {carpeta_dia}")
    print(f"🗑️ Archivos movidos correctamente a PAPELERA_API: {movidos}")

    # MENSAJE AUTOMÁTICO FINAL
    print("\n------------------------------------------------------------")
    print("✅ PROCESO COMPLETADO CON ÉXITO")
    print("📂 La carpeta del formulario quedó vacía.")
    print("📦 Los archivos se encuentran respaldados en:")
    print(f"   → {carpeta_dia}")
    print("🗂️ Los archivos del Drive fueron movidos a la carpeta: PAPELERA_API")
    print("💡 TIP: Cuando desees liberar espacio, entra a Google Drive → PAPELERA_API y elimina definitivamente los archivos.")
    print("------------------------------------------------------------\n")

# ============================================================
# EJECUCIÓN
# ============================================================
if __name__ == "__main__":
    service = crear_servicio()
    descargar_archivos(service)

