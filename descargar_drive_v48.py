# ============================================================
# DESCARGAR Y RENOMBRAR PDF DESDE GOOGLE SHEET - v4.12 FINAL
# Integración completa: Drive → OneDrive → Google Sheet
# Versión robusta con detección de entorno (Empresa / Personal)
# + creación automática de carpetas por responsable y actividad
# ============================================================

import os
import io
import gspread
import pandas as pd
import time
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from gspread.utils import rowcol_to_a1
from pathlib import Path

# ------------------------------------------------------------
# CONFIGURACIÓN BASE
# ------------------------------------------------------------
#CAMBIAR LA LINEA 26 POR ESTA 24
#CRED_PATH = r"C:\Users\hector.gaviria\Desktop\Control_ANS\control-ans-elite-f4ea102db569.json"

CRED_PATH = r"C:\Users\hector.gaviria\Desktop\Control_ANS\control-ans-elite-f4ea102db569.json"
SHEET_ID = "1bPLGVVz50k6PlNp382isJrqtW_3IsrrhGW0UUlMf-bM"

# ------------------------------------------------------------
# 🔄 DETECCIÓN AUTOMÁTICA DE ENTORNO (Empresa / Personal)
# ------------------------------------------------------------
RUTA_EMPRESA = Path(r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS\Evidencias_PDF")

if RUTA_EMPRESA.exists():
    RUTA_DESTINO = RUTA_EMPRESA
    print("🏢 Entorno detectado: Empresa (OneDrive conectado)")
else:
    RUTA_DESTINO = Path(r"C:\Users\Acer\Desktop\Evidencias_PDF")
    print("💻 Entorno detectado: Personal (modo pruebas en Desktop)")

# ------------------------------------------------------------
# 🧭 CONFIGURACIÓN DE FECHA SIN CREAR CARPETA BASE
# ------------------------------------------------------------
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
CARPETA_FECHA = RUTA_DESTINO  # ya no crea una carpeta por fecha

print(f"📂 Carpeta destino base: {RUTA_DESTINO}")

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
# CONECTAR A GOOGLE SHEET CON GSPREAD
# ------------------------------------------------------------
def conectar_gspread():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_ID)

    for ws in spreadsheet.worksheets():
        nombre = ws.title.lower().replace(" ", "")
        if "form" in nombre or "respuesta" in nombre:
            print(f"📄 Hoja activa detectada: {ws.title}")
            return ws

    print("⚠️ No se detectó hoja de respuestas; usando la primera hoja.")
    return spreadsheet.sheet1

# ------------------------------------------------------------
# LEER GOOGLE SHEET COMO CSV
# ------------------------------------------------------------
def leer_google_sheet(service):
    try:
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
# DESCARGAR Y RENOMBRAR PDFS POR RESPONSABLE Y ACTIVIDAD
# ------------------------------------------------------------
def descargar_pdfs(service, df):
    # Normalizar encabezados
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

    # Columnas clave
    col_pedido = next((c for c in df.columns if "pedido" in c), None)
    col_tecnico = next((c for c in df.columns if "tecnic" in c), None)
    col_actividad = next((c for c in df.columns if "actividad" in c), None)
    col_url = next((c for c in df.columns if "evidenc" in c), None)

    if not all([col_pedido, col_tecnico, col_actividad, col_url]):
        print("❌ No se pudieron identificar las columnas necesarias.")
        return

    # ------------------------------------------------------------
    # Mapeo de responsables por actividad (ajustado con formato real del formulario)
    # ------------------------------------------------------------
    RESPONSABLES = {
        "ARTER-(REPLANTEO PREPAGO)": "Lina",
        "ARTER-(REPLANTEO HV)": "Lina",
        "ACREV-(PUNTOS DE CONEXION)": "Lina",
        "ALEGA-(LEGALIZACION RESIDENCIAL)": "Frank",
        "ALEGN-(LEGALIZACION NO RESIDENCIAL)": "Frank",
        "ALECA-(REFORMA RESIDENCIAL)": "Frank",
        "ACAMN-(REFORMAS NO RESIDENCIAL)": "Frank",
        "AEJDO-(HV SENCILLO)": "Lina",
        "INPRE-(EJECUCION PREPAGO)": "Lina",
        "AMRTR-(MOVIMIENTOS DE REDES)": "Robinson",
        "AEJDO-(HV MAS INTERNA)": "Lina",
        "REEQU-(TRABAJOS PREPAGO)": "Lina",
        "DIPRE-(RETIRO PREPAGO)": "Lina"
    }

    # ------------------------------------------------------------
    # Función para obtener carpeta destino según actividad
    # ------------------------------------------------------------
    def obtener_ruta_destino(actividad):
        responsable = RESPONSABLES.get(actividad, "Sin_Asignar")

        # Si el responsable no tiene carpeta base, la crea automáticamente
        carpeta_responsable = RUTA_DESTINO / responsable
        if not carpeta_responsable.exists():
            carpeta_responsable.mkdir(parents=True, exist_ok=True)
            print(f"🆕 Carpeta creada para nuevo responsable: {responsable}")

        # Carpeta por fecha y actividad
        ruta_final = carpeta_responsable / fecha_hoy / actividad
        ruta_final.mkdir(parents=True, exist_ok=True)
        return ruta_final

    log_errores = CARPETA_FECHA / "log_errores_descarga.txt"
    errores = 0
    descargados = 0

    # ------------------------------------------------------------
    # Proceso de descarga
    # ------------------------------------------------------------
    for i, fila in df.iterrows():
        pedido = str(fila.get(col_pedido, "")).strip()
        tecnico = str(fila.get(col_tecnico, "")).strip()
        actividad = str(fila.get(col_actividad, "")).strip()
        url = str(fila.get(col_url, "")).strip()

        if not (pedido and tecnico and url):
            print(f"⚠️ Fila {i+1} incompleta, se omite.")
            continue

        if "id=" not in url:
            print(f"⚠️ URL inválida en la fila {i+1}: {url}")
            continue

        file_id = url.split("id=")[-1]
        nombre_archivo = f"EPM - {pedido} - {tecnico}.pdf"
        ruta_destino = obtener_ruta_destino(actividad)
        ruta_local = ruta_destino / nombre_archivo

        if ruta_local.exists():
            print(f"[INFO] Ya existe: {nombre_archivo}, se omite descarga.")
            continue

        try:
            print(f"⬇️ Descargando {nombre_archivo} en {ruta_destino} ...")
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
            descargados += 1
            time.sleep(0.8)

        except Exception as e:
            errores += 1
            print(f"❌ Error al descargar {nombre_archivo}: {e}")
            with open(log_errores, "a", encoding="utf-8") as log:
                log.write(f"{pedido} - {tecnico}: {e}\n")

    print("\n---------------------------------------------")
    print(f"✅ Descargas completadas: {descargados}")
    print(f"⚠️ Errores registrados: {errores}")
    if errores > 0:
        print(f"📄 Ver log: {log_errores}")
    print("---------------------------------------------\n")

# ------------------------------------------------------------
# ACTUALIZAR RUTAS EN GOOGLE SHEET
# ------------------------------------------------------------
def actualizar_rutas_locales(df):
    print("\n🔄 Iniciando actualización de rutas en Google Sheet...")

    try:
        sheet = conectar_gspread()
    except Exception as e:
        print(f"❌ Error conectando a Google Sheet: {e}")
        return

    data = sheet.get_all_records()
    encabezados_original = sheet.row_values(1)

    col_evidencia_index = None
    for idx, name in enumerate(encabezados_original, start=1):
        name_clean = str(name).strip().lower().replace(" ", "")
        if "evidenc" in name_clean or "subeaqu" in name_clean:
            col_evidencia_index = idx
            break

    if not col_evidencia_index:
        print("❌ No se detectó la columna de evidencia.")
        return

    total_registros = len(data)
    enlaces_actualizados = 0
    enlaces_no_encontrados = 0

    for i, fila in enumerate(data, start=2):
        pedido = str(fila.get("Número del pedido", "")).strip()
        tecnico = str(fila.get("Nombre del técnico", "")).strip()
        if not pedido or not tecnico:
            continue

        nombre_pdf = f"EPM - {pedido} - {tecnico}.pdf"
        ruta_local = next(CARPETA_FECHA.glob(f"*/**/{nombre_pdf}"), None)

        if ruta_local and ruta_local.exists():
            celda = rowcol_to_a1(i, col_evidencia_index)
            ruta_web = str(ruta_local).replace(
                r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS",
                "https://eliteingenierosas-my.sharepoint.com/personal/h_gaviria_eliteingenieros_com_co/Documents"
            ).replace("\\", "/")
            sheet.update_acell(celda, f'=HIPERVINCULO("{ruta_web}"; "Abrir PDF")')
            enlaces_actualizados += 1
            print(f"✅ Enlace actualizado: {nombre_pdf}")
        else:
            enlaces_no_encontrados += 1
            print(f"⚠️ No se encontró el PDF: {nombre_pdf}")

    print("\n🎯 Actualización completada.")
    print(f"✅ Enlaces correctos: {enlaces_actualizados}")
    print(f"⚠️ No encontrados: {enlaces_no_encontrados}")

# ------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ------------------------------------------------------------
if __name__ == "__main__":
    service = crear_servicio()
    df = leer_google_sheet(service)
    if df is not None:
        descargar_pdfs(service, df)
        actualizar_rutas_locales(df)
