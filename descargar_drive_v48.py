# ============================================================
# DESCARGAR Y RENOMBRAR PDF DESDE GOOGLE SHEET - v4.12 FINAL
# Integración completa: Drive → OneDrive → Google Sheet
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

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------
CRED_PATH = r"C:\Users\hector.gaviria\Desktop\Control_ANS\control-ans-elite-f4ea102db569.json"
SHEET_ID = "1bPLGVVz50k6PlNp382isJrqtW_3IsrrhGW0UUlMf-bM"
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
# CONECTAR A GOOGLE SHEET CON GSPREAD
# ------------------------------------------------------------
def conectar_gspread():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
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
# DESCARGAR Y RENOMBRAR PDFS EN ONEDRIVE
# ------------------------------------------------------------
def descargar_pdfs(service, df):
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

    col_pedido = next((c for c in df.columns if "pedido" in c), None)
    col_tecnico = next((c for c in df.columns if "tecnic" in c), None)
    col_url = next((c for c in df.columns if "evidenc" in c), None)

    if not all([col_pedido, col_tecnico, col_url]):
        print("❌ No se pudieron identificar las columnas necesarias.")
        return

    # Crear carpeta según la fecha del formulario
    df["marca_temporal"] = pd.to_datetime(df["marca_temporal"], errors="coerce", dayfirst=True)
    fecha_formulario = df["marca_temporal"].max().strftime("%Y-%m-%d")
    carpeta_dia = os.path.join(RUTA_ONEDRIVE, fecha_formulario)
    os.makedirs(carpeta_dia, exist_ok=True)
    print(f"📁 Carpeta destino creada según formulario: {carpeta_dia}")

    # Registro de errores
    log_errores = os.path.join(carpeta_dia, "log_errores_descarga.txt")
    errores = 0
    descargados = 0

    for i, fila in df.iterrows():
        pedido = str(fila.get(col_pedido, "")).strip()
        tecnico = str(fila.get(col_tecnico, "")).strip()
        url = str(fila.get(col_url, "")).strip()

        if not (pedido and tecnico and url):
            print(f"⚠️ Fila {i+1} incompleta, se omite.")
            continue

        if "id=" not in url:
            print(f"⚠️ URL inválida en la fila {i+1}: {url}")
            continue

        file_id = url.split("id=")[-1]
        nombre_archivo = f"{pedido} - {tecnico}.pdf"
        ruta_local = os.path.join(carpeta_dia, nombre_archivo)

        if os.path.exists(ruta_local):
            print(f"[INFO] Ya existe: {nombre_archivo}, se omite descarga.")
            continue

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
# ACTUALIZAR RUTAS LOCALES EN GOOGLE SHEET
# ------------------------------------------------------------
def actualizar_rutas_locales(df):
    print("\n🔄 Iniciando actualización de rutas en Google Sheet...")

    try:
        sheet = conectar_gspread()
    except Exception as e:
        print(f"❌ Error conectando a Google Sheet: {e}")
        return

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"[\r\n]+", " ", regex=True)
        .str.replace(" ", "_")
        .str.lower()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )

    col_pedido = next((c for c in df.columns if "pedido" in c), None)
    col_tecnico = next((c for c in df.columns if "tecnic" in c), None)
    if not all([col_pedido, col_tecnico]):
        print("❌ No se encontraron las columnas de pedido y técnico.")
        return

    data = sheet.get_all_records()
    encabezados_original = sheet.row_values(1)

    col_evidencia_index = None
    for idx, name in enumerate(encabezados_original, start=1):
        name_clean = (
            str(name)
            .replace("\n", "")
            .replace("\r", "")
            .strip()
            .lower()
            .replace(" ", "")
        )
        if "evidenc" in name_clean or "subeaqu" in name_clean:
            col_evidencia_index = idx
            print(f"📍 Columna de evidencia detectada: {name} (índice {idx})")
            break

    if not col_evidencia_index:
        print("❌ No se detectó la columna de evidencia.")
        print("Encabezados encontrados:", encabezados_original)
        return

    def normalizar_nombre(texto):
        return (
            str(texto)
            .strip()
            .lower()
            .replace(" ", "")
            .replace("á", "a")
            .replace("é", "e")
            .replace("í", "i")
            .replace("ó", "o")
            .replace("ú", "u")
            .replace("ñ", "n")
        )

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    carpeta_dia = os.path.join(RUTA_ONEDRIVE, fecha_hoy)
    print(f"📋 Registros totales: {len(data)}")

    total_registros = len(data)
    enlaces_actualizados = 0
    enlaces_no_encontrados = 0

    for i, fila in enumerate(data, start=2):
        fila_normalizada = {normalizar_nombre(k): v for k, v in fila.items()}
        pedido = str(fila_normalizada.get("numerodelpedido", "")).strip()
        tecnico = str(fila_normalizada.get("nombredeltecnico", "")).strip()
        if not pedido or not tecnico:
            continue

        nombre_pdf = f"{pedido} - {tecnico}.pdf"
        ruta_local = os.path.join(carpeta_dia, nombre_pdf)

        if os.path.exists(ruta_local):
            celda = rowcol_to_a1(i, col_evidencia_index)
            ruta_web = ruta_local.replace(
                r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS",
                "https://eliteingenierosas-my.sharepoint.com/personal/h_gaviria_eliteingenieros_com_co/Documents"
            ).replace("\\", "/")

            sheet.update_acell(celda, f'=HIPERVINCULO("{ruta_web}"; "Abrir PDF")')
            time.sleep(1)
            enlaces_actualizados += 1
            print(f"✅ Enlace web actualizado para {nombre_pdf}")
        else:
            enlaces_no_encontrados += 1
            print(f"⚠️ No se encontró el PDF: {nombre_pdf}")

    print("\n🎯 Actualización de rutas completada.\n")
    print(f"📊 Total de registros procesados: {total_registros}")
    print(f"✅ Enlaces actualizados correctamente: {enlaces_actualizados}")
    print(f"⚠️ PDFs no encontrados: {enlaces_no_encontrados}")
    print("\n💡 Tip: Verifica en Google Sheets que los enlaces 'Abrir PDF' sean clickeables.\n")

# ------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ------------------------------------------------------------
if __name__ == "__main__":
    service = crear_servicio()
    df = leer_google_sheet(service)
    if df is not None:
        descargar_pdfs(service, df)
        actualizar_rutas_locales(df)
