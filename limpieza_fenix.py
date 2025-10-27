"""
------------------------------------------------------------
LIMPIEZA BASE FÉNIX – Proyecto Control_ANS_FENIX
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
Descripción:
- Detecta automáticamente el CSV más reciente.
- Normaliza nombres de columnas.
- Mantiene las columnas clave, creando las faltantes vacías.
- Rellena celdas vacías con 'SIN DATOS'.
- Filtra actividades válidas.
- Limpia comillas y espacios.
- Exporta a Excel con tabla estructurada + hoja de resumen.
- Registra log de columnas y registros procesados.
------------------------------------------------------------
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# ------------------------------------------------------------
# CONFIGURACIÓN DE RUTAS
# ------------------------------------------------------------
base_path = Path(__file__).resolve().parent
ruta_clean = base_path / "data_clean" / "FENIX_CLEAN.xlsx"
ruta_log = base_path / "data_clean" / "log_limpieza.txt"

# Buscar archivo CSV más reciente
archivos_csv = sorted(base_path.glob("data_raw/pendientes_*.csv"), key=lambda x: x.stat().st_mtime, reverse=True)
if not archivos_csv:
    raise FileNotFoundError("No se encontró ningún archivo CSV en data_raw/")
ruta_raw = archivos_csv[0]

print(f"📂 Archivo detectado automáticamente: {ruta_raw.name}")

# ------------------------------------------------------------
# CARGA DE DATOS
# ------------------------------------------------------------
df = pd.read_csv(
    ruta_raw,
    encoding='latin-1',
    sep=',',
    quotechar='"',
    on_bad_lines='skip',
    engine='python'
)

# ------------------------------------------------------------
# LIMPIEZA BÁSICA
# ------------------------------------------------------------
import unicodedata

# Normaliza nombres de columnas: quita tildes, espacios y mayúsculas
def normalizar_columna(nombre):
    nombre = str(nombre).strip().upper().replace(" ", "_")
    # elimina tildes y caracteres especiales
    nombre = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )
    return nombre

df.columns = [normalizar_columna(c) for c in df.columns]


# Renombrar si hay tildes en columnas
if "TIPO_DIRECCIÓN" in df.columns and "TIPO_DIRECCION" not in df.columns:
    df.rename(columns={"TIPO_DIRECCIÓN": "TIPO_DIRECCION"}, inplace=True)

if "INSTALACIÓN" in df.columns and "INSTALACION" not in df.columns:
    df.rename(columns={"INSTALACIÓN": "INSTALACION"}, inplace=True)


# Columnas requeridas
columnas_utiles = [
    "PEDIDO", "PRODUCTO_ID", "TIPO_TRABAJO", "TIPO_ELEMENTO_ID",
    "FECHA_RECIBO", "FECHA_INICIO_ANS", "CLIENTEID", "NOMBRE_CLIENTE",
    "TELEFONO_CONTACTO", "CELULAR_CONTACTO", "DIRECCION",
    "MUNICIPIO", "INSTALACION", "AREA_TRABAJO", "ACTIVIDAD",
    "NOMBRE", "TIPO_DIRECCION"
]

# Crear columnas faltantes vacías
for col in columnas_utiles:
    if col not in df.columns:
        df[col] = None

# Reordenar columnas
df = df[columnas_utiles].copy()
print("✅ Todas las columnas requeridas presentes (faltantes creadas vacías).")

# ------------------------------------------------------------
# FILTRO DE ACTIVIDADES
# ------------------------------------------------------------
actividades_validas = [
    "ACREV", "ALEGN", "ALEGA", "ALEMN", "ACAMN",
    "AMRTR", "APLIN", "REEQU", "INPRE", "DIPRE",
    "ARTER", "AEJDO"
]
df = df[df["ACTIVIDAD"].isin(actividades_validas)]

# ------------------------------------------------------------
# LIMPIEZA DE TEXTO Y COMILLAS
# ------------------------------------------------------------
columnas_a_limpieza = ["DIRECCION", "INSTALACION"]
for col in columnas_a_limpieza:
    if col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("^'", "", regex=True)
            .str.replace("'", "", regex=False)
            .str.strip()
        )

# ------------------------------------------------------------
# RELLENAR VACÍOS CON 'SIN DATOS'
# ------------------------------------------------------------
df = df.fillna("SIN DATOS")
df.replace("", "SIN DATOS", inplace=True)

# ------------------------------------------------------------
# GENERAR RESUMEN
# ------------------------------------------------------------
total_registros = len(df)
filas_vacias = (df == "SIN DATOS").all(axis=1).sum()
duplicados_pedido = df.duplicated(subset="PEDIDO").sum()

resumen = pd.DataFrame({
    "MÉTRICA": ["Total registros", "Filas completamente vacías", "Duplicados por PEDIDO"],
    "VALOR": [total_registros, filas_vacias, duplicados_pedido]
})
# ------------------------------------------------------------
# CÁLCULO DE DIAS_PACTADOS SEGÚN ACTIVIDAD Y TIPO_DIRECCION
# ------------------------------------------------------------

def calcular_dias_pactados(fila):
    actividad = str(fila["ACTIVIDAD"]).upper().strip()
    tipo_dir = str(fila["TIPO_DIRECCION"]).upper().strip()

    # Reglas base (puedes ir agregando más)
    if actividad == "ALEGN":
        return 7 if tipo_dir == "URBANO" else 10 if tipo_dir == "RURAL" else 0
    if actividad == "ALEGA":
         return 7 if tipo_dir == "URBANO" else 10
    elif actividad == "ARTER":
        return 0 if tipo_dir == "URBANO" else 0
    else:
        return 0  # temporal mientras confirmas las demás reglas

# Aplicar la función a cada fila
df["DIAS_PACTADOS"] = df.apply(calcular_dias_pactados, axis=1)
print("🧮 Columna 'DIAS_PACTADOS' generada exitosamente.")

# ------------------------------------------------------------
# EXPORTACIÓN A EXCEL (2 hojas)
# ------------------------------------------------------------
ruta_clean.parent.mkdir(exist_ok=True)

with pd.ExcelWriter(ruta_clean, engine="openpyxl") as writer:
    # Hoja principal
    df.to_excel(writer, index=False, sheet_name="FENIX_CLEAN")
    ws = writer.sheets["FENIX_CLEAN"]

    n_filas, n_cols = df.shape
    ultima_col = chr(65 + n_cols - 1)
    rango_tabla = f"A1:{ultima_col}{n_filas + 1}"

    tabla = Table(displayName="TABLA_FENIX", ref=rango_tabla)
    estilo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True
    )
    tabla.tableStyleInfo = estilo
    ws.add_table(tabla)

    # Hoja de resumen
    resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
    ws2 = writer.sheets["RESUMEN"]

print("✅ Archivo limpio, con 'SIN DATOS' y resumen generado exitosamente.")
print(f"📁 Archivo: {ruta_clean}")
print(f"🧮 Registros: {len(df)}")
print(f"📝 Log: {ruta_log}")

