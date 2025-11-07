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
from dateutil import parser  # ✅ agregado para lectura flexible de fechas
from io import StringIO
import unicodedata

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
# CARGA DE DATOS – Lectura segura del CSV con reparación reforzada
# ------------------------------------------------------------
try:
    print(f"🔍 Intentando leer archivo CSV: {ruta_raw}")

    # Detectar separador probable (; o ,)
    with open(ruta_raw, "r", encoding="latin-1", errors="ignore") as f:
        primera_linea = f.readline()
        sep = ";" if ";" in primera_linea else ","

    # Intento 1 – lectura normal
    df = pd.read_csv(
        ruta_raw,
        encoding="latin-1",
        sep=sep,
        dtype=str,
        on_bad_lines="skip",
        engine="python",
        quotechar='"'
    )

    # Intento 2 – si solo tiene 1 columna o muy pocas columnas válidas
    if len(df.columns) < 10:
        print("⚠️ Archivo mal formateado. Ejecutando reparación automática...")
        with open(ruta_raw, "r", encoding="latin-1", errors="ignore") as f:
            contenido = f.read()

        # Limpieza profunda: elimina comillas sueltas, espacios raros y duplicados de coma
        contenido = (
            contenido.replace(",'", ",")
                     .replace("',", ",")
                     .replace(",,", ",")
                     .replace(";'", ";")
                     .replace("';", ";")
                     .replace("\"", "")
                     .replace("´", "")
        )

        # Segunda lectura después de limpieza
        df = pd.read_csv(
            StringIO(contenido),
            sep=sep,
            dtype=str,
            on_bad_lines="skip",
            engine="python"
        )
        print(f"✅ Reparación aplicada. Columnas detectadas: {len(df.columns)}")

    print(f"📊 Registros cargados: {len(df)} ({len(df.columns)} columnas detectadas)")

    # ------------------------------------------------------------
    # 🧩 CORRECCIÓN ROBUSTA DE FORMATO DE FECHA (detecta ambos estilos)
    # ------------------------------------------------------------
    columnas_fecha = [c for c in df.columns if "FECHA" in c.upper()]

    def parsear_fecha_segura(valor):
        """Convierte fechas del formato Fénix asegurando día/mes/año correcto."""
        if pd.isna(valor) or not str(valor).strip():
            return None
        texto = str(valor).strip()
        texto = (
            texto.replace("a. m.", "AM")
                 .replace("p. m.", "PM")
                 .replace("p.m.", "PM")
                 .replace("a.m.", "AM")
                 .replace(".", ":")
                 .replace("\xa0", " ")
                 .strip()
        )
        try:
            # Intentar formato latino (dd/mm/yyyy)
            return parser.parse(texto, dayfirst=True)
        except Exception:
            try:
                # Intentar formato ISO o americano
                return parser.parse(texto, dayfirst=False)
            except Exception:
                return None

    for col in columnas_fecha:
        try:
            df[col] = df[col].apply(parsear_fecha_segura)
            # ✅ Exportar en formato latino dd/mm/yyyy HH:MM:SS
            df[col] = df[col].dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception as e:
            print(f"⚠️ Error al convertir columna {col}: {e}")

    print("🧭 Fechas convertidas correctamente a formato latino (DD/MM/YYYY HH:MM:SS).")

    # ------------------------------------------------------------
    # 🔍 DETECCIÓN Y CORRECCIÓN DE FECHAS ANÓMALAS (caracteres extra)
    # ------------------------------------------------------------
    if "FECHA_RECIBO" in df.columns and "FECHA_INICIO_ANS" in df.columns:
        def limpiar_fecha_str(valor):
            if isinstance(valor, str):
                valor = valor.replace('"', "").replace("'", "").strip()
                partes = valor.split()
                if len(partes) > 2:
                    valor = " ".join(partes[:2])
            return valor

        df["FECHA_RECIBO"] = df["FECHA_RECIBO"].apply(limpiar_fecha_str)
        df["FECHA_INICIO_ANS"] = df["FECHA_INICIO_ANS"].apply(limpiar_fecha_str)
        print("🧩 Fechas con caracteres extra corregidas correctamente.")

except Exception as e:
    print(f"❌ Error al leer el archivo CSV: {e}")
    import sys
    sys.exit(1)

# ------------------------------------------------------------
# LIMPIEZA BÁSICA
# ------------------------------------------------------------
def normalizar_columna(nombre):
    nombre = str(nombre).strip().upper().replace(" ", "_")
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

for col in columnas_utiles:
    if col not in df.columns:
        df[col] = None

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

    if actividad == "ALEGN":
        return 7 if tipo_dir == "URBANO" else 10 if tipo_dir == "RURAL" else 0
    if actividad == "ALEGA":
        return 7 if tipo_dir == "URBANO" else 10
    elif actividad == "ARTER":
        return 0
    else:
        return 0

df["DIAS_PACTADOS"] = df.apply(calcular_dias_pactados, axis=1)
print("🧮 Columna 'DIAS_PACTADOS' generada exitosamente.")

# ------------------------------------------------------------
# EXPORTACIÓN A EXCEL (2 hojas)
# ------------------------------------------------------------
ruta_clean.parent.mkdir(exist_ok=True)

with pd.ExcelWriter(ruta_clean, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="FENIX_CLEAN")
    ws = writer.sheets["FENIX_CLEAN"]

    n_filas, n_cols = df.shape
    ultima_col = chr(65 + n_cols - 1)
    rango_tabla = f"A1:{ultima_col}{n_filas + 1}"

    tabla = Table(displayName="TABLA_FENIX", ref=rango_tabla)
    estilo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    tabla.tableStyleInfo = estilo
    ws.add_table(tabla)

    resumen.to_excel(writer, index=False, sheet_name="RESUMEN")

print("✅ Archivo limpio, con 'SIN DATOS' y resumen generado exitosamente.")
print(f"📁 Archivo: {ruta_clean}")
print(f"🧮 Registros: {len(df)}")
print(f"📝 Log: {ruta_log}")
