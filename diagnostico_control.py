"""
------------------------------------------------------------
DIAGNÓSTICO DE DATOS – Proyecto Control_ANS_FENIX
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
"""

import pandas as pd
from pathlib import Path

# ------------------------------------------------------------
# CONFIGURACIÓN DE RUTAS
# ------------------------------------------------------------
base_path = Path(__file__).resolve().parent
ruta_clean = base_path / "data_clean" / "FENIX_CLEAN.xlsx"

# ------------------------------------------------------------
# CARGA DEL ARCHIVO LIMPIO
# ------------------------------------------------------------
try:
    df = pd.read_excel(ruta_clean, sheet_name="FENIX_CLEAN")
    print(f"✅ Archivo cargado correctamente: {ruta_clean.name}")
except FileNotFoundError:
    print("❌ No se encontró el archivo limpio FENIX_CLEAN.xlsx. Ejecuta primero limpieza_fenix.py.")
    exit()

# ------------------------------------------------------------
# INFORME GENERAL
# ------------------------------------------------------------
print("\n------------------------------------------------------------")
print("📊 INFORME GENERAL DEL ARCHIVO")
print("------------------------------------------------------------")
print(f"Total de filas: {df.shape[0]}")
print(f"Total de columnas: {df.shape[1]}\n")

print("Primeras columnas detectadas:")
print(list(df.columns[:10]))  # mostramos solo las primeras 10 por si hay muchas

print("\nTipos de datos:")
print(df.dtypes.head(10))

# ------------------------------------------------------------
# DETECCIÓN DE VALORES NULOS
# ------------------------------------------------------------
print("\n------------------------------------------------------------")
print("⚠️  Columnas con valores vacíos")
print("------------------------------------------------------------")
nulos = df.isna().sum()
nulos = nulos[nulos > 0].sort_values(ascending=False)

if nulos.empty:
    print("No se encontraron valores vacíos.")
else:
    print(nulos)

# ------------------------------------------------------------
# DETECCIÓN DE DUPLICADOS
# ------------------------------------------------------------
print("\n------------------------------------------------------------")
print("🔁 POSIBLES DUPLICADOS (filas idénticas)")
print("------------------------------------------------------------")
duplicados = df.duplicated().sum()
print(f"Filas duplicadas detectadas: {duplicados}")

# ------------------------------------------------------------
# DETECCIÓN DE COLUMNAS VACÍAS O CONSTANTES
# ------------------------------------------------------------
print("\n------------------------------------------------------------")
print("🧱 Columnas sin variación (solo un valor único o vacías)")
print("------------------------------------------------------------")
constantes = [col for col in df.columns if df[col].nunique(dropna=True) <= 1]
if constantes:
    print(constantes)
else:
    print("No hay columnas constantes detectadas.")

# ------------------------------------------------------------
# SUGERENCIA DE CLAVES POSIBLES
# ------------------------------------------------------------
print("\n------------------------------------------------------------")
print("🔑 Posibles columnas clave (ID, PEDIDO, ORDEN, etc.)")
print("------------------------------------------------------------")
claves = [col for col in df.columns if "ID" in col or "PEDIDO" in col or "ORDEN" in col]
if claves:
    print(claves)
else:
    print("No se detectaron columnas con 'ID', 'PEDIDO' o 'ORDEN'.")
