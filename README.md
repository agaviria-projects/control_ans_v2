# 🧩 Proyecto Control_ANS_FENIX

Este proyecto automatiza la limpieza y control de pedidos ANS descargados desde el sistema Fénix de EPM.  
Permite estandarizar campos, detectar vacíos, generar alertas y preparar los datos para análisis en Power BI.

## Estructura general
- **data_raw/**: Archivos originales descargados.
- **data_clean/**: Archivos listos para Power BI.
- **scripts/**: Limpieza, diagnóstico y panel tkinter.
- **venv/**: Entorno virtual Python.

## Día 1 – Configuración y limpieza base
- Creación del entorno virtual
- Instalación de librerías
- Primer script: `limpieza_fenix.py`
