
**Ejemplo de registro guardado:**

| fecha_envio | pedido | observacion | estado_campo | metodo_envio | pdf | imagenes |
|--------------|---------|--------------|---------------|---------------|------|-----------|
| 2025-10-24 14:53:37 | 23260219 | Generado con satisfacción | Cumplido | Formulario | 23260219_1_20251024_145336.pdf | 23260219_1_20251024_145337.jpg |

---

### 2️⃣ **Cálculos ANS (calculos_ans.py)**

Script que procesa el archivo **`FENIX_CLEAN.xlsx`** y genera **`FENIX_ANS.xlsx`**, aplicando toda la lógica de tiempos y semáforos.

**Funcionalidades principales:**
- Calcula **días pactados** según actividad (urbano/rural).
- Excluye **sábados, domingos y festivos**.
- Calcula:
- `FECHA_LIMITE_ANS`
- `DIAS_TRANSCURRIDOS`
- `DIAS_RESTANTES`
- `ESTADO` (VENCIDO, ALERTA, A TIEMPO)
- Agrega formato condicional en Excel con colores:
- 🟥 **VENCIDO**
- 🟧 **ALERTA 0 días**
- 🟡 **ALERTA 1-2 días**
- 🟩 **A TIEMPO**
- Genera hoja adicional `CONFIG_DIAS_PACTADOS` y `META_INFO` con metadatos del proceso.
- Prepara salida lista para conexión a **Power BI**.

**Dependencias:**  
`pandas`, `numpy`, `openpyxl`, `tkinter`, `datetime`

---

### 3️⃣ **Control FÉNIX vs ALMACÉN (validar_export_almacen.py)**

Script principal para detectar **diferencias entre FÉNIX y Planilla de Consumos (Elite)**.

**Flujo de proceso:**
1. Detecta automáticamente si el archivo base es `.txt` o `.xlsx`.
2. Limpia encabezados y elimina hojas no relevantes.
3. Estandariza columnas de ambos orígenes (`pedido`, `codigo`, `cantidad`).
4. Realiza `merge` extendido (outer join) entre FÉNIX y Elite.
5. Calcula:
 - `cantidad_fenix`
 - `cantidad_elite`
 - `diferencia`
 - `status` (`OK`, `FALTANTE EN ELITE`, `EXCESO EN ELITE`)
6. Aplica reglas especiales para materiales complementarios (`200492 ↔ 200492A`).
7. Agrega columna `TÉCNICO` desde Planilla de Consumos.
8. Reconstruye hoja `NO_COINCIDEN` con cantidades reales.
9. Genera resumen global de estados.

**Salidas:**
- `CONTROL_ALMACEN.xlsx` con 3 hojas:
- 🧾 **CONTROL_ALMACEN** → cruce completo  
- 📊 **RESUMEN** → conteo por estado  
- 🚨 **NO_COINCIDEN** → faltantes o excesos  

**Formato automático en Excel:**
- Encabezados coloreados por tipo (FENIX / ELITE / DIFERENCIA / STATUS).
- Semáforo por estado (`OK`, `FALTANTE`, `EXCESO`).
- Bordes, centrado y ancho ajustado automáticamente.

---

### 4️⃣ **Limpieza de FENIX (limpieza_fenix.py)**

Limpia los datos brutos exportados del sistema FÉNIX:
- Elimina duplicados.
- Normaliza nombres de columnas.
- Corrige tipos de datos.
- Prepara estructura base para los cálculos ANS.

---

### 5️⃣ **Diagnóstico y Validación (diagnostico_control.py)**

Evalúa calidad de datos:
- Detecta columnas vacías o mal tipadas.
- Identifica diferencias entre versiones.
- Apoya depuración en entornos empresariales.

---

## 📊 Integración con Power BI

Los archivos generados (`FENIX_ANS.xlsx` y `CONTROL_ALMACEN.xlsx`) se cargan directamente en Power BI para análisis:

- **Indicadores:** % Cumplimiento, Pedidos Vencidos, Alertas.  
- **Filtros:** Zona, Municipio, Técnico, Contrato.  
- **Visualizaciones:** Tablas, mapas, KPIs, líneas de tendencia.

---

## 🧱 Dependencias e Instalación

Instalar en entorno virtual (recomendado):

```bash
python -m venv venv
source venv/Scripts/activate   # Windows
pip install -r requirements.txt

Requerimientos:
Flask
pandas
numpy
openpyxl
gunicorn

🧰 Buenas Prácticas y Tips

Ejecutar con todos los archivos Excel cerrados.

Evitar subir archivos temporales (~$*.xlsx) → ya incluidos en .gitignore.

Los nombres de archivo incluyen timestamp para evitar duplicados.

Al modificar lógica, crear un commit versionado:

git add .
git commit -m "vX.X Descripción del cambio"
git push origin main

| Versión  | Fecha    | Cambios principales                                          |
| -------- | -------- | ------------------------------------------------------------ |
| **v3.2** | Sep 2025 | Cruce FENIX vs Elite, lectura flexible TXT/XLSX.             |
| **v3.7** | Oct 2025 | Colores de encabezado, semáforo y ajuste técnico.            |
| **v4.0** | Oct 2025 | Reconstrucción de hoja NO_COINCIDEN con cantidades reales.   |
| **v4.4** | Oct 2025 | Limpieza final, mejora de duplicados y .gitignore.           |
| **v4.5** | Oct 2025 | Unificación de carga PDF+imágenes y footer móvil responsive. |

