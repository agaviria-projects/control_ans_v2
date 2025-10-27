from flask import Flask, render_template, request, jsonify, redirect, url_for, flash
from datetime import datetime
from pathlib import Path
import pandas as pd
import os

# ------------------------------------------------------------
# CONFIGURACIÓN BASE DE FLASK
# ------------------------------------------------------------
base_dir = Path(__file__).resolve().parent  # carpeta 'formularios_tecnicos'

app = Flask(__name__, static_url_path='/static', static_folder='static', template_folder='templates')

# Clave secreta para mensajes flash
app.secret_key = "clave_super_secreta_ans"

# Carpeta de cargas
app.config['UPLOAD_FOLDER'] = base_dir / "static" / "uploads"
app.config['UPLOAD_FOLDER'].mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------
# CARGA ARCHIVO FENIX
# ------------------------------------------------------------
ruta_fenix = base_dir.parent / "data_clean" / "FENIX_ANS.xlsx"
if ruta_fenix.exists():
    df_fenix = pd.read_excel(ruta_fenix)
    df_fenix.columns = df_fenix.columns.str.strip().str.upper()
else:
    df_fenix = pd.DataFrame()

# ------------------------------------------------------------
# FORMULARIO PRINCIPAL
# ------------------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def formulario():
    ruta_excel = base_dir / "registros_formulario.xlsx"
    df_registros = pd.read_excel(ruta_excel) if ruta_excel.exists() else pd.DataFrame()

    # Si es envío del formulario (POST)
    if request.method == "POST":
        pedido = str(request.form["pedido"]).strip()
        observacion = request.form["observacion"]
        estado = request.form["estado"]

        # 🔸 Validar duplicado
        if not df_registros.empty and pedido in df_registros["pedido"].astype(str).values:
            flash(f"⚠ El pedido {pedido} ya fue registrado anteriormente.", "warning")
            return redirect(url_for("formulario"))

        # 🔸 Validar existencia en FENIX
        df_fenix["PEDIDO"] = df_fenix["PEDIDO"].astype(str).str.strip()
        resultado = df_fenix[df_fenix["PEDIDO"] == pedido]

        if resultado.empty:
            flash(f"❌ Pedido {pedido} no existe en FENIX_ANS. Verifique nuevamente.", "danger")
            return redirect(url_for("formulario"))

        # 🔸 Procesar archivos
        archivo_pdf = request.files.get("archivo_pdf")
        nombre_pdf = None
        if archivo_pdf and archivo_pdf.filename:
            nombre_pdf = f"{pedido}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            archivo_pdf.save(app.config['UPLOAD_FOLDER'] / nombre_pdf)

        imagenes = request.files.getlist("imagenes")
        nombres_imagenes = []
        for i, imagen in enumerate(imagenes, start=1):
            if imagen.filename:
                ext = imagen.filename.split(".")[-1].lower()
                if ext in ["jpg", "jpeg", "png"]:
                    nombre_img = f"{pedido}_{i}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
                    imagen.save(app.config['UPLOAD_FOLDER'] / nombre_img)
                    nombres_imagenes.append(nombre_img)

        # 🔸 Registrar fila
        fila = resultado.iloc[0]
        registro = {
            "fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "pedido": pedido,
            "clienteid": fila.get("CLIENTEID", ""),
            "cliente": fila.get("NOMBRE_CLIENTE", ""),
            "direccion": fila.get("DIRECCION", ""),
            "observacion": observacion,
            "estado_campo": estado,
            "metodo_envio": request.form.get("metodo_envio", ""),
            "estado_fenix": fila.get("ESTADO", ""),
            "pdf": nombre_pdf or "Sin archivo",
            "imagenes": ", ".join(nombres_imagenes) if nombres_imagenes else "Sin imágenes"
        }

        # 🔸 Guardar registro
        df_final = pd.concat([df_registros, pd.DataFrame([registro])], ignore_index=True)
        df_final.to_excel(ruta_excel, index=False)

        # 🔸 Confirmar al usuario
        flash(f"✅ Registro guardado correctamente — Pedido {pedido}", "success")
        return redirect(url_for("formulario"))

    # ✅ Si es GET (abrir formulario)
    return render_template("form.html")

# ------------------------------------------------------------
# CONSULTA PEDIDO FENIX
# ------------------------------------------------------------
@app.route("/buscar_pedido/<pedido_id>")
def buscar_pedido(pedido_id):
    pedido_id = str(pedido_id).strip()
    if df_fenix.empty:
        return jsonify({"error": "Archivo FENIX_ANS no encontrado o vacío"})

    ruta_excel = base_dir / "registros_formulario.xlsx"
    if ruta_excel.exists():
        df_reg = pd.read_excel(ruta_excel)
        df_reg["pedido"] = df_reg["pedido"].astype(str)
        if pedido_id in df_reg["pedido"].values:
            return jsonify({"mensaje_duplicado": f"⚠ El pedido {pedido_id} ya fue reportado como Cumplido."})

    df_fenix["PEDIDO"] = df_fenix["PEDIDO"].astype(str).str.strip()
    resultado = df_fenix[df_fenix["PEDIDO"] == pedido_id]
    if not resultado.empty:
        fila = resultado.iloc[0]
        datos = {
            "clienteid": str(fila.get("CLIENTEID", "")),
            "nombre_cliente": str(fila.get("NOMBRE_CLIENTE", "")),
            "telefono": str(fila.get("TELEFONO_CONTACTO", "")),
            "celular": str(fila.get("CELULAR_CONTACTO", "")),
            "direccion": str(fila.get("DIRECCION", "")),
            "fecha_limite_ans": str(fila.get("FECHA_LIMITE_ANS", "")),
            "estado": str(fila.get("ESTADO", ""))
        }
        return jsonify(datos)
    else:
        return jsonify({"error": "Pedido no encontrado"})

# ------------------------------------------------------------
# EJECUCIÓN
# ------------------------------------------------------------
if __name__ == "__main__":
    print("Ruta absoluta esperada del static:")
    print(app.static_folder)
    print("Contenido real de esa carpeta:")
    print(os.listdir(app.static_folder))
    app.run(debug=True, host="0.0.0.0")
