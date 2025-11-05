"""
------------------------------------------------------------
PANEL DE CONTROL ANS – ELITE Ingenieros S.A.S.
------------------------------------------------------------
Autor: Héctor A. Gaviria + IA (2025)
------------------------------------------------------------
"""

import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox as mbox
from PIL import Image, ImageTk
import sys
import io
from datetime import datetime

# ------------------------------------------------------------
# CONFIGURACIÓN UTF-8 GLOBAL
# ------------------------------------------------------------
if sys.stdout.encoding is None or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
if sys.stderr.encoding is None or sys.stderr.encoding.lower() != "utf-8":
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

# ------------------------------------------------------------
# RUTA DE ARCHIVOS
# ------------------------------------------------------------
RUTA_LOGO = r"data_raw/elite.png"
RUTA_SCRIPT_ANS = r"calculos_ans.py"
RUTA_SCRIPT_LIMPIEZA = r"limpieza_fenix.py"

# ------------------------------------------------------------
# FUNCIONES DE INTERFAZ
# ------------------------------------------------------------
def resaltar_boton(boton):
    color_original = boton.cget("bg")
    boton.config(bg="#27AE60")
    ventana.update_idletasks()
    return color_original

def restaurar_boton(boton, color_original):
    boton.config(bg=color_original)
    ventana.update_idletasks()

# ------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# ------------------------------------------------------------
def ejecutar_comando(nombre, comando, boton=None):
    """Ejecuta un script externo mostrando logs y progreso animado."""
    def tarea():
        log_text.insert(tk.END, f"\n🚀 Iniciando {nombre}...\n", "info")
        log_text.see(tk.END)

        # Reiniciar barra
        barra_progreso["value"] = 0
        ventana.update_idletasks()

        hora = datetime.now().strftime("%I:%M %p")
        pie_estado.config(text=f"🔄 Procesando {nombre}... | {hora}", fg="#1A5276")
        ventana.update_idletasks()

        color_original = resaltar_boton(boton) if boton else None

        try:
            # Activar animación continua
            barra_progreso.config(mode="indeterminate")
            barra_progreso.start(20)

            proceso = subprocess.Popen(
                comando,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=1,
                universal_newlines=True,
                cwd=os.path.dirname(os.path.abspath(__file__)),
                encoding="utf-8"
            )

            for linea in iter(proceso.stdout.readline, ''):
                if not linea:
                    break
                log_text.insert(tk.END, linea)
                log_text.see(tk.END)
                ventana.update_idletasks()

            proceso.wait()

            barra_progreso.stop()
            barra_progreso.config(mode="determinate")

            if proceso.returncode == 0:
                barra_progreso["value"] = 100
                ventana.update_idletasks()
                log_text.insert(tk.END, f"\n✅ {nombre} completado con éxito.\n", "success")
                pie_estado.config(text=f"✅ {nombre} completado con éxito. | {hora}", fg="#27AE60")
            else:
                log_text.insert(tk.END, f"\n❌ Error en {nombre} (código {proceso.returncode}).\n", "error")
                pie_estado.config(text=f"⚠️ Error en {nombre}. Revisa el log.", fg="#C0392B")

        except Exception as e:
            barra_progreso.stop()
            barra_progreso.config(mode="determinate", value=100)
            log_text.insert(tk.END, f"\n⚠️ Error inesperado: {e}\n", "error")
            pie_estado.config(text=f"⚠️ Error en {nombre}. Revisa el log.", fg="#C0392B")

        finally:
            if boton and color_original:
                restaurar_boton(boton, color_original)
            log_text.insert(tk.END, "-" * 60 + "\n", "separador")
            log_text.see(tk.END)
            pie_estado.config(text="⚙️ Esperando acción del usuario...", fg="#1B263B")
            ventana.update_idletasks()
            ventana.after(1500, lambda: barra_progreso.config(value=0))

    threading.Thread(target=tarea, daemon=True).start()

# ------------------------------------------------------------
# COMANDO DE BOTÓN INFORME – SECUENCIA COMPLETA SEGURA Y FUNCIONAL
# ------------------------------------------------------------
def ejecutar_informe():
    """Ejecuta limpieza_fenix.py y luego calculos_ans.py en secuencia, usando el mismo Python activo."""
    def tarea():
        try:
            log_text.insert(tk.END, "\n🚀 Iniciando proceso completo Informe ANS...\n", "info")
            log_text.see(tk.END)
            ventana.update_idletasks()

            color_original = resaltar_boton(btn_informe)
            barra_progreso.config(mode="indeterminate")
            barra_progreso.start(20)

            # ✅ Detectar Python actual (entorno activo)
            python_exe = sys.executable
            base_dir = os.path.dirname(os.path.abspath(__file__))

            log_text.insert(tk.END, f"🔍 Python detectado: {python_exe}\n", "info")
            log_text.insert(tk.END, f"📂 Directorio base: {base_dir}\n\n", "info")

            # 1️⃣ Ejecutar limpieza_fenix.py
            log_text.insert(tk.END, "📂 Ejecutando limpieza de FENIX...\n", "info")
            proceso1 = subprocess.Popen(
                [python_exe, "-X", "utf8", os.path.join(base_dir, RUTA_SCRIPT_LIMPIEZA)],
                cwd=base_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="ignore"
            )
            for linea in iter(proceso1.stdout.readline, ''):
                log_text.insert(tk.END, linea)
                log_text.see(tk.END)
                ventana.update_idletasks()
            proceso1.wait()

            if proceso1.returncode != 0:
                log_text.insert(tk.END, "\n❌ Error en limpieza FENIX.\n", "error")
                pie_estado.config(text="⚠️ Error en limpieza FENIX", fg="#C0392B")
                return

            log_text.insert(tk.END, "✅ Limpieza completada correctamente.\n\n", "success")

            # 2️⃣ Ejecutar calculos_ans.py
            log_text.insert(tk.END, "📊 Ejecutando cálculos ANS...\n", "info")
            proceso2 = subprocess.Popen(
                [python_exe, "-X", "utf8", os.path.join(base_dir, RUTA_SCRIPT_ANS)],
                cwd=base_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="ignore"
            )
            for linea in iter(proceso2.stdout.readline, ''):
                log_text.insert(tk.END, linea)
                log_text.see(tk.END)
                ventana.update_idletasks()
            proceso2.wait()

            if proceso2.returncode == 0:
                log_text.insert(tk.END, "\n✅ Informe ANS generado correctamente.\n", "success")
                pie_estado.config(text="✅ Informe ANS actualizado correctamente.", fg="#27AE60")
                
                # 🟢 Nuevo popup de confirmación visual
                mbox.showinfo("Control ANS – ELITE Ingenieros S.A.S.",
                              "✅ El Informe ANS ha sido actualizado correctamente.\n\n"
                              "El archivo FENIX_ANS.xlsx está listo para revisión.")
            else:
                log_text.insert(tk.END, "\n❌ Error en cálculos ANS.\n", "error")
                pie_estado.config(text="⚠️ Error en cálculos ANS", fg="#C0392B")

        except Exception as e:
            log_text.insert(tk.END, f"\n⚠️ Error inesperado: {e}\n", "error")
            pie_estado.config(text="⚠️ Error inesperado", fg="#C0392B")

        finally:
            barra_progreso.stop()
            barra_progreso.config(mode="determinate", value=100)
            restaurar_boton(btn_informe, color_original)
            log_text.insert(tk.END, "-" * 60 + "\n", "separador")
            pie_estado.config(text="⚙️ Esperando acción del usuario...", fg="#1B263B")
            ventana.after(1500, lambda: barra_progreso.config(value=0))

    threading.Thread(target=tarea, daemon=True).start()

# ------------------------------------------------------------
# INTERFAZ PRINCIPAL
# ------------------------------------------------------------
ventana = tk.Tk()
ventana.title("Control ANS - ELITE Ingenieros S.A.S.")
ventana.config(bg="#EAEDED")

# ------------------------------------------------------------
# BARRA SUPERIOR CON RELOJ VERDE
# ------------------------------------------------------------
frame_topbar = tk.Frame(ventana, bg="#1E8449", height=22)
frame_topbar.pack(fill="x")

reloj_top = tk.Label(
    frame_topbar,
    font=("Segoe UI", 9, "bold"),
    fg="white",
    bg="#1E8449",
    anchor="e"
)
reloj_top.pack(side="right", padx=15)

def actualizar_hora_top():
    hora_actual = datetime.now().strftime("%I:%M:%S %p")
    reloj_top.config(text=f"{hora_actual}")
    ventana.after(1000, actualizar_hora_top)

actualizar_hora_top()

# ------------------------------------------------------------
# TAMAÑO Y CENTRADO
# ------------------------------------------------------------
screen_w = ventana.winfo_screenwidth()
screen_h = ventana.winfo_screenheight()
ancho = int(screen_w * 0.55)
alto = int(screen_h * 0.78)
x = (screen_w // 2) - (ancho // 2)
y = (screen_h // 2) - (alto // 2)
ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
ventana.resizable(False, False)

# ------------------------------------------------------------
# ENCABEZADO PROFESIONAL
# ------------------------------------------------------------
frame_banner = tk.Frame(ventana, bg="#EAEDED", height=120)
frame_banner.pack(fill="x")

frame_superior = tk.Frame(frame_banner, bg="#EAEDED")
frame_superior.pack(pady=(10, 0))

try:
    logo_img = Image.open(RUTA_LOGO)
    logo_img = logo_img.resize((70, 70), Image.Resampling.LANCZOS)
    logo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(frame_superior, image=logo, bg="#EAEDED")
    logo_label.pack(side="left", padx=15)
except Exception:
    logo_label = tk.Label(frame_superior, text="[Logo no encontrado]", fg="black", bg="#EAEDED", font=("Segoe UI", 10))
    logo_label.pack(side="left", padx=15)

elite_label = tk.Label(frame_superior, text="ELITE ", font=("Segoe UI", 18, "bold"), fg="black", bg="#EAEDED")
elite_label.pack(side="left")

ingenieros_label = tk.Label(frame_superior, text="Ingenieros S.A.S.", font=("Segoe UI", 18, "bold"), fg="#1E8449", bg="#EAEDED")
ingenieros_label.pack(side="left")

titulo_control = tk.Label(frame_banner, text="Control ANS", font=("Segoe UI", 14, "bold"), fg="#1B263B", bg="#EAEDED")
titulo_control.pack(pady=(0, 10))

# ------------------------------------------------------------
# BOTONES PRINCIPALES – 1 SOLA FILA
# ------------------------------------------------------------
frame_botones = tk.Frame(ventana, bg="#EAEDED")
frame_botones.pack(pady=5, fill="x")
frame_botones.columnconfigure((0, 1, 2, 3), weight=1)

# Botón 1: EJECUTAR INFORME ANS
btn_informe = tk.Button(frame_botones, text="EJECUTAR\nINFORME ANS", command=ejecutar_informe,
                        width=20, height=2, bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"),
                        relief="ridge", borderwidth=3, cursor="hand2",
                        activebackground="#229954", activeforeground="white")
btn_informe.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

# Botón 2: CONTROL ALMACÉN
RUTA_SCRIPT_VALIDACION = r"validar_export_almacen.py"

def ejecutar_validacion():
    comando = f'python -X utf8 "{RUTA_SCRIPT_VALIDACION}"'
    ejecutar_comando("Control Almacén ANS", comando, btn_validar)

btn_validar = tk.Button(frame_botones, text="CONTROL\nFENIX Vs ALMACÉN", command=ejecutar_validacion,
                        width=20, height=2, bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"),
                        relief="ridge", borderwidth=3, cursor="hand2",
                        activebackground="#229954", activeforeground="white")
btn_validar.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

# Botón 3: DESCARGAR EVIDENCIAS DRIVE
RUTA_SCRIPT_DESCARGA = r"descargar_drive_v48.py"

def ejecutar_descarga_drive():
    comando = f'python -X utf8 "{RUTA_SCRIPT_DESCARGA}"'
    ejecutar_comando("Descarga Evidencias Drive", comando, btn_descarga_drive)

btn_descarga_drive = tk.Button(frame_botones, text="DESCARGAR\nEVIDENCIAS DRIVE", command=ejecutar_descarga_drive,
                               width=20, height=2, bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"),
                               relief="ridge", borderwidth=3, cursor="hand2",
                               activebackground="#1E8449", activeforeground="white")
btn_descarga_drive.grid(row=0, column=2, padx=10, pady=5, sticky="ew")

# Botón 4: MOVER A PAPELERA_API
RUTA_SCRIPT_PAPELERA = r"descargar_evidencias_drive.py"

def ejecutar_papelera_drive():
    comando = f'python -X utf8 "{RUTA_SCRIPT_PAPELERA}"'
    ejecutar_comando("Mover Evidencias a PAPELERA_API", comando, btn_papelera_drive)

btn_papelera_drive = tk.Button(frame_botones, text="MOVER A\nPAPELERA API", command=ejecutar_papelera_drive,
                               width=20, height=2, bg="#C0392B", fg="white", font=("Segoe UI", 10, "bold"),
                               relief="ridge", borderwidth=3, cursor="hand2",
                               activebackground="#922B21", activeforeground="white")
btn_papelera_drive.grid(row=0, column=3, padx=10, pady=5, sticky="ew")

# ------------------------------------------------------------
# BARRA DE PROGRESO
# ------------------------------------------------------------
barra_progreso = ttk.Progressbar(ventana, orient="horizontal", mode="determinate", length=450, maximum=100)
barra_progreso.pack(pady=(5, 5))

# ------------------------------------------------------------
# ÁREA DE LOG
# ------------------------------------------------------------
frame_log = tk.Frame(ventana, bg="#EAEDED")
frame_log.pack(fill="both", expand=False, pady=(5, 0))

log_text = scrolledtext.ScrolledText(frame_log, width=90, height=14, bg="white", font=("Consolas", 9))
log_text.pack(padx=15, pady=(5, 10), expand=True, fill="both")

log_text.tag_config("info", foreground="#2980B9")
log_text.tag_config("success", foreground="#27AE60")
log_text.tag_config("error", foreground="#C0392B")
log_text.tag_config("separador", foreground="#95A5A6")

# ------------------------------------------------------------
# BOTÓN SALIR
# ------------------------------------------------------------
frame_salida = tk.Frame(ventana, bg="#EAEDED")
frame_salida.pack(pady=(0, 10))
btn_salir = tk.Button(frame_salida, text="SALIR DEL PANEL", command=ventana.quit,
                      width=25, height=2, bg="#1E8449", fg="white",
                      font=("Segoe UI", 10, "bold"), relief="ridge", borderwidth=3, cursor="hand2",
                      activebackground="#C0392B", activeforeground="white")
btn_salir.pack(pady=3)

# ------------------------------------------------------------
# PIE DE PÁGINA
# ------------------------------------------------------------
frame_footer = tk.Frame(ventana, bg="#EAEDED")
frame_footer.pack(side="bottom", fill="x", pady=(0, 5), ipady=4)
tk.Frame(frame_footer, bg="#B3B6B7", height=2).pack(fill="x", pady=(2, 3))

frame_pie = tk.Frame(frame_footer, bg="#EAEDED")
frame_pie.pack(fill="x", pady=(0, 3))

pie_estado = tk.Label(frame_pie, text="\u2699 Esperando acción del usuario...",
                      font=("Segoe UI", 9, "italic"), fg="#1B263B", bg="#EAEDED", anchor="w")
pie_estado.pack(side="left", padx=(15, 0))

pie_corporativo = tk.Label(frame_pie,
    text="© 2025 ELITE Ingenieros S.A.S.  |  Pasión por lo que hacemos.",
    font=("Segoe UI", 9, "italic"), fg="#1B263B", bg="#EAEDED", anchor="e")
pie_corporativo.pack(side="right", padx=(0, 15))

# ------------------------------------------------------------
# INICIAR INTERFAZ
# ------------------------------------------------------------
ventana.mainloop()
