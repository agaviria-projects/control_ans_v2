@echo off
:: ============================================================
:: DESCARGAR EVIDENCIAS DRIVE (Automático cada 2 horas)
:: ============================================================

cd "C:\Users\hector.gaviria\Desktop\Control_ANS"

:: Activar entorno virtual
call "C:\Users\hector.gaviria\Desktop\Control_ANS\venv\Scripts\activate.bat"

:: Registrar fecha y hora de inicio en log
echo ============================================================ >> logs_descargas.txt
echo 🕒 Inicio de ejecución: %date% %time% >> logs_descargas.txt

:: Ejecutar script Python y registrar salida
python descargar_evidencias_drive.py >> logs_descargas.txt 2>&1

:: Registrar fin de ejecución
echo ✅ Fin de ejecución: %date% %time% >> logs_descargas.txt
echo. >> logs_descargas.txt

:: Desactivar entorno virtual
deactivate
exit
