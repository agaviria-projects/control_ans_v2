@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python menu_control_ans.py
pause
