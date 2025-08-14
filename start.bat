@echo off
ECHO Instalando/verificando dependencias desde requirements.txt...
pip install -r requirements.txt

ECHO.
ECHO Iniciando la aplicacion...
python gui_app.py

pause
