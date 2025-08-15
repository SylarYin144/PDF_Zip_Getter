@echo off
REM Ejecuta el script de Python y redirige la salida estándar a stdout.log y el error estándar a stderr.log
python scihub_downloader.py > stdout.log 2> stderr.log

echo.
echo Proceso terminado.
echo Si la aplicacion no se abrio, por favor revisa el archivo stderr.log para ver los detalles del error.
echo.
pause
