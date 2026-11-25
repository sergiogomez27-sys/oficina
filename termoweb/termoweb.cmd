@echo off
rem Uso: termoweb ver | termoweb horario | termoweb ajustar   (anade --prueba para no cambiar nada)
cd /d "%~dp0"
python termoweb.py %*
if "%~1"=="" pause
