@echo off
rem Crea una tarea de Windows que ajusta las temperaturas segun el tiempo cada dia a las 6:00.
cd /d "%~dp0"
schtasks /Create /F /SC DAILY /ST 06:00 /TN "TermoWeb ajustar" /TR "\"%~dp0termoweb.cmd\" ajustar"
echo.
echo Hecho. Para quitarla: schtasks /Delete /TN "TermoWeb ajustar" /F
pause
