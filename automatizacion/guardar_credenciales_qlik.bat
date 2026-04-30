@echo off
REM Lanzador para guardar las credenciales de Qlik en Windows Credential Manager.
REM Doble clic y sigue las instrucciones en la ventana.

setlocal
cd /d "%~dp0"

set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

if not exist %PY% (
    echo ERROR: No se encontro Python. Reinstala con:
    echo   winget install Python.Python.3.12 --scope user
    pause
    exit /b 1
)

%PY% guardar_credenciales_qlik.py
echo.
pause
