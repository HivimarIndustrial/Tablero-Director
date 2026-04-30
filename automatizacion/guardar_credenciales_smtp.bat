@echo off
REM Guarda las credenciales SMTP (Office 365) en Windows Credential Manager.
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

%PY% guardar_credenciales_smtp.py
echo.
pause
