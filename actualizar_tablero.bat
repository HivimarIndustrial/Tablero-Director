@echo off
REM =======================================================================
REM   Hivimar Tablero Industrial - actualizar desde Base Tablero Industria.xlsx
REM   Uso: doble clic, o "actualizar_tablero.bat" en una terminal.
REM   NO descarga nada de QLIK/Odoo todavia: eso es manual por ahora.
REM =======================================================================

setlocal
cd /d "%~dp0"

REM Ruta al Python instalado para el usuario consultorindustrial
set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

if not exist %PY% (
    echo ERROR: No se encontro Python en %PY%
    echo Instala Python 3.12 desde winget: winget install Python.Python.3.12 --scope user
    pause
    exit /b 1
)

echo.
echo === 1/3  Respaldo del HTML actual ===
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value ^| find "="') do set DT=%%a
set BACKUP=Tablero_Director_FUENTE.backup_%DT:~0,8%_%DT:~8,6%.html
copy /Y "Tablero_Director_FUENTE.html" "%BACKUP%" >nul
echo   Respaldo: %BACKUP%

echo.
echo === 2/3  Regenerando DB desde el Excel ===
%PY% regenerar_db.py
if errorlevel 1 (
    echo.
    echo ERROR ejecutando regenerar_db.py. Abortando.
    pause
    exit /b 1
)

echo.
echo === 3/3  Inyectando DB en el HTML ===
%PY% update_html.py
if errorlevel 1 (
    echo.
    echo ERROR ejecutando update_html.py. Abortando.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Listo. Abre: Tablero_Director_FUENTE.html
echo ============================================================
pause
