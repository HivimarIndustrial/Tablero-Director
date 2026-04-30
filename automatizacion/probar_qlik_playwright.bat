@echo off
REM Lanzador del smoke test de Qlik con Playwright (navegador VISIBLE).
setlocal
cd /d "%~dp0"
set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
%PY% probar_qlik_playwright.py
pause
