@echo off
REM ============================================================
REM  PDF -> Word PRO - arranque en Windows
REM  El entorno y los datos viven en %LOCALAPPDATA%\PDFWordPRO,
REM  nunca en la carpeta de instalacion (Archivos de programa no
REM  permite escribir a un usuario normal).
REM ============================================================
setlocal
title PDF Word PRO
cd /d "%~dp0"

set "BASE=%LOCALAPPDATA%\PDFWordPRO"
set "VENV=%BASE%\venv"
set "PDFWORDPRO_DATA=%BASE%\data"
if not exist "%PDFWORDPRO_DATA%" mkdir "%PDFWORDPRO_DATA%"

REM --- localizar Python 3 ---
set "PYCMD="
py -3 --version >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD python --version >nul 2>nul && set "PYCMD=python"
if not defined PYCMD (
  echo.
  echo [ERROR] No se encuentra Python 3 en este equipo.
  echo Instalalo desde https://www.python.org marcando
  echo "Add python.exe to PATH" y vuelve a intentarlo.
  pause
  exit /b 1
)

REM --- crear el entorno la primera vez ---
if not exist "%VENV%\Scripts\python.exe" (
  echo Preparando PDF Word PRO por primera vez...
  echo (se crea el entorno en %BASE% - necesita internet)
  %PYCMD% -m venv "%VENV%"
  if errorlevel 1 (
    echo [ERROR] No se pudo crear el entorno virtual.
    pause
    exit /b 1
  )
  "%VENV%\Scripts\python.exe" -m pip install --quiet --upgrade pip
  "%VENV%\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias. Revisa tu conexion.
    pause
    exit /b 1
  )
)

REM --- si ya esta en marcha, solo abrir el navegador ---
curl -s http://localhost:8000/api/health >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  start "" http://localhost:8000
  exit /b 0
)

REM --- arrancar el servidor en su propia ventana minimizada ---
start "PDF Word PRO (servidor)" /min "%VENV%\Scripts\python.exe" -m uvicorn app.main:app --host 0.0.0.0 --port 8000

REM --- esperar a que responda y abrir el navegador ---
set /a n=0
:espera
timeout /t 1 /nobreak >nul
curl -s http://localhost:8000/api/health >nul 2>nul
if %ERRORLEVEL% EQU 0 goto listo
set /a n+=1
if %n% LSS 90 goto espera
echo El servidor esta tardando en arrancar. Revisa la ventana minimizada
echo "PDF Word PRO (servidor)" por si muestra algun error.
pause
exit /b 1

:listo
start "" http://localhost:8000
echo.
echo PDF Word PRO esta en marcha en http://localhost:8000
echo El servidor sigue en la ventana minimizada "PDF Word PRO (servidor)":
echo cierrala para detener la aplicacion. Ya puedes cerrar esta ventana.
timeout /t 6 /nobreak >nul
exit /b 0
