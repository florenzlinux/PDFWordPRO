@echo off
REM ============================================================
REM  PDF -> Word PRO : compila pdfwordpro.msi en Windows
REM  Requisitos:
REM    1) WiX Toolset v4  ->  dotnet tool install --global wix
REM       (o el instalador de https://wixtoolset.org)
REM    2) Este archivo junto a pdfwordpro.wxs y la carpeta app\
REM  Resultado: pdfwordpro.msi junto a este archivo.
REM ============================================================
setlocal
cd /d "%~dp0"

where wix >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo [ERROR] No se encuentra WiX Toolset v4 ^(comando "wix"^).
  echo Instala .NET SDK y despues ejecuta:
  echo     dotnet tool install --global wix
  echo O descarga WiX desde https://wixtoolset.org
  pause
  exit /b 1
)

echo Compilando pdfwordpro.msi ...
wix build pdfwordpro.wxs -o pdfwordpro.msi
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo [ERROR] La compilacion ha fallado. Revisa los mensajes de WiX.
  pause
  exit /b 1
)

echo.
echo OK: pdfwordpro.msi creado.
echo El instalador deja la app en Archivos de programa\PDFWordPRO.
echo Para usarla: doble clic en iniciar_windows.bat de esa carpeta
echo (requiere Python 3.11+ desde https://www.python.org, marcado
echo  "Add python.exe to PATH" al instalarlo).
pause
