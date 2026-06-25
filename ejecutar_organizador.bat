@echo off
chcp 65001 >nul
title Organizador de Carpetas Educativas - CDCOMPRI1P

echo.
echo ============================================================
echo    ORGANIZADOR DE CARPETAS EDUCATIVAS - CDCOMPRI1P
echo ============================================================
echo.

REM ── Verificar que Python este instalado ──────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo.
    echo  Descargalo desde: https://www.python.org/downloads/
    echo  Asegurate de marcar "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)

echo [1/3] Python detectado:
python --version
echo.

REM ── Instalar dependencias si hacen falta ─────────────────────
echo [2/3] Verificando dependencias (pandas, openpyxl, beautifulsoup4)...
pip install pandas openpyxl beautifulsoup4 --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [WARN] No se pudieron instalar algunas dependencias.
    echo        Si el script falla, ejecuta manualmente:
    echo        pip install pandas openpyxl beautifulsoup4
)
echo.

REM ── Ejecutar el organizador ───────────────────────────────────
echo [3/3] Iniciando organizador...
echo ============================================================
echo.

python "%~dp0edu_folder_organizer.py"

echo.
if errorlevel 1 (
    echo [ERROR] El script termino con errores. Revisa los mensajes arriba.
) else (
    echo [OK] Proceso finalizado correctamente.
    echo      Revisa la carpeta OUTPUT en:
    echo      C:\Users\bangulo\Downloads\CDCOMPRI1P\OUTPUT
)

echo.
pause
