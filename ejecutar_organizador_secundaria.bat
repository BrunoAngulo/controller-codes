@echo off
chcp 65001 >nul
title Organizador de Carpetas Educativas — Secundaria

REM ── Si no hay argumentos, mostrar instrucciones ───────────────
if "%~1"=="" (
    echo.
    echo  ============================================================
    echo    ORGANIZADOR DE CARPETAS EDUCATIVAS — SECUNDARIA
    echo  ============================================================
    echo.
    echo  Como usar:
    echo    1. Arrastra una o varias carpetas sobre ESTE archivo .bat
    echo    2. O instala el menu contextual con:
    echo       instalar_menu_contextual.bat
    echo.
    pause
    exit /b 0
)

REM ── Verificar Python ──────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERROR] Python no encontrado en el sistema.
    echo          Descargalo en: https://www.python.org/downloads/
    echo          Marca "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

REM ── Instalar dependencias silenciosamente ─────────────────────
pip install pandas openpyxl beautifulsoup4 --quiet --disable-pip-version-check 2>nul

REM ── Ejecutar con todas las carpetas recibidas ─────────────────
python "%~dp0edu_folder_organizer.py" --secundaria %*
