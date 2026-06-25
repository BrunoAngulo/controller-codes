@echo off
chcp 65001 >nul
title Instalar - Organizar carpeta educativa

set "SCRIPT_DIR=%~dp0"
set "BAT_FILE=%SCRIPT_DIR%ejecutar_organizador.bat"
set "SENDTO_DIR=%APPDATA%\Microsoft\Windows\SendTo"
set "REG_KEY=OrganizarCarpetaEducativa"
set "MENU_LABEL=Organizar carpeta educativa"

echo.
echo ============================================================
echo   INSTALACION: Organizar carpeta educativa
echo ============================================================
echo.
echo   Script: %BAT_FILE%
echo.

REM ── [0] Verificar Python y dependencias ──────────────────────
echo [0/3] Verificando Python y dependencias...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Python no instalado.
    echo           Descargalo en: https://www.python.org/downloads/
    pause
    exit /b 1
)
pip install pandas openpyxl beautifulsoup4 --quiet --disable-pip-version-check
echo   [OK] Dependencias instaladas.
echo.

REM ── [1] Acceso directo en "Enviar a" ─────────────────────────
echo [1/3] Creando acceso directo en menu "Enviar a"...
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut('%SENDTO_DIR%\%MENU_LABEL%.lnk'); ^
   $s.TargetPath = '%BAT_FILE%'; ^
   $s.WorkingDirectory = '%SCRIPT_DIR%'; ^
   $s.Description = 'Organizar carpetas educativas'; ^
   $s.Save()" >nul 2>&1

if exist "%SENDTO_DIR%\%MENU_LABEL%.lnk" (
    echo   [OK] Aparecera en: clic derecho → "Enviar a" → "%MENU_LABEL%"
) else (
    echo   [WARN] No se pudo crear en "Enviar a".
)
echo.

REM ── [2] Menu contextual al hacer clic derecho en carpeta ──────
echo [2/3] Registrando en menu contextual (clic derecho sobre carpeta)...
reg add "HKCU\Software\Classes\Directory\shell\%REG_KEY%" /ve /d "%MENU_LABEL%" /f >nul 2>&1
reg add "HKCU\Software\Classes\Directory\shell\%REG_KEY%\command" /ve /d "\"%BAT_FILE%\" \"%%1\"" /f >nul 2>&1

reg query "HKCU\Software\Classes\Directory\shell\%REG_KEY%" >nul 2>&1
if errorlevel 1 (
    echo   [WARN] No se pudo registrar el menu contextual.
) else (
    echo   [OK] Aparecera en: clic derecho en carpeta → "%MENU_LABEL%"
)
echo.

REM ── [3] Menu al hacer clic derecho en fondo de carpeta ───────
echo [3/3] Registrando para fondo de carpeta...
reg add "HKCU\Software\Classes\Directory\Background\shell\%REG_KEY%" /ve /d "%MENU_LABEL%" /f >nul 2>&1
reg add "HKCU\Software\Classes\Directory\Background\shell\%REG_KEY%\command" /ve /d "\"%BAT_FILE%\" \"%%V\"" /f >nul 2>&1
echo   [OK] Registrado.
echo.

echo ============================================================
echo   INSTALACION COMPLETADA
echo ============================================================
echo.
echo   Formas de usar:
echo.
echo   A) Arrastra carpetas sobre "ejecutar_organizador.bat"
echo      (admite multiples carpetas a la vez)
echo.
echo   B) Selecciona carpetas → clic derecho → "Enviar a"
echo      → "%MENU_LABEL%"
echo      (se procesa una por una, cada una crea su OUTPUT)
echo.
echo   C) Clic derecho en una carpeta
echo      → "%MENU_LABEL%"
echo.
echo   NOTA: Si mueves los archivos .bat y .py a otra carpeta,
echo   ejecuta este instalador nuevamente.
echo.
pause
