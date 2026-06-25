@echo off
setlocal

if "%~1"=="" (
    echo Selecciona uno o varios archivos, o una carpeta.
    exit /b 1
)

python "%~dp0format_report.py" %*
set "exit_code=%errorlevel%"

if not "%exit_code%"=="0" (
    echo.
    echo Ocurrio un error al formatear uno o mas archivos.
)

exit /b %exit_code%
