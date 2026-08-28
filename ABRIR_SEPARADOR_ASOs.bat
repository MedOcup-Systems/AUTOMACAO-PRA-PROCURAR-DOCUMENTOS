@echo off
title Separador de ASOs

set SCRIPT_DIR=%~dp0
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

py "%SCRIPT_DIR%\SEPARACAODEASO.py" >nul 2>&1
if not errorlevel 1 exit /b 0

python "%SCRIPT_DIR%\SEPARACAODEASO.py"
if errorlevel 1 (
    echo.
    echo ERRO ao iniciar. Execute primeiro: INSTALAR_E_CONFIGURAR.bat
    pause
)
