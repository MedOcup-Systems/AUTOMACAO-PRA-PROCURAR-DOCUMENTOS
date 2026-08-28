@echo off
setlocal EnableDelayedExpansion
title Gerar Executavel - Separador de ASOs

echo.
echo ============================================================
echo  GERADOR DE EXECUTAVEL - SEPARADOR DE ASOs
echo  Gera um .exe para distribuir sem precisar do Python
echo ============================================================
echo.
echo Este processo pode demorar alguns minutos.
echo NAO feche esta janela.
echo.

REM -- Pasta do script --
set SCRIPT_DIR=%~dp0
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%
set APP_PY=%SCRIPT_DIR%\SEPARACAODEASO.py

echo Pasta de trabalho: %SCRIPT_DIR%
echo Arquivo Python: %APP_PY%
echo.

if not exist "%APP_PY%" (
    echo ERRO: SEPARACAODEASO.py nao encontrado em: %SCRIPT_DIR%
    pause
    exit /b 1
)

REM -- Tenta py ou python --
set PY_CMD=py
py --version >nul 2>&1
if errorlevel 1 (
    set PY_CMD=python
)

echo [1/3] Verificando instalador Python com comando: %PY_CMD%
%PY_CMD% --version
echo.

echo [2/3] Instalando dependencias e PyInstaller se necessario...
%PY_CMD% -m pip install --upgrade pip --quiet
%PY_CMD% -m pip install pyinstaller pdfplumber pypdf pdf2image pytesseract openpyxl pillow --quiet
echo.

echo [3/3] Compilando executavel...
if exist "%SCRIPT_DIR%\dist" rmdir /s /q "%SCRIPT_DIR%\dist"
if exist "%SCRIPT_DIR%\build" rmdir /s /q "%SCRIPT_DIR%\build"

%PY_CMD% -m PyInstaller "%APP_PY%" --onefile --windowed --name "Separador_ASOs" --distpath "%SCRIPT_DIR%\dist" --workpath "%SCRIPT_DIR%\build" --noconfirm --hidden-import pdfplumber --hidden-import pdfminer --hidden-import pdfminer.six --hidden-import openpyxl --hidden-import openpyxl.cell --hidden-import openpyxl.styles --hidden-import pypdf --hidden-import pytesseract --hidden-import pdf2image --hidden-import PIL --hidden-import PIL.Image --hidden-import PIL._tkinter_finder --hidden-import unicodedata --hidden-import queue --hidden-import threading --collect-all pdfplumber --collect-all pdfminer --collect-all openpyxl

echo.
if exist "%SCRIPT_DIR%\dist\Separador_ASOs.exe" (
    echo ============================================================
    echo  EXECUTAVEL GERADO COM SUCESSO: dist\Separador_ASOs.exe
    echo ============================================================
    echo.
    set /p ABRIR="Deseja abrir a pasta com o executavel agora? (S/N): "
    if /i "!ABRIR!"=="S" explorer "%SCRIPT_DIR%\dist"
) else (
    echo ============================================================
    echo  ERRO: O executavel NAO foi gerado.
    echo ============================================================
    pause
)

endlocal
