@echo off
setlocal EnableDelayedExpansion
title Instalador — Separador de ASOs

echo.
echo ============================================================
echo  INSTALADOR — SEPARADOR DE ASOs
echo  Este script instala tudo que o programa precisa
echo ============================================================
echo.

set PY_CMD=py
py --version >nul 2>&1
if errorlevel 1 (
    set PY_CMD=python
)

echo [1/4] Verificando Python com comando: %PY_CMD%
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERRO: Python nao encontrado.
    echo Por favor instale o Python: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo Python OK.
echo.

echo [2/4] Instalando bibliotecas Python...
%PY_CMD% -m pip install --upgrade pip --quiet
%PY_CMD% -m pip install pdfplumber pypdf pdf2image pytesseract openpyxl pillow pyinstaller --quiet
echo Bibliotecas OK.
echo.

echo [3/4] Verificando Tesseract e Poppler...
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo Tesseract OK.
) else (
    echo AVISO: Tesseract nao instalado em C:\Program Files\Tesseract-OCR.
)

where pdftoppm >nul 2>&1
if errorlevel 1 (
    if exist "C:\poppler\bin\pdftoppm.exe" (
        echo Poppler em C:\poppler\bin.
    ) else (
        echo AVISO: Poppler nao encontrado no PATH.
    )
) else (
    echo Poppler OK.
)
echo.

echo [4/4] Criando atalho ABRIR_SEPARADOR_ASOs.bat...
set SCRIPT_DIR=%~dp0
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%
set APP_PATH=%SCRIPT_DIR%\SEPARACAODEASO.py

echo ============================================================
echo  INSTALACAO CONCLUIDA!
echo  Use ABRIR_SEPARADOR_ASOs.bat para iniciar o programa.
echo ============================================================
echo.
set /p INICIAR="Deseja abrir o programa agora? (S/N): "
if /i "!INICIAR!"=="S" (
    %PY_CMD% "%APP_PATH%"
)

endlocal
