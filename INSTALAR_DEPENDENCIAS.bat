@echo off
setlocal EnableDelayedExpansion
title Instalador de Dependencias — Separador de ASOs

echo.
echo ============================================================
echo  INSTALADOR DE DEPENDENCIAS
echo  Separador Automatico de ASOs
echo ============================================================
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo ERRO: Execute como Administrador (clique com botao direito ^> Executar como administrador).
    pause
    exit /b 1
)

set PASTA_INST=%~dp0dependencias_temp
mkdir "%PASTA_INST%" >nul 2>&1

echo [1/2] Verificando Tesseract OCR...
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo Tesseract ja esta instalado.
    goto :POPPLER
)

echo Baixando Tesseract OCR...
powershell -Command "& { $url='https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.4.20240503/tesseract-ocr-w64-setup-5.3.4.20240503.exe'; $dest='%PASTA_INST%\tesseract_setup.exe'; Write-Host '  Baixando...'; Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing; Write-Host '  Download concluido.' }"

if exist "%PASTA_INST%\tesseract_setup.exe" (
    echo Instalando Tesseract OCR...
    "%PASTA_INST%\tesseract_setup.exe" /S /D=C:\Program Files\Tesseract-OCR
)

:POPPLER
echo.
echo [2/2] Verificando Poppler...
where pdftoppm >nul 2>&1
if not errorlevel 1 (
    echo Poppler ja esta no PATH.
    goto :FINALIZAR
)

if exist "C:\poppler\Library\bin\pdftoppm.exe" goto :ADD_PATH

echo Baixando Poppler...
powershell -Command "& { $url='https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip'; $dest='%PASTA_INST%\poppler.zip'; Write-Host '  Baixando...'; Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing; Write-Host '  Download concluido.' }"

if exist "%PASTA_INST%\poppler.zip" (
    echo Extraindo Poppler...
    powershell -Command "Expand-Archive -Path '%PASTA_INST%\poppler.zip' -DestinationPath 'C:\poppler\' -Force"
)

:ADD_PATH
set POPPLER_BIN=C:\poppler\Library\bin
if exist "%POPPLER_BIN%\pdftoppm.exe" (
    echo Adicionando Poppler ao PATH...
    powershell -Command "& { $path = [System.Environment]::GetEnvironmentVariable('Path','Machine'); if ($path -notlike '*C:\poppler\*') { [System.Environment]::SetEnvironmentVariable('Path', $path + ';C:\poppler\Library\bin', 'Machine'); } }"
)

:FINALIZAR
rmdir /s /q "%PASTA_INST%" >nul 2>&1
echo.
echo ============================================================
echo  DEPENDENCIAS CONFIGURADAS!
echo ============================================================
echo.
pause
endlocal
