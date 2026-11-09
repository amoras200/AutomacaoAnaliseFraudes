@echo off
rem Duplo clique para gerar o relatorio de chargeback (abre janelas para escolher
rem a planilha e o periodo). Tambem aceita argumentos: "Gerar Relatorio.bat" planilha.xlsx --inicio 2026-03
cd /d "%~dp0"

set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY (
    python --version >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo.
    echo [ERRO] Python nao foi encontrado neste computador.
    echo Instale em https://www.python.org/downloads/ e, na primeira tela do
    echo instalador, marque a opcao "Add python.exe to PATH". Depois de instalar,
    echo de duplo clique neste arquivo novamente.
    echo.
    pause
    exit /b 1
)

%PY% -c "import pandas, openpyxl, matplotlib" >nul 2>nul
if errorlevel 1 (
    echo Preparando o programa pela primeira vez. Isso leva 1 a 2 minutos
    echo e precisa de internet. Nao feche esta janela...
    echo.
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERRO] Nao foi possivel instalar as bibliotecas. Verifique a internet
        echo e tente novamente.
        pause
        exit /b 1
    )
)

%PY% relatorio_chargeback.py %*
if errorlevel 1 pause
