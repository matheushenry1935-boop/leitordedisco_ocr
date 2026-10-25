@echo off
rem ============================================================
rem  OCR MIDIAS JUDICIAIS V6 - Instalador de requisitos (Windows)
rem  Instala: Python 3.12 (se faltar) e as bibliotecas Python
rem  (em uma pasta .venv). A leitura dos discos e feita por IA
rem  via internet, entao nao precisa instalar Tesseract.
rem  Arquivo mantido sem acentos de proposito (compatibilidade).
rem ============================================================
setlocal EnableExtensions
title Instalador - OCR Midias Judiciais
cd /d "%~dp0"

echo ============================================================
echo   OCR MIDIAS JUDICIAIS V6 - Instalador de requisitos
echo ============================================================
echo.

if not exist "requirements.txt" (
    echo ERRO: requirements.txt nao foi encontrado. Deixe todos os arquivos na mesma pasta.
    goto :fim_erro
)

rem ---- 1. Python
echo [1/3] Verificando o Python...
call :achar_python
if not defined PYV (
    echo       Python 3.10 a 3.13 com Tkinter nao encontrado. Instalando o Python 3.12...
    where winget >nul 2>&1
    if errorlevel 1 (
        echo ERRO: o winget nao foi encontrado neste Windows.
        echo       Instale o Python manualmente em https://www.python.org/downloads/
        echo       marcando "Add to PATH" e "tcl/tk and IDLE", e execute este instalador de novo.
        goto :fim_erro
    )
    winget install -e --id Python.Python.3.12 --source winget --scope user --accept-package-agreements --accept-source-agreements
    call :achar_python
)
if not defined PYV (
    echo ERRO: nao foi possivel instalar o Python automaticamente.
    echo       Instale manualmente em https://www.python.org/downloads/ e execute este instalador de novo.
    goto :fim_erro
)
echo       Usando Python %PYV%.

rem ---- 2. Ambiente virtual + bibliotecas
echo [2/3] Instalando as bibliotecas Python na pasta .venv...
if not exist ".venv\Scripts\python.exe" (
    py -%PYV% -m venv .venv
    if errorlevel 1 (
        echo ERRO: nao foi possivel criar o ambiente virtual.
        goto :fim_erro
    )
)
".venv\Scripts\python.exe" -m pip install --upgrade pip --disable-pip-version-check
".venv\Scripts\python.exe" -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 (
    echo ERRO: falha ao instalar as bibliotecas. Verifique a internet ou o proxy e tente de novo.
    goto :fim_erro
)

rem ---- 3. Verificacao final
echo [3/3] Testando a instalacao...
".venv\Scripts\python.exe" -c "import cv2, PIL, openpyxl, tkinter; print('      Bibliotecas Python OK')"
if errorlevel 1 (
    echo ERRO: alguma biblioteca nao carregou corretamente.
    goto :fim_erro
)

echo.
echo ============================================================
echo   INSTALACAO CONCLUIDA
echo   Para abrir o programa, de dois cliques em iniciar.bat
echo   Na primeira vez, clique em Configurar IA e cole a chave de API.
echo ============================================================
pause
exit /b 0

:fim_erro
echo.
echo A instalacao NAO foi concluida. Leia as mensagens acima.
pause
exit /b 1

rem ------------------------------------------------------------
rem  Sub-rotinas
rem ------------------------------------------------------------
:achar_python
set "PYV="
for %%V in (3.12 3.11 3.13 3.10) do (
    if not defined PYV (
        py -%%V -c "import tkinter" >nul 2>&1 && set "PYV=%%V"
    )
)
exit /b 0
