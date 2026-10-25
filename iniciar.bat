@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Ambiente nao encontrado. Execute primeiro o instalar.bat.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" ocr_midias.py
if errorlevel 1 pause
