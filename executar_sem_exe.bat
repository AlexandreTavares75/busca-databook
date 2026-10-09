@echo off
chcp 65001 >nul
rem Roda direto do codigo-fonte (para testes). Usuarios finais usam o instalador.
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
start "" pythonw src\app.py