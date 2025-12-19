@echo off
echo Iniciando o Sistema de RMs...
cd /d "%~dp0"
python -m streamlit run app.py
pause