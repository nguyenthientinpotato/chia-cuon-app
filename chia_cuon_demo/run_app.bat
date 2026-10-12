@echo off
chcp 65001 >nul
cd /d "%~dp0"
python --version >nul 2>&1 || (echo Chua cai Python. Tai tai https://www.python.org/downloads/ va tick "Add Python to PATH". & pause & exit /b)
python -c "import streamlit, ortools, numpy, openpyxl" >nul 2>&1 || (
  echo Lan dau chay: dang cai thu vien, vui long doi 1-3 phut...
  python -m pip install -r requirements.txt
)
python -m streamlit run app.py
pause
