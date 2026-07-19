@echo off
setlocal
cd /d "%~dp0"
title Jirau V11 Profissional
where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Falha ao instalar as bibliotecas. Verifique sua internet e o Python.
  pause
  exit /b 1
)
%PY% -m streamlit run app.py
pause
