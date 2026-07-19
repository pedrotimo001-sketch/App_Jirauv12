@echo off
chcp 65001 >nul
title Jirau - Instalação de Bibliotecas
cd /d "%~dp0"
echo ===============================================
echo   JIRAU - INSTALANDO BIBLIOTECAS NECESSARIAS
echo ===============================================
echo.
where py >nul 2>nul
if %errorlevel%==0 (
    py -m pip install --upgrade pip
    py -m pip install -r requirements.txt
) else (
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
)
if errorlevel 1 (
    echo.
    echo ERRO: Nao foi possivel instalar as bibliotecas.
    echo Verifique sua internet e se o Python esta instalado.
    pause
    exit /b 1
)
echo.
echo Bibliotecas instaladas com sucesso.
pause
