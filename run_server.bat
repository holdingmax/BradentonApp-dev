@echo off
title BradentonApp - servidor
cd /d "%~dp0"

:loop
echo [%date% %time%] Iniciando webapp.py...
python webapp.py
echo [%date% %time%] El servidor se cerro (codigo %errorlevel%). Reiniciando en 2 segundos...
timeout /t 2 /nobreak >nul
goto loop
