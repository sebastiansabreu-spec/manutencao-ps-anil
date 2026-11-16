@echo off
title Sistema de Manutencao - PS Anil
echo ========================================================
echo   Iniciando Painel Central - PS Anil
echo   Abrindo no navegador: http://127.0.0.1:8000
echo ========================================================
echo.
timeout /t 2 /nobreak >nul
start http://127.0.0.1:8000
.venv\Scripts\uvicorn.exe main:app --reload --host 127.0.0.1 --port 8000
pause
