@echo off
title Sistema de Manutencao - PS Anil
echo ========================================================
echo   Iniciando Painel Central - PS Anil
echo   Acesso Local: http://127.0.0.1:8000
echo   Acesso em Rede: http://0.0.0.0:8000
echo ========================================================
echo.
timeout /t 2 /nobreak >nul
start http://127.0.0.1:8000
.venv\Scripts\uvicorn.exe main:app --reload --host 0.0.0.0 --port 8000
pause
