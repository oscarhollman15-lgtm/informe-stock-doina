@echo off
title Reiniciar Stock Doina

echo Cerrando servicios anteriores...
taskkill /f /im streamlit.exe >nul 2>&1
taskkill /f /im ngrok.exe >nul 2>&1
timeout /t 2 /nobreak >nul

set APP_DIR=%~dp0
set STREAMLIT=C:\Users\Operador\AppData\Local\Programs\Python\Python312\Scripts\streamlit.exe

:: Leer dominio guardado
set /p NGROK_DOMAIN=<"%APP_DIR%ngrok_domain.txt"
set NGROK_DOMAIN=%NGROK_DOMAIN: =%

echo Iniciando app...
start "" wscript.exe "%APP_DIR%_start_app.vbs"
timeout /t 4 /nobreak >nul

echo Iniciando tunel...
start "" wscript.exe "%APP_DIR%_start_ngrok.vbs"
timeout /t 3 /nobreak >nul

echo.
echo Listo. App corriendo en:
echo https://%NGROK_DOMAIN%
echo.
pause
