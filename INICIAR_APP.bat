@echo off
title Stock Doina — Ventas

:: Matar cualquier instancia anterior de Streamlit en el puerto 8501
echo Cerrando instancias anteriores...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8501"') do taskkill /f /pid %%a 2>nul
timeout /t 2 /nobreak >nul

:: Ir a la carpeta del script
cd /d "%~dp0"

:: Buscar python en el entorno virtual del proyecto, o usar el del sistema
if exist "%~dp0..\proyecto reportes web doina\venv\Scripts\python.exe" (
    set PYTHON="%~dp0..\proyecto reportes web doina\venv\Scripts\python.exe"
) else (
    set PYTHON=python
)

echo.
echo =========================================
echo  Stock Doina — Ventas
echo  Iniciando servidor en puerto 8501...
echo =========================================
echo.
echo  Acceso local:    http://localhost:8501
echo  Acceso en red:   http://%COMPUTERNAME%:8501
echo.
echo  Para acceso externo ver ngrok mas abajo
echo  (Ctrl+C para detener)
echo =========================================
echo.

%PYTHON% -m streamlit run informe_stock.py ^
    --server.port 8501 ^
    --server.address 0.0.0.0 ^
    --server.headless true ^
    --server.enableCORS false ^
    --server.enableXsrfProtection false ^
    --browser.gatherUsageStats false

pause
