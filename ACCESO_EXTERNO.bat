@echo off
title Stock Doina — Acceso Externo (Cloudflare Tunnel)

echo.
echo =========================================
echo  Stock Doina — Tunel publico gratuito
echo  (Cloudflare Tunnel - sin registro)
echo =========================================
echo.
echo  IMPORTANTE: La app debe estar corriendo.
echo  Ejecuta INICIAR_APP.bat primero en otra ventana.
echo.

:: Verificar si cloudflared ya esta instalado
where cloudflared >nul 2>&1
if %errorlevel% equ 0 goto :iniciar

:: Buscar en carpeta local
if exist "%~dp0cloudflared.exe" (
    set PATH=%PATH%;%~dp0
    goto :iniciar
)

:: Descargar cloudflared automaticamente (sin necesidad de cuenta)
echo  Descargando cloudflared...
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile '%~dp0cloudflared.exe'"
if not exist "%~dp0cloudflared.exe" (
    echo.
    echo  Error al descargar. Descargalo manualmente desde:
    echo  https://github.com/cloudflare/cloudflared/releases/latest
    echo  Archivo: cloudflared-windows-amd64.exe
    echo  Copialo en esta carpeta y volvé a ejecutar este .bat
    pause
    exit
)
set PATH=%PATH%;%~dp0
echo  Descargado correctamente.
echo.

:iniciar
echo  Creando tunel publico...
echo  La URL aparecera en unos segundos (busca la linea con trycloudflare.com)
echo  Compartila con quien quiera ver el reporte.
echo.
echo  (Ctrl+C para cerrar el tunel)
echo =========================================
echo.

cloudflared tunnel --url http://localhost:8501

pause
