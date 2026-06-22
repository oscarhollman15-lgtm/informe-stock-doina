@echo off
setlocal enabledelayedexpansion
title Setup Completo — Stock Doina Ventas
color 0A

set PYTHON=C:\Users\Operador\AppData\Local\Programs\Python\Python312\python.exe
set STREAMLIT=C:\Users\Operador\AppData\Local\Programs\Python\Python312\Scripts\streamlit.exe
set APP_DIR=%~dp0
set APP_SCRIPT=%~dp0informe_stock.py
set NGROK_CFG=%USERPROFILE%\AppData\Local\ngrok\ngrok.yml

echo.
echo  =====================================================
echo   SETUP COMPLETO — Stock Doina Ventas
echo   Configura arranque automatico + URL publica fija
echo  =====================================================
echo.

:: ─── PASO 1: Instalar ngrok ───────────────────────────────────────────────────
echo  [1/5] Verificando ngrok...
where ngrok >nul 2>&1
if %errorlevel% equ 0 (
    echo       OK - ngrok ya instalado
) else (
    echo       Instalando ngrok...
    winget install ngrok.ngrok --silent
    if %errorlevel% neq 0 (
        echo.
        echo       winget no disponible. Descargando manualmente...
        powershell -Command "Invoke-WebRequest -Uri 'https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-v3-stable-windows-amd64.zip' -OutFile '%TEMP%\ngrok.zip'; Expand-Archive '%TEMP%\ngrok.zip' -DestinationPath 'C:\Windows\System32' -Force"
    )
    where ngrok >nul 2>&1
    if %errorlevel% neq 0 (
        echo  ERROR: No se pudo instalar ngrok.
        echo  Descargalo de https://ngrok.com/download y copialo a C:\Windows\System32\
        pause & exit
    )
    echo       OK - ngrok instalado
)

:: ─── PASO 2: Registrar cuenta ngrok ──────────────────────────────────────────
echo.
echo  [2/5] Configurar cuenta ngrok (GRATIS)
echo.
echo  Si todavia no tenes cuenta:
echo   1. Abre el navegador en: https://dashboard.ngrok.com/signup
echo   2. Registrate (email + contrasena, no necesitas tarjeta)
echo   3. Volvé acá
echo.
start https://dashboard.ngrok.com/signup
echo  Presiona una tecla cuando tengas la cuenta lista...
pause >nul

:: ─── PASO 3: Auth token ───────────────────────────────────────────────────────
echo.
echo  [3/5] Auth token
echo.
echo  Andá a: https://dashboard.ngrok.com/authtokens
echo  Copia el token que aparece (empieza con "2..." o similar)
echo.
start https://dashboard.ngrok.com/authtokens
echo.
set /p NGROK_TOKEN="  Pega el token acá y presiona Enter: "
if "%NGROK_TOKEN%"=="" (
    echo  ERROR: Token vacío.
    pause & exit
)
ngrok config add-authtoken %NGROK_TOKEN%
echo  OK - Token configurado

:: ─── PASO 4: Dominio estático gratuito ───────────────────────────────────────
echo.
echo  [4/5] Dominio estatico gratuito
echo.
echo  Andá a: https://dashboard.ngrok.com/domains
echo  Hace clic en "New Domain" y ngrok te genera uno gratis.
echo  Ejemplo: doina-ventas-abc123.ngrok-free.app
echo  (ese dominio sera SIEMPRE el mismo, no cambia nunca)
echo.
start https://dashboard.ngrok.com/domains
echo.
set /p NGROK_DOMAIN="  Pega el dominio que te dieron (solo el nombre, sin https://): "
if "%NGROK_DOMAIN%"=="" (
    echo  ERROR: Dominio vacío.
    pause & exit
)

:: Guardar dominio en archivo para el script de arranque
echo %NGROK_DOMAIN% > "%APP_DIR%ngrok_domain.txt"
echo  OK - Dominio guardado: %NGROK_DOMAIN%

:: ─── PASO 5: Crear tareas en el Programador de Windows ───────────────────────
echo.
echo  [5/5] Configurando arranque automatico con Windows...

:: Crear script VBS para iniciar Streamlit sin ventana visible
echo Set WShell = CreateObject("WScript.Shell") > "%APP_DIR%_start_app.vbs"
echo WShell.Run """" ^& "%STREAMLIT%" ^& """" ^& " run """ ^& "%APP_SCRIPT%" ^& """ --server.port 8501 --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false", 0, False >> "%APP_DIR%_start_app.vbs"

:: Crear script VBS para iniciar ngrok sin ventana visible
echo Set WShell = CreateObject("WScript.Shell") > "%APP_DIR%_start_ngrok.vbs"
echo WShell.Run "ngrok http --domain=%NGROK_DOMAIN% 8501", 0, False >> "%APP_DIR%_start_ngrok.vbs"

:: Registrar tarea: Streamlit
schtasks /delete /tn "StockDoina_App" /f >nul 2>&1
schtasks /create /tn "StockDoina_App" ^
    /tr "wscript.exe \"%APP_DIR%_start_app.vbs\"" ^
    /sc ONLOGON /delay 0001:00 ^
    /ru "%USERNAME%" /f >nul
if %errorlevel% equ 0 (
    echo  OK - Tarea StockDoina_App creada
) else (
    echo  AVISO: No se pudo crear la tarea del app ^(puede requerir admin^)
)

:: Registrar tarea: ngrok
schtasks /delete /tn "StockDoina_Ngrok" /f >nul 2>&1
schtasks /create /tn "StockDoina_Ngrok" ^
    /tr "wscript.exe \"%APP_DIR%_start_ngrok.vbs\"" ^
    /sc ONLOGON /delay 0001:30 ^
    /ru "%USERNAME%" /f >nul
if %errorlevel% equ 0 (
    echo  OK - Tarea StockDoina_Ngrok creada
) else (
    echo  AVISO: No se pudo crear la tarea del tunel ^(puede requerir admin^)
)

:: ─── Iniciar ahora mismo ──────────────────────────────────────────────────────
echo.
echo  Iniciando servicios ahora...
start "" wscript.exe "%APP_DIR%_start_app.vbs"
timeout /t 4 /nobreak >nul
start "" wscript.exe "%APP_DIR%_start_ngrok.vbs"
timeout /t 3 /nobreak >nul

echo.
echo  =====================================================
echo   LISTO. Todo configurado.
echo.
echo   URL publica permanente:
echo   https://%NGROK_DOMAIN%
echo.
echo   Compartila con quien quiera ver el reporte.
echo   No cambia nunca. Arranca sola con Windows.
echo  =====================================================
echo.
start https://%NGROK_DOMAIN%
pause
