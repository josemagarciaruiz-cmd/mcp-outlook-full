@echo off
title Instalar MCP Outlook (conector completo + agente)
echo ============================================================
echo   INSTALADOR MCP OUTLOOK (conector completo + agente)
echo ============================================================
echo.

rem --- Comprobar que existe el lanzador de Python (py) o python en PATH ---
where py >nul 2>&1
if not errorlevel 1 goto CHECKFOLDER
where python >nul 2>&1
if not errorlevel 1 goto CHECKFOLDER
goto NOPYTHON

:CHECKFOLDER
rem --- Debe ejecutarse DENTRO de la carpeta (donde esta agent.py) ---
if exist agent.py goto RUN
goto NOFOLDER

:RUN
echo Lanzando el instalador. Se abrira el navegador para iniciar sesion en Microsoft 365.
echo.
powershell -ExecutionPolicy Bypass -File ".\instalar.ps1"
if errorlevel 1 goto FAILED
echo.
echo ============================================================
echo   LISTO. Abre Claude: 'outlook' y 'agente-outlook'.
echo ============================================================
pause
exit /b 0

:NOPYTHON
echo FALTA Python. Instalalo desde https://www.python.org/downloads/
echo IMPORTANTE: marca la casilla "Add python.exe to PATH".
echo Luego vuelve a hacer doble clic en este archivo.
echo.
pause
exit /b 1

:NOFOLDER
echo No encuentro agent.py.
echo Descomprime el ZIP y ejecuta este archivo DESDE DENTRO de la carpeta.
echo.
pause
exit /b 1

:FAILED
echo.
echo Hubo un problema durante la instalacion. Abre install_log.txt para ver el detalle.
echo.
pause
exit /b 1
