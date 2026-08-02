@echo off
title Instalar MCP Outlook (conector completo + agente)
echo ============================================================
echo   INSTALADOR MCP OUTLOOK (conector 63 tools + agente)
echo ============================================================
echo.
where python >nul 2>&1
if errorlevel 1 (
  echo FALTA Python. Instalalo desde https://www.python.org/downloads/
  echo   IMPORTANTE: marca la casilla "Add python.exe to PATH".
  echo Luego vuelve a hacer doble clic en este archivo.
  echo.
  pause
  exit /b
)
if exist agent.py goto INSTALL
where git >nul 2>&1
if errorlevel 1 (
  echo No encuentro el proyecto ni Git.
  echo Descomprime el ZIP y ejecuta este archivo DESDE DENTRO de la carpeta.
  echo.
  pause
  exit /b
)
if not exist mcp-outlook-full git clone https://github.com/josemagarciaruiz-cmd/mcp-outlook-full
cd mcp-outlook-full
:INSTALL
echo Lanzando el instalador. Se abrira el navegador para iniciar sesion en Microsoft 365.
echo.
powershell -ExecutionPolicy Bypass -File .\instalar.ps1
echo.
echo ============================================================
echo   LISTO. Abre Claude: 'outlook' (63 tools) y 'agente-outlook'.
echo ============================================================
pause
