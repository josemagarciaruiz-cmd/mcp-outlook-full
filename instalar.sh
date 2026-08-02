#!/usr/bin/env bash
# Instalador COMPLETO de MCP Outlook (conector 63 tools + agente local).
# macOS / Linux. Ejecutar con Claude CERRADO:  ./instalar.sh
set -e
cd "$(dirname "$0")"

echo "==> 1/3  Entorno y dependencias (conector + agente en un solo venv)..."
python3 -m venv .venv
./.venv/bin/pip install -q --upgrade pip
./.venv/bin/pip install -q -e .

if [ ! -f .env ]; then
  cp .env.example .env
fi

# Cargar variables del .env para el login
set -a
# shellcheck disable=SC1091
. ./.env
set +a

if [ -z "${OUTLOOK_CLIENT_ID:-}" ] || [ -z "${OUTLOOK_TENANT_ID:-}" ]; then
  echo
  echo "!! Falta configurar el .env."
  echo "   Abre  $(pwd)/.env  y rellena OUTLOOK_CLIENT_ID y OUTLOOK_TENANT_ID."
  echo "   (Cómo obtenerlos: GUIA DETALLADA / README, sección 'Variables de Azure'.)"
  echo "   Luego repite ./instalar.sh"
  exit 1
fi

echo
echo "==> 2/3  Iniciando sesion en Microsoft 365 (una sola vez)."
echo "         Se abrira el navegador con un codigo: TECLEALO e inicia sesion con TU cuenta."
./.venv/bin/python -m outlook_mcp login

echo
echo "==> 3/3  Registrando los conectores 'outlook' + 'agente-outlook' en Claude..."
./.venv/bin/python conectar_claude.py

echo
echo "LISTO. Abre Claude. Tendras dos conectores:"
echo "  - 'outlook'        (63 herramientas: correo, agenda, tareas, OneDrive, contactos)"
echo "  - 'agente-outlook' (puente con tu disco: adjuntar/guardar ficheros, .eml, OneDrive)"
