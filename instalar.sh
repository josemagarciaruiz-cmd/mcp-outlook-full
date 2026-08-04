#!/usr/bin/env bash
# Instalador COMPLETO de MCP Outlook (conector + agente) para macOS / Linux.
# Ejecutar con Claude CERRADO:  ./instalar.sh
# Control de errores explícito (set -e) + mensajes claros.
set -euo pipefail
cd "$(dirname "$0")"

echo "==> 1/4  Buscando Python 3.13 / 3.12 / 3.11..."
PYBIN=""
for v in 3.13 3.12 3.11; do
  if command -v "python$v" >/dev/null 2>&1; then PYBIN="python$v"; break; fi
done
if [ -z "$PYBIN" ]; then
  # Aceptar python3 si está en el rango 3.11-3.13
  if command -v python3 >/dev/null 2>&1; then
    ok=$(python3 -c "import sys;print(1 if (3,11)<=sys.version_info[:2]<=(3,13) else 0)")
    [ "$ok" = "1" ] && PYBIN="python3"
  fi
fi
if [ -z "$PYBIN" ]; then
  echo "!! No hay Python 3.11, 3.12 o 3.13. Instala 3.13 desde python.org y repite."
  exit 1
fi
echo "     Usando $PYBIN"

"$PYBIN" -m venv .venv
./.venv/bin/python -m pip install -q --upgrade pip
echo "==> 2/4  Instalando dependencias..."
if [ -d wheels ]; then
  ./.venv/bin/python -m pip install -q --no-index --find-links wheels -r requirements.txt
else
  ./.venv/bin/python -m pip install -q -r requirements.txt
fi

export PYTHONPATH="$(pwd)"
echo "==> 3/4  Comprobando que todo carga (conector + agente)..."
./.venv/bin/python -c "import mcp, msal, httpx; from mcp.server.mcpserver import MCPServer; from outlook_mcp.server import build_server; import agent; print('OK imports')"

[ -f .env ] || cp .env.example .env
set -a; . ./.env; set +a
if [ -z "${OUTLOOK_CLIENT_ID:-}" ] || [ -z "${OUTLOOK_TENANT_ID:-}" ]; then
  echo "!! Falta configurar el .env: rellena OUTLOOK_CLIENT_ID y OUTLOOK_TENANT_ID (ver GUIA DETALLADA)."
  exit 1
fi

echo "==> 4/4  Iniciando sesion en Microsoft 365 y registrando en Claude..."
echo "     Se abrira el navegador con un codigo: TECLEALO e inicia sesion con TU cuenta."
./.venv/bin/python -m outlook_mcp login
./.venv/bin/python conectar_claude.py

echo
echo "LISTO. Abre Claude: 'outlook' (conector completo) y 'agente-outlook' (puente con disco)."
