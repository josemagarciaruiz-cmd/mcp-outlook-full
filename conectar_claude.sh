#!/usr/bin/env bash
# Da de alta los conectores 'outlook' + 'agente-outlook' en Claude Desktop (stdio).
# Ejecutar tras ./instalar.sh y con el .env relleno + login hecho.
set -e
cd "$(dirname "$0")"
if [ -x ./.venv/bin/python ]; then
  ./.venv/bin/python conectar_claude.py "$@"
else
  python3 conectar_claude.py "$@"
fi
