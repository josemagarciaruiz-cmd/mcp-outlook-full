#!/bin/bash
cd "$(dirname "$0")"
echo "Cerrando Claude si estuviera abierto..."
osascript -e 'quit app "Claude"' 2>/dev/null || true
sleep 2
echo "Instalando MCP Outlook (conector completo + agente)..."
chmod +x instalar.sh 2>/dev/null || true
./instalar.sh
echo ""
echo "Puedes cerrar esta ventana. Abre Claude y prueba: 'usa outlook para ver mis ultimos correos'."
read -n 1 -s -r -p "Pulsa una tecla para salir..."
