"""Configuración central del servidor MCP de Outlook.

Todo se toma de variables de entorno para poder desplegar la MISMA pieza
en local (stdio, una por Mac) y en el VPS (HTTP) sin tocar código.
"""

from __future__ import annotations

import os
from pathlib import Path

# --- Identidad de la app de Azure (registro único, cuenta M365 del despacho) ---
# OUTLOOK_CLIENT_ID  -> Application (client) ID del registro de Entra ID.
# OUTLOOK_TENANT_ID  -> GUID del tenant del despacho (o dominio). Para trabajo,
#                       usa el tenant EXACTO; "organizations" acepta cualquier
#                       cuenta de trabajo/escuela; "common" también personales.
CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "").strip()
TENANT_ID = os.environ.get("OUTLOOK_TENANT_ID", "organizations").strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Permisos DELEGADOS (el server actúa en tu nombre). MSAL añade solo
# openid/profile/offline_access. offline_access es lo que da el refresh token.
DEFAULT_SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "MailboxSettings.ReadWrite",
    "Tasks.ReadWrite",              # Microsoft To-Do (plazos/tareas)
    "Files.ReadWrite",              # OneDrive (ficheros, enlaces)
    "Contacts.ReadWrite",           # contactos
    "Mail.ReadWrite.Shared",        # buzones compartidos (leer/gestionar)
    "Mail.Send.Shared",             # enviar desde/en nombre de buzon compartido
    "Calendars.ReadWrite.Shared",   # calendarios compartidos/de otros
    "People.Read",                  # personas frecuentes (buscar emails)
    "User.ReadBasic.All",           # directorio de la organizacion
]
SCOPES = [
    s.strip()
    for s in os.environ.get("OUTLOOK_SCOPES", " ".join(DEFAULT_SCOPES)).split()
    if s.strip()
]

# --- Persistencia del token (llavero/archivo). En el VPS: volumen persistente. ---
TOKEN_CACHE_PATH = Path(
    os.environ.get(
        "OUTLOOK_TOKEN_CACHE",
        str(Path.home() / ".outlook-mcp" / "token_cache.bin"),
    )
).expanduser()

# --- Microsoft Graph ---
GRAPH_BASE = os.environ.get("OUTLOOK_GRAPH_BASE", "https://graph.microsoft.com/v1.0")

# Zona horaria para lecturas de calendario y para interpretar fechas de eventos.
TIMEZONE = os.environ.get("OUTLOOK_TIMEZONE", "Europe/Madrid")

# --- Modo HTTP (VPS) ---
HTTP_HOST = os.environ.get("OUTLOOK_HTTP_HOST", "0.0.0.0")
HTTP_PORT = int(os.environ.get("OUTLOOK_HTTP_PORT", "8000"))
# Si se define, el endpoint HTTP exige Authorization: Bearer <token>.
# (Además de la ruta-token secreta que ponga Traefik por delante.)
HTTP_BEARER_TOKEN = os.environ.get("OUTLOOK_HTTP_BEARER", "").strip()

# Solo lectura: desactiva toda operación de escritura/envío/borrado.
READ_ONLY = os.environ.get("OUTLOOK_READ_ONLY", "").lower() in ("1", "true", "yes")


def require_client_id() -> str:
    if not CLIENT_ID:
        raise RuntimeError(
            "Falta OUTLOOK_CLIENT_ID. Registra una app en Entra ID "
            "(portal.azure.com -> App registrations) y exporta su Application "
            "(client) ID como OUTLOOK_CLIENT_ID. Ver README, sección 'Registro en Azure'."
        )
    return CLIENT_ID
