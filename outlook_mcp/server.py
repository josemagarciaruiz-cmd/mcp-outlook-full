"""Servidor MCP de Outlook: correo + calendario sobre Microsoft Graph."""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from . import (
    auth,
    config,
    tools_calendar,
    tools_contacts,
    tools_directory,
    tools_files,
    tools_mail,
    tools_mail_extra,
    tools_tasks,
)

INSTRUCTIONS = """\
Servidor MCP de Outlook/Microsoft 365 (correo, calendario, tareas, OneDrive, contactos).

Correo: mail_*  (listar, buscar, leer, redactar, enviar, responder/reenviar CON
adjuntos, marcar, mover, categorías, adjuntos, guardar/exportar, hilo, lote,
reglas de bandeja, respuesta automática/fuera de oficina). Carpetas ANIDADAS:
mail_folder_tree da todo el árbol (carpetas por cliente bajo la Bandeja de entrada);
mail_list_folders(parent='inbox') sus subcarpetas; luego mail_list_messages(folder=<id>).
Calendario: calendar_*  (agenda, crear/editar eventos y reuniones de Teams,
recurrencias, invitaciones, cancelar, huecos y disponibilidad).
Tareas (Microsoft To-Do): tasks_*  (listas, crear/editar/completar tareas con plazo).
OneDrive: files_*  (listar, buscar, descargar, subir, crear enlace de compartir).
Contactos: contacts_*  (listar, buscar, crear, editar; carpetas de contactos).
Directorio: people_search (personas frecuentes) y directory_search (organización).
Buzón compartido: casi todas las tools de correo aceptan 'mailbox' (email del buzón).
Calendarios de otros: calendar_list_events(user=email); crear en calendario concreto: calendar_create_event(calendar_id=...).

Convenciones:
- Correo: por defecto se opera sobre la bandeja PRINCIPAL (Focused). mail_list_messages
  devuelve solo Principal salvo classification='other' o 'all'. Incluir "Otros" solo si
  el usuario lo pide.
- Adjuntar/guardar/descargar ficheros LOCALES (attachment_paths, *_download, files_upload,
  mail_save_attachments) requiere el conector LOCAL; por URL/OneDrive (attachment_urls) sirve
  también el VPS.
- Las fechas son ISO 8601. La zona por defecto es la de OUTLOOK_TIMEZONE.
- El borrado de correo va a Papelera salvo permanent=True.
- Enviar correo y responder invitaciones son acciones que salen al exterior:
  confirma con el usuario antes de ejecutarlas.
"""


def build_server() -> MCPServer:
    # En mcp 2.0 host/port van en run()/streamable_http_app(), no en el constructor.
    mcp = MCPServer("outlook", instructions=INSTRUCTIONS, version="0.8.0")

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def whoami() -> dict:
        """Devuelve la cuenta autenticada y verifica el acceso a Graph.

        Útil como primera comprobación: si falla, hay que hacer login
        (python -m outlook_mcp login).
        """
        from . import graph

        me = await graph.request("GET", "/me", params={"$select": "displayName,userPrincipalName,mail,id"})
        return {
            "displayName": me.get("displayName"),
            "email": me.get("mail") or me.get("userPrincipalName"),
            "id": me.get("id"),
            "cached_account": (auth.current_account() or {}).get("account"),
        }

    tools_mail.register(mcp)
    tools_mail_extra.register(mcp)
    tools_calendar.register(mcp)
    tools_tasks.register(mcp)
    tools_files.register(mcp)
    tools_contacts.register(mcp)
    tools_directory.register(mcp)
    return mcp
