"""Herramientas de correo avanzadas: guardar adjuntos, exportar, hilo, acciones
en lote, categorías, reglas de bandeja y respuestas automáticas (fuera de oficina).
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Annotated, Optional

from pydantic import Field

from . import format as fmt
from . import graph, helpers


def register(mcp) -> None:
    # ------------------------------------------------------- guardar / exportar
    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_save_attachments(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        dest_dir: Annotated[str, Field(description="Carpeta LOCAL donde guardar (conector local; el VPS no ve tu disco)")],
        attachment_id: Annotated[Optional[str], Field(None, description="Guardar solo este adjunto; por defecto todos")] = None,
    ) -> dict:
        """Descarga y GUARDA en disco los adjuntos de un correo (conector local)."""
        helpers.guard_write()
        d = Path(dest_dir).expanduser()
        d.mkdir(parents=True, exist_ok=True)
        if attachment_id:
            atts = [await graph.request("GET", f"/me/messages/{message_id}/attachments/{attachment_id}")]
        else:
            res = await graph.request("GET", f"/me/messages/{message_id}/attachments")
            atts = res.get("value", []) if isinstance(res, dict) else []
        saved = []
        for a in atts:
            cb = a.get("contentBytes")
            if not cb:
                continue
            p = d / (a.get("name") or "adjunto")
            p.write_bytes(base64.b64decode(cb))
            saved.append({"name": p.name, "path": str(p), "size": a.get("size")})
        return {"saved": saved, "dest_dir": str(d)}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_export_eml(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        dest_dir: Annotated[str, Field(description="Carpeta LOCAL donde guardar el .eml (conector local)")],
        filename: Annotated[Optional[str], Field(None, description="Nombre del fichero; por defecto el asunto")] = None,
    ) -> dict:
        """Exporta un correo completo como fichero .eml (MIME) al disco (conector local)."""
        helpers.guard_write()
        raw = await graph.request("GET", f"/me/messages/{message_id}/$value")
        if not isinstance(raw, (bytes, bytearray)):
            raw = str(raw).encode("utf-8")
        d = Path(dest_dir).expanduser()
        d.mkdir(parents=True, exist_ok=True)
        name = filename or "correo"
        if not name.lower().endswith(".eml"):
            name += ".eml"
        name = name.replace("/", "-").replace(":", "-")
        p = d / name
        p.write_bytes(raw)
        return {"path": str(p), "bytes": len(raw)}

    # --------------------------------------------------------------- hilo
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_get_conversation(
        message_id: Annotated[str, Field(description="Id de un mensaje del hilo")],
        top: Annotated[int, Field(50, ge=1, le=100)] = 50,
    ) -> dict:
        """Trae TODOS los mensajes de la conversación (hilo) de un correo, ordenados."""
        m = await graph.request("GET", f"/me/messages/{message_id}", params={"$select": "conversationId"})
        cid = m.get("conversationId")
        if not cid:
            return {"messages": [], "note": "El mensaje no tiene conversationId."}
        # $filter conversationId no admite $orderby -> ordenamos en cliente
        res = await graph.request(
            "GET", "/me/messages",
            params={"$filter": f"conversationId eq '{cid}'", "$top": top,
                    "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,parentFolderId"},
        )
        items = res.get("value", []) if isinstance(res, dict) else []
        items.sort(key=lambda x: x.get("receivedDateTime", ""))
        return {"conversationId": cid, "count": len(items), "messages": [fmt.fmt_message(x) for x in items]}

    # ------------------------------------------------------------- lote
    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_bulk(
        message_ids: Annotated[list[str], Field(description="Ids de los mensajes")],
        action: Annotated[str, Field(description="mark_read | mark_unread | move | trash | delete_permanent | flag | unflag")],
        destination_folder: Annotated[Optional[str], Field(None, description="Carpeta destino (solo para action=move)")] = None,
    ) -> dict:
        """Aplica una acción a VARIOS correos de una vez (leer, mover, borrar, marcar)."""
        helpers.guard_write()
        ok, errors = 0, []
        for mid in message_ids:
            try:
                if action == "mark_read":
                    await graph.request("PATCH", f"/me/messages/{mid}", json={"isRead": True})
                elif action == "mark_unread":
                    await graph.request("PATCH", f"/me/messages/{mid}", json={"isRead": False})
                elif action == "flag":
                    await graph.request("PATCH", f"/me/messages/{mid}", json={"flag": {"flagStatus": "flagged"}})
                elif action == "unflag":
                    await graph.request("PATCH", f"/me/messages/{mid}", json={"flag": {"flagStatus": "notFlagged"}})
                elif action == "move":
                    if not destination_folder:
                        raise ValueError("Falta destination_folder para move.")
                    await graph.request("POST", f"/me/messages/{mid}/move", json={"destinationId": destination_folder})
                elif action == "trash":
                    await graph.request("POST", f"/me/messages/{mid}/move", json={"destinationId": "deleteditems"})
                elif action == "delete_permanent":
                    await graph.request("DELETE", f"/me/messages/{mid}")
                else:
                    raise ValueError(f"Acción no soportada: {action}")
                ok += 1
            except Exception as e:  # noqa: BLE001
                errors.append({"id": mid[:16] + "…", "error": str(e)[:120]})
        return {"action": action, "ok": ok, "failed": len(errors), "errors": errors}

    # -------------------------------------------------------- categorías
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_list_categories() -> dict:
        """Lista las categorías maestras de Outlook (nombre y color)."""
        res = await graph.request("GET", "/me/outlook/masterCategories")
        return {"categories": [{"id": c.get("id"), "name": c.get("displayName"), "color": c.get("color")}
                               for c in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_create_category(
        name: Annotated[str, Field(description="Nombre de la categoría")],
        color: Annotated[str, Field("preset0", description="Color preset0..preset24 de Outlook")] = "preset0",
    ) -> dict:
        """Crea una categoría maestra (luego se aplica con mail_update categories=[...])."""
        helpers.guard_write()
        res = await graph.request("POST", "/me/outlook/masterCategories", json={"displayName": name, "color": color})
        return {"id": res.get("id"), "name": res.get("displayName"), "color": res.get("color")}

    # ------------------------------------------------------------ reglas
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_list_rules() -> dict:
        """Lista las reglas de la Bandeja de entrada (auto-archivado, etc.)."""
        res = await graph.request("GET", "/me/mailFolders/inbox/messageRules")
        out = []
        for r in (res.get("value", []) if isinstance(res, dict) else []):
            out.append({"id": r.get("id"), "name": r.get("displayName"),
                        "enabled": r.get("isEnabled"), "sequence": r.get("sequence"),
                        "conditions": r.get("conditions"), "actions": r.get("actions")})
        return {"rules": out}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_create_rule(
        name: Annotated[str, Field(description="Nombre de la regla")],
        move_to_folder: Annotated[Optional[str], Field(None, description="Id/carpeta destino a la que mover")] = None,
        from_addresses: Annotated[Optional[list[str]], Field(None, description="Aplicar si el remitente es uno de estos emails")] = None,
        sender_contains: Annotated[Optional[list[str]], Field(None, description="Aplicar si el remitente contiene estos textos")] = None,
        subject_contains: Annotated[Optional[list[str]], Field(None, description="Aplicar si el asunto contiene estos textos")] = None,
        mark_as_read: Annotated[bool, Field(False)] = False,
        sequence: Annotated[int, Field(1, ge=1)] = 1,
    ) -> dict:
        """Crea una regla de bandeja: p.ej. correos de un cliente → su carpeta.

        Debe llevar al menos una condición y una acción (mover y/o marcar leído).
        """
        helpers.guard_write()
        conditions: dict = {}
        if from_addresses:
            conditions["fromAddresses"] = [{"emailAddress": {"address": a}} for a in from_addresses]
        if sender_contains:
            conditions["senderContains"] = sender_contains
        if subject_contains:
            conditions["subjectContains"] = subject_contains
        if not conditions:
            raise ValueError("Indica al menos una condición (from_addresses, sender_contains o subject_contains).")
        actions: dict = {}
        if move_to_folder:
            actions["moveToFolder"] = move_to_folder
        if mark_as_read:
            actions["markAsRead"] = True
        if not actions:
            raise ValueError("Indica al menos una acción (move_to_folder y/o mark_as_read).")
        rule = {"displayName": name, "sequence": sequence, "isEnabled": True,
                "conditions": conditions, "actions": actions}
        res = await graph.request("POST", "/me/mailFolders/inbox/messageRules", json=rule)
        return {"id": res.get("id"), "name": res.get("displayName"), "enabled": res.get("isEnabled")}

    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def mail_delete_rule(
        rule_id: Annotated[str, Field(description="Id de la regla a eliminar")],
    ) -> dict:
        """Elimina una regla de la Bandeja de entrada."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/mailFolders/inbox/messageRules/{rule_id}")
        return {"status": "deleted", "rule_id": rule_id}

    # --------------------------------------------- respuestas automáticas (OOO)
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_get_automatic_replies() -> dict:
        """Consulta el estado de la respuesta automática (fuera de oficina)."""
        res = await graph.request("GET", "/me/mailboxSettings", params={"$select": "automaticRepliesSetting"})
        return res.get("automaticRepliesSetting", {}) if isinstance(res, dict) else {}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_set_automatic_replies(
        status: Annotated[str, Field(description="disabled | alwaysEnabled | scheduled")],
        message: Annotated[str, Field("", description="Mensaje de respuesta (interno; y externo si no se da otro)")] = "",
        external_message: Annotated[Optional[str], Field(None, description="Mensaje para externos; por defecto = message")] = None,
        start: Annotated[Optional[str], Field(None, description="Inicio ISO 8601 (solo status=scheduled)")] = None,
        end: Annotated[Optional[str], Field(None, description="Fin ISO 8601 (solo status=scheduled)")] = None,
        external_audience: Annotated[str, Field("all", description="none | contactsOnly | all")] = "all",
    ) -> dict:
        """Activa/desactiva/programa la respuesta automática de fuera de oficina."""
        helpers.guard_write()
        setting: dict = {
            "status": status,
            "externalAudience": external_audience,
            "internalReplyMessage": message,
            "externalReplyMessage": external_message if external_message is not None else message,
        }
        if status == "scheduled" and start and end:
            setting["scheduledStartDateTime"] = helpers.date_time_zone(start)
            setting["scheduledEndDateTime"] = helpers.date_time_zone(end)
        await graph.request("PATCH", "/me/mailboxSettings", json={"automaticRepliesSetting": setting})
        return {"status": status}
