"""Herramientas de CORREO de Outlook (Microsoft Graph).

Cobertura completa: carpetas, listar/buscar/leer, redactar/enviar, responder,
reenviar, marcar, mover, categorizar, adjuntos y borrado a Papelera.
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import format as fmt
from . import graph, helpers

# Campos que pedimos por defecto al listar (respuestas compactas).
_LIST_SELECT = (
    "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,"
    "isRead,hasAttachments,importance,flag,categories,bodyPreview,webLink,conversationId"
)
_FULL_SELECT = _LIST_SELECT + ",bccRecipients,replyTo,parentFolderId,body"


def _iso(d: str) -> str:
    """Normaliza una fecha ISO a datetime de Graph con Z (medianoche si falta hora)."""
    d = d.strip()
    if "T" not in d:
        d = d + "T00:00:00"
    if not (d.endswith("Z") or "+" in d[10:] or "-" in d[11:]):
        d = d + "Z"
    return d


def register(mcp) -> None:
    # ---------------------------------------------------------------- carpetas
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_list_folders(
        parent: Annotated[Optional[str], Field(None, description="Id/nombre de carpeta cuyas SUBCARPETAS listar (p.ej. 'inbox' para ver las carpetas de clientes bajo la Bandeja de entrada). Por defecto, las carpetas de primer nivel.")] = None,
        top: Annotated[int, Field(200, ge=1, le=500)] = 200,
        mailbox: Annotated[Optional[str], Field(None, description="Email de un buzón COMPARTIDO al que tengas acceso; por defecto el tuyo")] = None,
    ) -> dict:
        """Lista carpetas de correo con contadores. Con 'parent' devuelve las
        SUBCARPETAS de esa carpeta (así se ven las carpetas anidadas por cliente).

        Devuelve id, nombre, no leídos, total y nº de subcarpetas. Usa los ids
        (o nombres conocidos: inbox, sentitems, drafts, deleteditems, archive)
        en las demás tools. Para ver TODA la jerarquía de golpe: mail_folder_tree.
        Con 'mailbox' opera sobre un buzón compartido.
        """
        root = helpers.mbox(mailbox)
        base = f"{root}/mailFolders/{parent}/childFolders" if parent else f"{root}/mailFolders"
        params = {"$top": min(top, 100), "$select": "id,displayName,parentFolderId,unreadItemCount,totalItemCount,childFolderCount"}
        res = await graph.get_paged(base, params=params, max_items=top)
        return {"folders": [fmt.fmt_folder(f) for f in res["items"]], "count": len(res["items"]), "next": res["next"]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_folder_tree(
        parent: Annotated[Optional[str], Field(None, description="Carpeta raíz desde la que mapear (p.ej. 'inbox'); por defecto todo el buzón")] = None,
        max_depth: Annotated[int, Field(4, ge=1, le=8)] = 4,
        max_folders: Annotated[int, Field(400, ge=1, le=2000)] = 400,
        mailbox: Annotated[Optional[str], Field(None, description="Email de un buzón COMPARTIDO; por defecto el tuyo")] = None,
    ) -> dict:
        """Devuelve el ÁRBOL COMPLETO de carpetas (anidadas incluidas) como lista
        plana con la ruta de cada una: p.ej. 'Bandeja de entrada/CLIENTE X'.

        Ideal para ver de una vez todas tus carpetas por cliente/colaborador con
        sus ids, para luego leer sus correos con mail_list_messages(folder=<id>).
        """
        out: list[dict] = []
        root = helpers.mbox(mailbox)

        async def walk(folder_id: Optional[str], prefix: str, depth: int) -> None:
            if len(out) >= max_folders:
                return
            path = f"{root}/mailFolders/{folder_id}/childFolders" if folder_id else f"{root}/mailFolders"
            res = await graph.get_paged(
                path,
                params={"$top": 100, "$select": "id,displayName,parentFolderId,unreadItemCount,totalItemCount,childFolderCount"},
                max_items=max_folders,
            )
            for f in res["items"]:
                if len(out) >= max_folders:
                    return
                name = f.get("displayName") or ""
                full = f"{prefix}/{name}" if prefix else name
                out.append({
                    "id": f.get("id"),
                    "path": full,
                    "name": name,
                    "total": f.get("totalItemCount"),
                    "unread": f.get("unreadItemCount"),
                    "children": f.get("childFolderCount", 0),
                })
                if f.get("childFolderCount", 0) and depth < max_depth:
                    await walk(f["id"], full, depth + 1)

        await walk(parent, "", 1)
        return {"folders": out, "count": len(out), "truncated": len(out) >= max_folders}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_create_folder(
        name: Annotated[str, Field(description="Nombre de la nueva carpeta")],
        parent_folder_id: Annotated[Optional[str], Field(None, description="Id de la carpeta padre; por defecto raíz")] = None,
    ) -> dict:
        """Crea una carpeta de correo (opcionalmente dentro de otra)."""
        helpers.guard_write()
        path = f"/me/mailFolders/{parent_folder_id}/childFolders" if parent_folder_id else "/me/mailFolders"
        res = await graph.request("POST", path, json={"displayName": name})
        return fmt.fmt_folder(res)

    # ------------------------------------------------------------ listar/buscar
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_list_messages(
        folder: Annotated[str, Field("inbox", description="Id de carpeta o nombre conocido: inbox, sentitems, drafts, deleteditems, archive, junkemail")] = "inbox",
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
        unread_only: Annotated[bool, Field(False)] = False,
        classification: Annotated[str, Field("focused", description="Solo en 'inbox': 'focused' = bandeja PRINCIPAL (por defecto), 'other' = bandeja Otros, 'all' = toda la bandeja mezclada.")] = "focused",
        since: Annotated[Optional[str], Field(None, description="Solo correos recibidos DESDE esta fecha (ISO, p.ej. 2026-01-01)")] = None,
        until: Annotated[Optional[str], Field(None, description="Solo correos recibidos HASTA esta fecha (ISO)")] = None,
        order_by: Annotated[str, Field("receivedDateTime desc", description="Campo OData de orden")] = "receivedDateTime desc",
        filter: Annotated[Optional[str], Field(None, description="Filtro OData avanzado, p.ej. \"from/emailAddress/address eq 'x@y.com'\"")] = None,
        mailbox: Annotated[Optional[str], Field(None, description="Email de un buzón COMPARTIDO al que tengas acceso; por defecto el tuyo")] = None,
    ) -> dict:
        """Lista los correos de una carpeta. Por defecto, la bandeja PRINCIPAL (Focused).

        En la Bandeja de entrada, Outlook separa 'Principal' (focused) de 'Otros'
        (other). Por defecto se devuelve SOLO Principal; usa classification='other'
        para Otros o 'all' para todo mezclado. Filtra por rango de fechas con
        since/until y opera sobre un buzón compartido con 'mailbox'. Respuesta
        compacta; cuerpo completo con mail_get_message; pagina con 'next'.
        """
        root = helpers.mbox(mailbox)
        filters = []
        if unread_only:
            filters.append("isRead eq false")
        if since:
            filters.append(f"receivedDateTime ge {_iso(since)}")
        if until:
            filters.append(f"receivedDateTime le {_iso(until)}")
        if filter:
            filters.append(f"({filter})")
        base_filter = " and ".join(filters) if filters else None
        cls = (classification or "focused").lower()

        # Carpeta distinta de inbox, o 'all': listado simple de una página.
        if folder != "inbox" or cls == "all":
            params = {"$top": top, "$orderby": order_by, "$select": _LIST_SELECT}
            if base_filter:
                params["$filter"] = base_filter
            res = await graph.get_paged(f"{root}/mailFolders/{folder}/messages", params=params, max_items=top)
            return {"messages": [fmt.fmt_message(m) for m in res["items"]], "next": res["next"]}

        # 'focused'/'other': Graph no admite $filter inferenceClassification junto a
        # $orderby (error InefficientFilter). Traemos ordenado por fecha con el campo
        # de clasificación y filtramos en cliente, paginando hasta reunir 'top'.
        select = _LIST_SELECT + ",inferenceClassification"
        params = {"$top": 50, "$orderby": order_by, "$select": select}
        if base_filter:
            params["$filter"] = base_filter
        collected: list = []
        data = await graph.request("GET", f"{root}/mailFolders/inbox/messages", params=params)
        page = data.get("value", [])
        next_link = data.get("@odata.nextLink")
        pages = 0
        while True:
            for m in page:
                if (m.get("inferenceClassification") or "focused") == cls:
                    collected.append(m)
            pages += 1
            if len(collected) >= top or not next_link or pages >= 6:
                break
            nxt = await graph.get_next(next_link, max_items=50)
            page, next_link = nxt["items"], nxt["next"]
        return {
            "messages": [fmt.fmt_message(m) for m in collected[:top]],
            "classification": cls,
            "next": None,
        }

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_search(
        query: Annotated[str, Field(description="Búsqueda de texto libre (KQL de Outlook), busca en asunto, cuerpo, remitente, etc.")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
        folder: Annotated[Optional[str], Field(None, description="Acotar la búsqueda a una carpeta (id o nombre conocido); por defecto todo el buzón")] = None,
        mailbox: Annotated[Optional[str], Field(None, description="Email de un buzón COMPARTIDO; por defecto el tuyo")] = None,
    ) -> dict:
        """Busca correos por texto libre (asunto, cuerpo, gente), en todo el buzón
        o acotado a una carpeta, y opcionalmente en un buzón compartido.

        Usa $search de Graph (relevancia). Para filtros por fecha o campo exacto usa
        mail_list_messages (since/until/filter). $search no combina con $orderby.
        """
        root = helpers.mbox(mailbox)
        base = f"{root}/mailFolders/{folder}/messages" if folder else f"{root}/messages"
        params = {"$search": f'"{query}"', "$top": top, "$select": _LIST_SELECT}
        res = await graph.request("GET", base, params=params)
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"messages": [fmt.fmt_message(m) for m in items], "next": res.get("@odata.nextLink")}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_get_message(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        body_as_text: Annotated[bool, Field(True, description="Devolver el cuerpo en texto plano (True) o HTML (False)")] = True,
        mailbox: Annotated[Optional[str], Field(None, description="Email de un buzón COMPARTIDO; por defecto el tuyo")] = None,
    ) -> dict:
        """Lee un correo completo: cuerpo, destinatarios, cc/bcc, adjuntos flag."""
        headers = graph.prefer_headers(text_body=body_as_text)
        res = await graph.request(
            "GET", f"{helpers.mbox(mailbox)}/messages/{message_id}", params={"$select": _FULL_SELECT}, headers=headers
        )
        return fmt.fmt_message(res, full=True)

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_next_page(
        next_link: Annotated[str, Field(description="URL @odata.nextLink devuelta en 'next'")],
        max_items: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Continúa una paginación de correo a partir del enlace 'next'."""
        res = await graph.get_next(next_link, max_items=max_items)
        return {"messages": [fmt.fmt_message(m) for m in res["items"]], "next": res["next"]}

    # --------------------------------------------------------------- adjuntos
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_list_attachments(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        include_content: Annotated[bool, Field(False, description="Incluir contenido en base64 (puede ser grande)")] = False,
    ) -> dict:
        """Lista los adjuntos de un correo (nombre, tipo, tamaño; contenido opcional)."""
        select = "id,name,contentType,size,isInline"
        if include_content:
            select += ",contentBytes"
        res = await graph.request("GET", f"/me/messages/{message_id}/attachments", params={"$select": select})
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"attachments": [fmt.fmt_attachment(a, with_content=include_content) for a in items]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def mail_get_attachment(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        attachment_id: Annotated[str, Field(description="Id del adjunto")],
    ) -> dict:
        """Descarga un adjunto concreto con su contenido en base64."""
        res = await graph.request("GET", f"/me/messages/{message_id}/attachments/{attachment_id}")
        return fmt.fmt_attachment(res, with_content=True)

    # ----------------------------------------------------------- redactar/enviar
    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_create_draft(
        subject: Annotated[str, Field(description="Asunto")],
        body: Annotated[str, Field(description="Cuerpo del mensaje")],
        to: Annotated[list[str], Field(description="Destinatarios: 'a@x.com' o 'Nombre <a@x.com>'")],
        cc: Annotated[Optional[list[str]], Field(None)] = None,
        bcc: Annotated[Optional[list[str]], Field(None)] = None,
        body_type: Annotated[str, Field("HTML", description="HTML o Text")] = "HTML",
        attachments: Annotated[Optional[list[dict]], Field(None, description="Adjuntos. Cada item: {\"path\": \"/ruta/fichero\"} (lo lee el servidor; PREFERIDO) o {name, contentType, contentBytes(base64)}. Los >3 MB van por upload session (hasta ~150 MB).")] = None,
        attachment_paths: Annotated[Optional[list[str]], Field(None, description="Atajo: rutas de ficheros a adjuntar. El SERVIDOR los lee y codifica; el base64 no pasa por el modelo (así no se corrompe). Para ficheros de tu ordenador usa el conector LOCAL.")] = None,
        attachment_urls: Annotated[Optional[list[str]], Field(None, description="URLs http(s) de ficheros a adjuntar (OneDrive/Drive/enlace público). El SERVIDOR las descarga; sirve también desde el VPS (no necesita ver tu disco).")] = None,
    ) -> dict:
        """Crea un BORRADOR en la carpeta Borradores (no lo envía).

        Devuelve el id del borrador; luego puedes revisarlo y enviarlo con
        mail_send_draft, o editarlo desde Outlook. Adjunta por RUTA
        (attachment_paths, conector local) o por ENLACE (attachment_urls, sirve
        también en VPS); grandes (>3 MB) por subida troceada.
        """
        helpers.guard_write()
        resolved = helpers.resolve_local(list(attachments or []) + [{"path": p} for p in (attachment_paths or [])])
        for u in (attachment_urls or []):
            resolved.append(await graph.fetch_attachment(u))
        small, large = helpers.classify(resolved)
        msg = {
            "subject": subject,
            "body": helpers.body(body, body_type),
            "toRecipients": helpers.recipients(to),
        }
        if cc:
            msg["ccRecipients"] = helpers.recipients(cc)
        if bcc:
            msg["bccRecipients"] = helpers.recipients(bcc)
        if small:
            msg["attachments"] = [
                {"@odata.type": "#microsoft.graph.fileAttachment", **a} for a in small
            ]
        res = await graph.request("POST", "/me/messages", json=msg)
        for a in large:
            await graph.upload_large_attachment(res["id"], a["name"], a["contentType"], a["data"])
        out = fmt.fmt_message(res, full=True)
        if large:
            out["large_attachments_added"] = len(large)
        return out

    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": False, "openWorldHint": True})
    async def mail_send(
        subject: Annotated[str, Field(description="Asunto")],
        body: Annotated[str, Field(description="Cuerpo del mensaje")],
        to: Annotated[list[str], Field(description="Destinatarios")],
        cc: Annotated[Optional[list[str]], Field(None)] = None,
        bcc: Annotated[Optional[list[str]], Field(None)] = None,
        body_type: Annotated[str, Field("HTML", description="HTML o Text")] = "HTML",
        attachments: Annotated[Optional[list[dict]], Field(None, description="Adjuntos. Cada item: {\"path\": \"/ruta/fichero\"} (lo lee el servidor; PREFERIDO) o {name, contentType, contentBytes(base64)}. Los >3 MB van por upload session (hasta ~150 MB).")] = None,
        attachment_paths: Annotated[Optional[list[str]], Field(None, description="Atajo: rutas de ficheros a adjuntar. El SERVIDOR los lee y codifica; el base64 no pasa por el modelo (así no se corrompe). Para ficheros de tu ordenador usa el conector LOCAL.")] = None,
        attachment_urls: Annotated[Optional[list[str]], Field(None, description="URLs http(s) de ficheros a adjuntar (OneDrive/Drive/enlace público). El SERVIDOR las descarga; sirve también desde el VPS (no necesita ver tu disco).")] = None,
        save_to_sent: Annotated[bool, Field(True)] = True,
        mailbox: Annotated[Optional[str], Field(None, description="Enviar DESDE/en nombre de un buzón COMPARTIDO (su email); por defecto el tuyo")] = None,
    ) -> dict:
        """ENVÍA un correo nuevo directamente (acción irreversible).

        Para revisar antes de enviar, usa mail_create_draft. Guarda copia en
        Enviados salvo que save_to_sent=False. Adjunta por RUTA (attachment_paths,
        conector local) o por ENLACE (attachment_urls, sirve también en VPS).
        Adjuntos grandes (>3 MB): se crea borrador, se suben por trozos y se envía.
        Con 'mailbox' envía desde un buzón compartido del despacho.
        """
        helpers.guard_write()
        root = helpers.mbox(mailbox)
        resolved = helpers.resolve_local(list(attachments or []) + [{"path": p} for p in (attachment_paths or [])])
        for u in (attachment_urls or []):
            resolved.append(await graph.fetch_attachment(u))
        small, large = helpers.classify(resolved)
        msg = {
            "subject": subject,
            "body": helpers.body(body, body_type),
            "toRecipients": helpers.recipients(to),
        }
        if cc:
            msg["ccRecipients"] = helpers.recipients(cc)
        if bcc:
            msg["bccRecipients"] = helpers.recipients(bcc)
        if small:
            msg["attachments"] = [
                {"@odata.type": "#microsoft.graph.fileAttachment", **a} for a in small
            ]
        if not large:
            # sin adjuntos grandes: envío directo en una sola llamada
            await graph.request("POST", f"{root}/sendMail", json={"message": msg, "saveToSentItems": save_to_sent})
            return {"status": "sent", "subject": subject, "to": to}
        # con adjuntos grandes: crear borrador, subirlos por trozos y enviar
        draft = await graph.request("POST", f"{root}/messages", json=msg)
        for a in large:
            await graph.upload_large_attachment(draft["id"], a["name"], a["contentType"], a["data"], base=root)
        await graph.request("POST", f"{root}/messages/{draft['id']}/send")
        return {"status": "sent", "subject": subject, "to": to, "large_attachments": len(large)}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_send_draft(
        message_id: Annotated[str, Field(description="Id del borrador a enviar")],
    ) -> dict:
        """Envía un borrador ya creado (acción irreversible)."""
        helpers.guard_write()
        await graph.request("POST", f"/me/messages/{message_id}/send")
        return {"status": "sent", "message_id": message_id}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_update_draft(
        message_id: Annotated[str, Field(description="Id del BORRADOR a editar")],
        subject: Annotated[Optional[str], Field(None, description="Nuevo asunto")] = None,
        body: Annotated[Optional[str], Field(None, description="Nuevo cuerpo del mensaje")] = None,
        body_type: Annotated[str, Field("HTML", description="HTML o Text")] = "HTML",
        to: Annotated[Optional[list[str]], Field(None, description="Reemplaza los destinatarios")] = None,
        cc: Annotated[Optional[list[str]], Field(None, description="Reemplaza la copia")] = None,
        bcc: Annotated[Optional[list[str]], Field(None, description="Reemplaza la copia oculta")] = None,
    ) -> dict:
        """Edita un BORRADOR ya existente: cambia asunto, cuerpo o destinatarios,
        CONSERVANDO sus adjuntos. Ideal cuando dejas un borrador con tus documentos
        adjuntos para que yo le redacte el texto y luego lo envíe con mail_send_draft.

        Solo funciona sobre borradores: Graph no permite editar el cuerpo/asunto de un
        correo ya enviado o recibido (para esos, usa mail_update: leído/categorías...).
        """
        helpers.guard_write()
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if body is not None:
            patch["body"] = helpers.body(body, body_type)
        if to is not None:
            patch["toRecipients"] = helpers.recipients(to)
        if cc is not None:
            patch["ccRecipients"] = helpers.recipients(cc)
        if bcc is not None:
            patch["bccRecipients"] = helpers.recipients(bcc)
        if not patch:
            raise ValueError("Indica al menos un campo a cambiar (subject, body, to, cc o bcc).")
        meta = await graph.request("GET", f"/me/messages/{message_id}", params={"$select": "isDraft"})
        if not meta.get("isDraft"):
            raise ValueError(
                "Ese correo NO es un borrador. Graph solo deja editar asunto/cuerpo/destinatarios "
                "de BORRADORES. Si quieres reenviarlo o responderlo, usa mail_forward / mail_reply."
            )
        res = await graph.request("PATCH", f"/me/messages/{message_id}", json=patch)
        return fmt.fmt_message(res, full=True)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_add_attachments(
        message_id: Annotated[str, Field(description="Id del BORRADOR al que añadir adjuntos")],
        attachment_paths: Annotated[Optional[list[str]], Field(None, description="Rutas de ficheros a adjuntar (conector local; el servidor los lee)")] = None,
        attachment_urls: Annotated[Optional[list[str]], Field(None, description="URLs http(s) de ficheros a adjuntar (el servidor las descarga; sirve también en VPS)")] = None,
        attachments: Annotated[Optional[list[dict]], Field(None, description="Adjuntos {path} o {name, contentType, contentBytes}")] = None,
    ) -> dict:
        """Añade adjuntos a un BORRADOR ya creado (por ruta local o por URL/OneDrive).

        Grandes (>3 MB) por subida troceada. Útil para completar un borrador antes de
        enviarlo con mail_send_draft.
        """
        helpers.guard_write()
        resolved = helpers.resolve_local(list(attachments or []) + [{"path": p} for p in (attachment_paths or [])])
        for u in (attachment_urls or []):
            resolved.append(await graph.fetch_attachment(u))
        small, large = helpers.classify(resolved)
        await graph.add_attachments(message_id, small, large)
        return {"message_id": message_id, "adjuntos_anadidos": len(small) + len(large)}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_reply(
        message_id: Annotated[str, Field(description="Id del mensaje al que responder")],
        comment: Annotated[str, Field(description="Texto de la respuesta")],
        reply_all: Annotated[bool, Field(False, description="Responder a todos")] = False,
        send: Annotated[bool, Field(True, description="True: envía; False: crea borrador de respuesta")] = True,
        attachment_paths: Annotated[Optional[list[str]], Field(None, description="Rutas de ficheros a adjuntar (conector local)")] = None,
        attachment_urls: Annotated[Optional[list[str]], Field(None, description="URLs http(s) de ficheros a adjuntar")] = None,
    ) -> dict:
        """Responde a un correo (o a todos), opcionalmente CON adjuntos.

        Sin adjuntos y send=True: respuesta directa en una llamada. Con adjuntos:
        se crea el borrador de respuesta, se adjuntan y se envía (o se deja en
        borradores si send=False).
        """
        helpers.guard_write()
        resolved = helpers.resolve_local([{"path": p} for p in (attachment_paths or [])])
        for u in (attachment_urls or []):
            resolved.append(await graph.fetch_attachment(u))
        small, large = helpers.classify(resolved)

        if not small and not large and send:
            action = "replyAll" if reply_all else "reply"
            await graph.request("POST", f"/me/messages/{message_id}/{action}", json={"comment": comment})
            return {"status": "sent", "message_id": message_id, "replyAll": reply_all}
        # con adjuntos o para dejar borrador: createReply(All) -> adjuntar -> (enviar)
        create = "createReplyAll" if reply_all else "createReply"
        draft = await graph.request("POST", f"/me/messages/{message_id}/{create}", json={"comment": comment})
        await graph.add_attachments(draft["id"], small, large)
        if send:
            await graph.request("POST", f"/me/messages/{draft['id']}/send")
            return {"status": "sent", "message_id": message_id, "replyAll": reply_all, "attachments": len(small) + len(large)}
        return {"status": "draft_created", "draft": fmt.fmt_message(draft)}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_forward(
        message_id: Annotated[str, Field(description="Id del mensaje a reenviar")],
        to: Annotated[list[str], Field(description="Destinatarios del reenvío")],
        comment: Annotated[str, Field("", description="Comentario a añadir")] = "",
        attachment_paths: Annotated[Optional[list[str]], Field(None, description="Rutas de ficheros a adjuntar además de los originales (conector local)")] = None,
        attachment_urls: Annotated[Optional[list[str]], Field(None, description="URLs http(s) de ficheros a adjuntar además de los originales")] = None,
    ) -> dict:
        """Reenvía un correo (conserva sus adjuntos), opcionalmente añadiendo más."""
        helpers.guard_write()
        resolved = helpers.resolve_local([{"path": p} for p in (attachment_paths or [])])
        for u in (attachment_urls or []):
            resolved.append(await graph.fetch_attachment(u))
        small, large = helpers.classify(resolved)

        if not small and not large:
            await graph.request(
                "POST", f"/me/messages/{message_id}/forward",
                json={"comment": comment, "toRecipients": helpers.recipients(to)},
            )
            return {"status": "forwarded", "message_id": message_id, "to": to}
        # con adjuntos extra: createForward -> adjuntar -> enviar
        draft = await graph.request(
            "POST", f"/me/messages/{message_id}/createForward",
            json={"comment": comment, "toRecipients": helpers.recipients(to)},
        )
        await graph.add_attachments(draft["id"], small, large)
        await graph.request("POST", f"/me/messages/{draft['id']}/send")
        return {"status": "forwarded", "message_id": message_id, "to": to, "extra_attachments": len(small) + len(large)}

    # ----------------------------------------------------------- marcar/organizar
    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True, "openWorldHint": True})
    async def mail_update(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        is_read: Annotated[Optional[bool], Field(None, description="Marcar leído/no leído")] = None,
        flag: Annotated[Optional[str], Field(None, description="Estado de seguimiento: flagged, complete, notFlagged")] = None,
        categories: Annotated[Optional[list[str]], Field(None, description="Reemplaza las categorías del mensaje")] = None,
        importance: Annotated[Optional[str], Field(None, description="low, normal, high")] = None,
    ) -> dict:
        """Actualiza propiedades de un correo: leído, seguimiento, categorías, importancia."""
        helpers.guard_write()
        patch: dict = {}
        if is_read is not None:
            patch["isRead"] = is_read
        if flag is not None:
            patch["flag"] = {"flagStatus": flag}
        if categories is not None:
            patch["categories"] = categories
        if importance is not None:
            patch["importance"] = importance
        if not patch:
            raise ValueError("Indica al menos un campo a actualizar.")
        res = await graph.request("PATCH", f"/me/messages/{message_id}", json=patch)
        return fmt.fmt_message(res)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_move(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        destination_folder: Annotated[str, Field(description="Id de carpeta o nombre conocido (archive, deleteditems, inbox...)")],
    ) -> dict:
        """Mueve un correo a otra carpeta (devuelve el mensaje con su nuevo id)."""
        helpers.guard_write()
        res = await graph.request("POST", f"/me/messages/{message_id}/move", json={"destinationId": destination_folder})
        return fmt.fmt_message(res)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def mail_copy(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        destination_folder: Annotated[str, Field(description="Id de carpeta o nombre conocido")],
    ) -> dict:
        """Copia un correo a otra carpeta (deja el original donde está)."""
        helpers.guard_write()
        res = await graph.request("POST", f"/me/messages/{message_id}/copy", json={"destinationId": destination_folder})
        return fmt.fmt_message(res)

    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def mail_delete(
        message_id: Annotated[str, Field(description="Id del mensaje")],
        permanent: Annotated[bool, Field(False, description="False: mueve a Elementos eliminados (reversible). True: borrado permanente.")] = False,
    ) -> dict:
        """Borra un correo. Por defecto lo manda a Papelera (Elementos eliminados).

        Con permanent=True se elimina de forma IRREVERSIBLE (DELETE directo).
        """
        helpers.guard_write()
        if permanent:
            await graph.request("DELETE", f"/me/messages/{message_id}")
            return {"status": "permanently_deleted", "message_id": message_id}
        res = await graph.request("POST", f"/me/messages/{message_id}/move", json={"destinationId": "deleteditems"})
        return {"status": "moved_to_trash", "message_id": res.get("id")}
