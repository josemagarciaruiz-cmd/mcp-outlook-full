#!/usr/bin/env python3
"""
Agente Outlook (local) — el "puente con el disco" del conector Outlook.

Complementa al conector completo `outlook`: resuelve lo que una dirección remota
(o el propio modelo) no puede hacer con seguridad sobre TU disco:
  - adjuntar a un correo ficheros de tu ordenador y enviarlos/responder/reenviar,
  - guardar los adjuntos de un correo en una carpeta de tu disco,
  - exportar un correo a .eml a tu disco,
  - puente OneDrive <-> disco (subir/bajar ficheros).

Reutiliza el núcleo probado del conector (outlook_mcp: auth MSAL, Graph, adjuntos
pequeños/grandes, OneDrive). Por seguridad solo toca las carpetas de
OUTLOOK_ALLOWED_DIRS. Transporte: stdio (conector local en Claude Desktop).

Autor: José María García Ruiz · josemaria.ai
"""

from __future__ import annotations

import base64
import mimetypes
import os
import sys
from pathlib import Path
from typing import Annotated, Optional

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from outlook_mcp import auth, config, graph, helpers

# --------------------------------------------------------------------------- #
# Seguridad de rutas: el agente SOLO puede tocar OUTLOOK_ALLOWED_DIRS
# --------------------------------------------------------------------------- #

def _allowed_dirs() -> list[str]:
    raw = os.environ.get("OUTLOOK_ALLOWED_DIRS", str(Path.home()))
    raw = raw.replace(";", ":")
    dirs = []
    for p in raw.split(":"):
        p = p.strip()
        if p:
            dirs.append(os.path.realpath(os.path.expanduser(os.path.expandvars(p))))
    return dirs or [str(Path.home())]


ALLOWED = _allowed_dirs()


def _safe(path: str, *, must_exist_file: bool = False, make_dir: bool = False) -> str:
    """Valida que la ruta cae dentro de las carpetas permitidas y la devuelve."""
    rp = os.path.realpath(os.path.expanduser(os.path.expandvars(path)))
    if not any(rp == a or rp.startswith(a + os.sep) for a in ALLOWED):
        raise ValueError(
            f"Ruta fuera de las carpetas permitidas: {path}. "
            f"Permitidas: {ALLOWED} (configúralas en OUTLOOK_ALLOWED_DIRS)."
        )
    if must_exist_file and not os.path.isfile(rp):
        raise FileNotFoundError(f"No existe el fichero: {path}")
    if make_dir:
        os.makedirs(rp, exist_ok=True)
    return rp


def _inline(small: list[dict]) -> list[dict]:
    return [{"@odata.type": "#microsoft.graph.fileAttachment", **a} for a in small]


INSTRUCTIONS = """\
Agente Outlook local: el puente con el disco del conector `outlook`.
Todas las herramientas empiezan por outlook_local_ y SOLO tocan las carpetas de
OUTLOOK_ALLOWED_DIRS. Úsalo cuando haya que leer/escribir ficheros del ORDENADOR:
adjuntar ficheros locales a un correo (send/reply/forward), guardar adjuntos al
disco, exportar .eml, o mover ficheros entre disco y OneDrive. Para todo lo demás
(leer/buscar correo, agenda, tareas...), usa el conector `outlook`.
Enviar correo sale al exterior: confirma con el usuario antes.
"""


def build_agent() -> MCPServer:
    mcp = MCPServer("agente-outlook", instructions=INSTRUCTIONS, version="1.0.0")

    @mcp.tool(annotations={"readOnlyHint": True})
    async def outlook_local_allowed_dirs() -> dict:
        """Muestra las carpetas del disco que el Agente Outlook tiene permitido tocar."""
        return {"allowed": ALLOWED}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_send(
        to: Annotated[list[str], Field(description="Destinatarios")],
        subject: Annotated[str, Field(description="Asunto")],
        body: Annotated[str, Field(description="Cuerpo del mensaje")],
        file_paths: Annotated[Optional[list[str]], Field(None, description="Rutas de ficheros del disco a adjuntar")] = None,
        cc: Annotated[Optional[list[str]], Field(None, description="Copia")] = None,
        bcc: Annotated[Optional[list[str]], Field(None, description="Copia oculta")] = None,
        body_type: Annotated[str, Field("HTML", description="HTML o Text")] = "HTML",
        save_to_sent: Annotated[bool, Field(True)] = True,
    ) -> dict:
        """Envía un correo adjuntando ficheros de TU disco (los lee el agente, el
        base64 no pasa por el modelo). Grandes (>3 MB) por subida troceada."""
        paths = [_safe(p, must_exist_file=True) for p in (file_paths or [])]
        small, large = helpers.classify(helpers.resolve_local([{"path": p} for p in paths]))
        msg = {"subject": subject, "body": helpers.body(body, body_type),
               "toRecipients": helpers.recipients(to)}
        if cc:
            msg["ccRecipients"] = helpers.recipients(cc)
        if bcc:
            msg["bccRecipients"] = helpers.recipients(bcc)
        if small:
            msg["attachments"] = _inline(small)
        if not large:
            await graph.request("POST", "/me/sendMail",
                                json={"message": msg, "saveToSentItems": save_to_sent})
            return {"status": "sent", "to": to, "attachments": len(small)}
        draft = await graph.request("POST", "/me/messages", json=msg)
        for a in large:
            await graph.upload_large_attachment(draft["id"], a["name"], a["contentType"], a["data"])
        await graph.request("POST", f"/me/messages/{draft['id']}/send")
        return {"status": "sent", "to": to, "attachments": len(small) + len(large)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_reply(
        message_id: Annotated[str, Field(description="Id del correo a responder")],
        comment: Annotated[str, Field(description="Texto de la respuesta")],
        file_paths: Annotated[Optional[list[str]], Field(None, description="Ficheros del disco a adjuntar")] = None,
        reply_all: Annotated[bool, Field(False)] = False,
        send: Annotated[bool, Field(True, description="Enviar; si False, deja borrador")] = True,
    ) -> dict:
        """Responde a un correo adjuntando ficheros de tu disco."""
        paths = [_safe(p, must_exist_file=True) for p in (file_paths or [])]
        small, large = helpers.classify(helpers.resolve_local([{"path": p} for p in paths]))
        create = "createReplyAll" if reply_all else "createReply"
        draft = await graph.request("POST", f"/me/messages/{message_id}/{create}",
                                    json={"comment": comment})
        await graph.add_attachments(draft["id"], small, large)
        if send:
            await graph.request("POST", f"/me/messages/{draft['id']}/send")
            return {"status": "sent", "message_id": message_id, "attachments": len(small) + len(large)}
        return {"status": "draft_created", "draft_id": draft.get("id")}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_forward(
        message_id: Annotated[str, Field(description="Id del correo a reenviar")],
        to: Annotated[list[str], Field(description="Destinatarios")],
        comment: Annotated[str, Field("", description="Comentario opcional")] = "",
        file_paths: Annotated[Optional[list[str]], Field(None, description="Ficheros del disco a añadir")] = None,
    ) -> dict:
        """Reenvía un correo (conserva sus adjuntos) añadiendo ficheros de tu disco."""
        paths = [_safe(p, must_exist_file=True) for p in (file_paths or [])]
        small, large = helpers.classify(helpers.resolve_local([{"path": p} for p in paths]))
        draft = await graph.request("POST", f"/me/messages/{message_id}/createForward",
                                    json={"comment": comment, "toRecipients": helpers.recipients(to)})
        await graph.add_attachments(draft["id"], small, large)
        await graph.request("POST", f"/me/messages/{draft['id']}/send")
        return {"status": "forwarded", "to": to, "extra_attachments": len(small) + len(large)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_save_attachments(
        message_id: Annotated[str, Field(description="Id del correo")],
        dest_dir: Annotated[str, Field(description="Carpeta del disco donde guardar")],
        attachment_id: Annotated[Optional[str], Field(None, description="Un adjunto concreto; si se omite, todos")] = None,
    ) -> dict:
        """Descarga y GUARDA en tu disco los adjuntos de un correo."""
        d = _safe(dest_dir, make_dir=True)
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
            p = os.path.join(d, (a.get("name") or "adjunto").replace("/", "-"))
            with open(p, "wb") as fh:
                fh.write(base64.b64decode(cb))
            saved.append({"name": os.path.basename(p), "path": p, "size": a.get("size")})
        return {"saved": saved, "dest_dir": d}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_export_eml(
        message_id: Annotated[str, Field(description="Id del correo")],
        dest_dir: Annotated[str, Field(description="Carpeta del disco de destino")],
        filename: Annotated[Optional[str], Field(None, description="Nombre del .eml")] = None,
    ) -> dict:
        """Exporta un correo completo como fichero .eml a tu disco."""
        d = _safe(dest_dir, make_dir=True)
        raw = await graph.request("GET", f"/me/messages/{message_id}/$value")
        if not isinstance(raw, (bytes, bytearray)):
            raw = str(raw).encode("utf-8")
        name = (filename or "correo").replace("/", "-").replace(":", "-")
        if not name.lower().endswith(".eml"):
            name += ".eml"
        p = os.path.join(d, name)
        with open(p, "wb") as fh:
            fh.write(raw)
        return {"path": p, "bytes": len(raw)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_onedrive_download(
        item_id: Annotated[str, Field(description="Id del fichero en OneDrive")],
        dest_dir: Annotated[str, Field(description="Carpeta del disco de destino")],
        filename: Annotated[Optional[str], Field(None, description="Nombre a usar; por defecto el de OneDrive")] = None,
    ) -> dict:
        """Descarga un fichero de OneDrive a tu disco."""
        d = _safe(dest_dir, make_dir=True)
        meta = await graph.request("GET", f"/me/drive/items/{item_id}", params={"$select": "name"})
        data = await graph.download_item(item_id)
        p = os.path.join(d, (filename or meta.get("name") or "descarga").replace("/", "-"))
        with open(p, "wb") as fh:
            fh.write(data)
        return {"path": p, "bytes": len(data)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def outlook_local_onedrive_upload(
        local_path: Annotated[str, Field(description="Ruta del fichero del disco a subir")],
        dest_folder: Annotated[str, Field("", description="Carpeta destino en OneDrive; raíz si se omite")] = "",
        name: Annotated[Optional[str], Field(None, description="Nombre a usar en OneDrive")] = None,
    ) -> dict:
        """Sube un fichero de tu disco a OneDrive."""
        p = _safe(local_path, must_exist_file=True)
        with open(p, "rb") as fh:
            data = fh.read()
        target = name or os.path.basename(p)
        dest = f"{dest_folder.strip('/')}/{target}" if dest_folder.strip("/") else target
        ctype = mimetypes.guess_type(target)[0] or "application/octet-stream"
        res = await graph.onedrive_upload(dest, data, ctype)
        return {"name": res.get("name"), "id": res.get("id"), "webUrl": res.get("webUrl")}

    return mcp


# --------------------------------------------------------------------------- #
# Arranque / login
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        try:
            config.require_client_id()
            info = auth.device_login()
        except Exception as exc:  # mensaje limpio, sin traceback
            sys.exit(f"ERROR: {exc}")
        print(f"\n✓ Sesión iniciada como {info.get('name')} <{info.get('account')}>",
              file=sys.stderr)
        sys.exit(0)
    build_agent().run(transport="stdio")
