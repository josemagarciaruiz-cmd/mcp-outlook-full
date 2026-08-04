"""Herramientas de OneDrive: listar, buscar, descargar, subir y compartir.
Scope: Files.ReadWrite. Sinergia: el downloadUrl/enlace sirve para attachment_urls.
"""

from __future__ import annotations

import mimetypes
from pathlib import Path
from typing import Annotated, Optional

from pydantic import Field

from . import graph, helpers


def _fmt_item(i: dict) -> dict:
    return {
        "id": i.get("id"),
        "name": i.get("name"),
        "isFolder": "folder" in i,
        "size": i.get("size"),
        "childCount": (i.get("folder") or {}).get("childCount") if "folder" in i else None,
        "lastModified": i.get("lastModifiedDateTime"),
        "webUrl": i.get("webUrl"),
        "downloadUrl": i.get("@microsoft.graph.downloadUrl"),
    }


def register(mcp) -> None:
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_list(
        folder_id: Annotated[Optional[str], Field(None, description="Id de carpeta; por defecto la raíz de OneDrive")] = None,
        top: Annotated[int, Field(50, ge=1, le=200)] = 50,
    ) -> dict:
        """Lista ficheros y carpetas de OneDrive (raíz o una carpeta)."""
        path = f"/me/drive/items/{folder_id}/children" if folder_id else "/me/drive/root/children"
        res = await graph.get_paged(path, params={"$top": top}, max_items=top)
        return {"items": [_fmt_item(i) for i in res["items"]], "next": res["next"]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_search(
        query: Annotated[str, Field(description="Texto a buscar en nombre/contenido de tus ficheros")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Busca ficheros en OneDrive por nombre o contenido."""
        q = query.replace("'", "''")
        res = await graph.request("GET", f"/me/drive/root/search(q='{q}')", params={"$top": top})
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"items": [_fmt_item(i) for i in items]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_get(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
    ) -> dict:
        """Metadatos de un fichero, incluido su downloadUrl (usable en attachment_urls)."""
        res = await graph.request("GET", f"/me/drive/items/{item_id}")
        return _fmt_item(res)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def files_download(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
        dest_dir: Annotated[str, Field(description="Carpeta LOCAL donde guardar (conector local)")],
        filename: Annotated[Optional[str], Field(None, description="Nombre del fichero; por defecto el de OneDrive")] = None,
    ) -> dict:
        """Descarga un fichero de OneDrive al disco local (conector local)."""
        helpers.guard_write()
        meta = await graph.request("GET", f"/me/drive/items/{item_id}", params={"$select": "name"})
        data = await graph.download_item(item_id)
        d = Path(dest_dir).expanduser()
        d.mkdir(parents=True, exist_ok=True)
        p = d / (filename or meta.get("name") or "descarga")
        p.write_bytes(data)
        return {"path": str(p), "bytes": len(data)}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def files_upload(
        local_path: Annotated[str, Field(description="Ruta del fichero LOCAL a subir (conector local)")],
        dest_folder: Annotated[str, Field("", description="Carpeta de OneDrive destino, p.ej. 'Documentos' (vacío = raíz)")] = "",
        name: Annotated[Optional[str], Field(None, description="Nombre en OneDrive; por defecto el del fichero")] = None,
    ) -> dict:
        """Sube un fichero local a OneDrive (grandes por sesión de subida)."""
        helpers.guard_write()
        p = Path(local_path).expanduser()
        if not p.is_file():
            raise FileNotFoundError(f"No existe el fichero local '{local_path}'.")
        data = p.read_bytes()
        target = name or p.name
        dest = f"{dest_folder.strip('/')}/{target}" if dest_folder.strip("/") else target
        ctype = mimetypes.guess_type(target)[0] or "application/octet-stream"
        res = await graph.onedrive_upload(dest, data, ctype)
        return _fmt_item(res)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def files_create_link(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
        link_type: Annotated[str, Field("view", description="view | edit")] = "view",
        scope: Annotated[str, Field("anonymous", description="anonymous (cualquiera con el enlace) | organization")] = "anonymous",
    ) -> dict:
        """Crea un enlace para compartir un fichero de OneDrive."""
        helpers.guard_write()
        res = await graph.request(
            "POST", f"/me/drive/items/{item_id}/createLink",
            json={"type": link_type, "scope": scope},
        )
        link = res.get("link", {}) if isinstance(res, dict) else {}
        return {"webUrl": link.get("webUrl"), "type": link.get("type"), "scope": link.get("scope")}

    # ----------------------------------------------------------------- #
    # Gestión de ficheros/carpetas
    # ----------------------------------------------------------------- #
    @mcp.tool(annotations={"openWorldHint": True})
    async def files_create_folder(
        name: Annotated[str, Field(description="Nombre de la nueva carpeta")],
        parent_id: Annotated[Optional[str], Field(None, description="Carpeta padre; vacío = raíz de OneDrive")] = None,
    ) -> dict:
        """Crea una carpeta en OneDrive."""
        helpers.guard_write()
        path = f"/me/drive/items/{parent_id}/children" if parent_id else "/me/drive/root/children"
        res = await graph.request("POST", path, json={
            "name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"})
        return _fmt_item(res)

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_move(
        item_id: Annotated[str, Field(description="Id del item a mover")],
        new_parent_id: Annotated[str, Field(description="Id de la carpeta destino")],
        new_name: Annotated[Optional[str], Field(None, description="Renombrar al mover (opcional)")] = None,
    ) -> dict:
        """Mueve un fichero o carpeta a otra carpeta."""
        helpers.guard_write()
        body = {"parentReference": {"id": new_parent_id}}
        if new_name:
            body["name"] = new_name
        res = await graph.request("PATCH", f"/me/drive/items/{item_id}", json=body)
        return _fmt_item(res)

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_copy(
        item_id: Annotated[str, Field(description="Id del item a copiar")],
        new_parent_id: Annotated[str, Field(description="Id de la carpeta destino")],
        new_name: Annotated[Optional[str], Field(None, description="Nombre de la copia (opcional)")] = None,
    ) -> dict:
        """Copia un fichero o carpeta (operación asíncrona en Graph)."""
        helpers.guard_write()
        body = {"parentReference": {"id": new_parent_id}}
        if new_name:
            body["name"] = new_name
        await graph.request("POST", f"/me/drive/items/{item_id}/copy", json=body)
        return {"status": "copia_en_proceso", "destino": new_parent_id}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_rename(
        item_id: Annotated[str, Field(description="Id del item")],
        new_name: Annotated[str, Field(description="Nuevo nombre")],
    ) -> dict:
        """Renombra un fichero o carpeta."""
        helpers.guard_write()
        res = await graph.request("PATCH", f"/me/drive/items/{item_id}", json={"name": new_name})
        return _fmt_item(res)

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_delete(
        item_id: Annotated[str, Field(description="Id del item")],
    ) -> dict:
        """Envía un fichero o carpeta a la PAPELERA de OneDrive (reversible)."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/drive/items/{item_id}")
        return {"status": "a_papelera", "item_id": item_id}

    # ----------------------------------------------------------------- #
    # Versiones y conversión
    # ----------------------------------------------------------------- #
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_list_versions(
        item_id: Annotated[str, Field(description="Id del item")],
    ) -> dict:
        """Lista las versiones anteriores de un fichero (para recuperar un cambio)."""
        res = await graph.request("GET", f"/me/drive/items/{item_id}/versions")
        vs = res.get("value", []) if isinstance(res, dict) else []
        return {"versions": [
            {"id": v.get("id"), "lastModified": v.get("lastModifiedDateTime"), "size": v.get("size")}
            for v in vs]}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_restore_version(
        item_id: Annotated[str, Field(description="Id del item")],
        version_id: Annotated[str, Field(description="Id de la versión (de files_list_versions)")],
    ) -> dict:
        """Restaura una versión anterior de un fichero como versión actual."""
        helpers.guard_write()
        await graph.request("POST", f"/me/drive/items/{item_id}/versions/{version_id}/restoreVersion")
        return {"status": "version_restaurada", "version_id": version_id}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_save_pdf(
        item_id: Annotated[str, Field(description="Id del documento (Word/Excel/PPT/etc.)")],
        dest_folder: Annotated[str, Field("", description="Carpeta destino en OneDrive (vacío = raíz)")] = "",
        pdf_name: Annotated[Optional[str], Field(None, description="Nombre del PDF (por defecto, el del original)")] = None,
    ) -> dict:
        """Convierte un documento a PDF y guarda el PDF en OneDrive (nativo de Graph)."""
        helpers.guard_write()
        meta = await graph.request("GET", f"/me/drive/items/{item_id}", params={"$select": "name"})
        data = await graph.get_content(f"/me/drive/items/{item_id}/content?format=pdf")
        base = (pdf_name or meta.get("name") or "documento").rsplit(".", 1)[0] + ".pdf"
        dest = f"{dest_folder.strip('/')}/{base}" if dest_folder.strip("/") else base
        res = await graph.onedrive_upload(dest, data, "application/pdf")
        return _fmt_item(res)

    # ----------------------------------------------------------------- #
    # Compartición con control y auditoría (confidencialidad / RGPD)
    # ----------------------------------------------------------------- #
    @mcp.tool(annotations={"openWorldHint": True})
    async def files_share(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
        link_type: Annotated[str, Field("view", description="view | edit")] = "view",
        scope: Annotated[str, Field("organization", description="organization (solo del despacho) | anonymous (cualquiera con enlace)")] = "organization",
        expiration_date: Annotated[Optional[str], Field(None, description="Caducidad ISO, p.ej. 2026-09-30T23:59:59Z")] = None,
        password: Annotated[Optional[str], Field(None, description="Contraseña del enlace (según plan de M365)")] = None,
    ) -> dict:
        """Crea un enlace de compartir CON CONTROL: caducidad y/o contraseña."""
        helpers.guard_write()
        body: dict = {"type": link_type, "scope": scope}
        if expiration_date:
            body["expirationDateTime"] = expiration_date
        if password:
            body["password"] = password
        res = await graph.request("POST", f"/me/drive/items/{item_id}/createLink", json=body)
        link = res.get("link", {}) if isinstance(res, dict) else {}
        return {"webUrl": link.get("webUrl"), "type": link.get("type"), "scope": link.get("scope"),
                "expira": res.get("expirationDateTime")}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_invite(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
        emails: Annotated[list[str], Field(description="Correos de las personas a invitar")],
        role: Annotated[str, Field("read", description="read | write")] = "read",
        message: Annotated[Optional[str], Field(None, description="Mensaje de la invitación")] = None,
        require_sign_in: Annotated[bool, Field(True, description="Exigir inicio de sesión")] = True,
        expiration_date: Annotated[Optional[str], Field(None, description="Caducidad del acceso ISO (opcional)")] = None,
    ) -> dict:
        """Da acceso a personas CONCRETAS por email (lectura o edición), con caducidad opcional."""
        helpers.guard_write()
        body: dict = {
            "recipients": [{"email": e} for e in emails],
            "roles": [role],
            "requireSignIn": require_sign_in,
            "sendInvitation": True,
        }
        if message:
            body["message"] = message
        if expiration_date:
            body["expirationDateTime"] = expiration_date
        res = await graph.request("POST", f"/me/drive/items/{item_id}/invite", json=body)
        granted = res.get("value", []) if isinstance(res, dict) else []
        return {"invitados": len(granted), "role": role}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_list_permissions(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
    ) -> dict:
        """Muestra QUIÉN tiene acceso a un fichero y cómo (enlaces, personas, alcance)."""
        res = await graph.request("GET", f"/me/drive/items/{item_id}/permissions")
        perms = res.get("value", []) if isinstance(res, dict) else []
        out = []
        for p in perms:
            link = p.get("link") or {}
            ident = (p.get("grantedToV2") or p.get("grantedTo") or {}).get("user") or {}
            out.append({
                "id": p.get("id"), "roles": p.get("roles"),
                "tipo": "enlace" if link else "persona",
                "scope": link.get("scope") if link else None,
                "url": link.get("webUrl") if link else None,
                "persona": ident.get("displayName") or ident.get("email"),
                "expira": p.get("expirationDateTime"),
            })
        return {"permisos": out}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_remove_permission(
        item_id: Annotated[str, Field(description="Id del item de OneDrive")],
        permission_id: Annotated[str, Field(description="Id del permiso (de files_list_permissions)")],
    ) -> dict:
        """Revoca un acceso concreto (corta un enlace o quita a una persona)."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/drive/items/{item_id}/permissions/{permission_id}")
        return {"status": "acceso_revocado", "permission_id": permission_id}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_audit_shares(
        folder_id: Annotated[Optional[str], Field(None, description="Carpeta a auditar; vacío = raíz de OneDrive")] = None,
        top: Annotated[int, Field(100, ge=1, le=200)] = 100,
    ) -> dict:
        """Audita una carpeta y señala los ficheros COMPARTIDOS (sobre todo enlaces 'cualquiera'/externos).

        Recorre los items de la carpeta y revisa sus permisos; marca los que salen del
        despacho. Útil para revisión de confidencialidad/RGPD.
        """
        path = f"/me/drive/items/{folder_id}/children" if folder_id else "/me/drive/root/children"
        res = await graph.get_paged(path, params={"$top": top}, max_items=top)
        riesgos = []
        for it in res["items"]:
            try:
                pr = await graph.request("GET", f"/me/drive/items/{it['id']}/permissions")
            except Exception:
                continue
            perms = pr.get("value", []) if isinstance(pr, dict) else []
            expuestos = [p for p in perms if (p.get("link") or {}).get("scope") == "anonymous"
                         or "external" in str(p.get("grantedToV2", "")).lower()]
            if expuestos:
                riesgos.append({"name": it.get("name"), "id": it.get("id"),
                                "accesos_externos": len(expuestos),
                                "webUrl": it.get("webUrl")})
        return {"auditados": len(res["items"]), "con_exposicion": riesgos,
                "aviso": "Revisa 'con_exposicion'; usa files_remove_permission para cortar accesos."}

    # ----------------------------------------------------------------- #
    # SharePoint: bibliotecas de documentos del despacho (scope Sites.*)
    # ----------------------------------------------------------------- #
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sharepoint_find_sites(
        query: Annotated[str, Field(description="Texto para buscar el sitio/equipo de SharePoint")],
    ) -> dict:
        """Busca sitios de SharePoint del despacho (repositorios de equipo)."""
        q = query.replace("'", "''")
        res = await graph.request("GET", "/sites", params={"search": q})
        sites = res.get("value", []) if isinstance(res, dict) else []
        return {"sites": [{"id": s.get("id"), "name": s.get("displayName") or s.get("name"),
                           "webUrl": s.get("webUrl")} for s in sites]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sharepoint_list_libraries(
        site_id: Annotated[str, Field(description="Id del sitio (de sharepoint_find_sites)")],
    ) -> dict:
        """Lista las bibliotecas de documentos (drives) de un sitio de SharePoint."""
        res = await graph.request("GET", f"/sites/{site_id}/drives")
        drives = res.get("value", []) if isinstance(res, dict) else []
        return {"bibliotecas": [{"id": d.get("id"), "name": d.get("name"),
                                 "webUrl": d.get("webUrl")} for d in drives]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sharepoint_list_items(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        folder_id: Annotated[Optional[str], Field(None, description="Carpeta; vacío = raíz de la biblioteca")] = None,
        top: Annotated[int, Field(50, ge=1, le=200)] = 50,
    ) -> dict:
        """Lista ficheros y carpetas de una biblioteca de SharePoint."""
        tail = f"items/{folder_id}/children" if folder_id else "root/children"
        res = await graph.get_paged(f"/drives/{drive_id}/{tail}", params={"$top": top}, max_items=top)
        return {"items": [_fmt_item(i) for i in res["items"]], "next": res["next"]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sharepoint_search(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        query: Annotated[str, Field(description="Texto a buscar")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Busca ficheros dentro de una biblioteca de SharePoint (incluye grabaciones de canal)."""
        q = query.replace("'", "''")
        res = await graph.request("GET", f"/drives/{drive_id}/root/search(q='{q}')", params={"$top": top})
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"items": [_fmt_item(i) for i in items]}
