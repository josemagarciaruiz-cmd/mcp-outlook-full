"""OneDrive/SharePoint avanzado (parte 2): búsqueda unificada, novedades (delta),
recientes, compartidos conmigo, navegar por ruta, árbol, previsualización,
actualizar permisos, plantillas, subir desde URL, crear ruta, operaciones por lote,
analítica y copia entre OneDrive y SharePoint. Scopes ya existentes (Files.ReadWrite,
Sites.ReadWrite.All).
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph, helpers


def _fmt(i: dict) -> dict:
    return {
        "id": i.get("id"), "name": i.get("name"),
        "isFolder": "folder" in i, "size": i.get("size"),
        "lastModified": i.get("lastModifiedDateTime"),
        "webUrl": i.get("webUrl"),
        "downloadUrl": i.get("@microsoft.graph.downloadUrl"),
        "parentDriveId": (i.get("parentReference") or {}).get("driveId"),
    }


def register(mcp) -> None:
    # ------------------------- Búsqueda y descubrimiento ------------------------- #
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_search_all(
        query: Annotated[str, Field(description="Texto a buscar en TODO (OneDrive + todos los sitios de SharePoint)")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Búsqueda UNIFICADA de documentos en tu OneDrive y en todo SharePoint a la vez."""
        body = {"requests": [{"entityTypes": ["driveItem"],
                              "query": {"queryString": query}, "size": top}]}
        res = await graph.request("POST", "/search/query", json=body)
        hits = []
        try:
            for cont in res["value"][0]["hitsContainers"]:
                for h in cont.get("hits", []):
                    r = h.get("resource", {})
                    hits.append({"name": r.get("name"), "id": r.get("id"),
                                 "webUrl": r.get("webUrl"),
                                 "driveId": (r.get("parentReference") or {}).get("driveId"),
                                 "summary": h.get("summary")})
        except Exception:
            pass
        return {"resultados": hits}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_delta(
        folder_id: Annotated[Optional[str], Field(None, description="Carpeta a vigilar; vacío = raíz de OneDrive")] = None,
        since_link: Annotated[Optional[str], Field(None, description="deltaLink de una llamada anterior para traer SOLO lo nuevo")] = None,
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """NOVEDADES: qué ha cambiado desde la última vez. Guarda 'deltaLink' y pásalo como since_link la próxima."""
        if since_link:
            res = await graph.get_json(since_link)
        else:
            base = f"/drives/{drive_id}" if drive_id else "/me/drive"
            tail = f"items/{folder_id}/delta" if folder_id else "root/delta"
            res = await graph.get_json(f"{base}/{tail}")
        items = [_fmt(i) for i in res.get("value", [])]
        return {"cambios": items, "deltaLink": res.get("@odata.deltaLink"),
                "nextLink": res.get("@odata.nextLink")}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_recent(
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Ficheros que has usado recientemente."""
        res = await graph.request("GET", "/me/drive/recent", params={"$top": top})
        return {"recientes": [_fmt(i) for i in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_shared_with_me(
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Ficheros y carpetas que OTROS han compartido contigo."""
        res = await graph.request("GET", "/me/drive/sharedWithMe", params={"$top": top})
        return {"compartidos_conmigo": [_fmt(i) for i in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_get_by_path(
        path: Annotated[str, Field(description="Ruta desde la raíz, p.ej. 'Clientes/Fulano/Contrato.docx'")],
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Localiza un fichero o carpeta por su RUTA (sin necesitar el id)."""
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        p = path.strip("/")
        res = await graph.request("GET", f"{base}/root:/{p}")
        return _fmt(res)

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_tree(
        folder_id: Annotated[Optional[str], Field(None, description="Carpeta raíz del árbol; vacío = raíz de OneDrive")] = None,
        max_depth: Annotated[int, Field(2, ge=1, le=4, description="Profundidad máxima")] = 2,
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Árbol de carpetas/ficheros hasta cierta profundidad (recorrido recursivo acotado)."""
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"

        async def walk(fid: Optional[str], depth: int) -> list:
            tail = f"items/{fid}/children" if fid else "root/children"
            res = await graph.get_paged(f"{base}/{tail}", params={"$top": 200}, max_items=200)
            out = []
            for it in res["items"]:
                node = {"name": it.get("name"), "id": it.get("id"), "isFolder": "folder" in it}
                if "folder" in it and depth < max_depth:
                    node["children"] = await walk(it["id"], depth + 1)
                out.append(node)
            return out

        return {"arbol": await walk(folder_id, 1)}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_preview(
        item_id: Annotated[str, Field(description="Id del fichero")],
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Enlace de PREVISUALIZACIÓN incrustable (solo lectura, temporal) de un documento."""
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        res = await graph.request("POST", f"{base}/items/{item_id}/preview", json={})
        return {"getUrl": res.get("getUrl"), "postUrl": res.get("postUrl")}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def files_analytics(
        item_id: Annotated[str, Field(description="Id del fichero")],
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Actividad de un documento: nº de vistas y de usuarios (histórico). Puede no estar disponible en OneDrive personal."""
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        try:
            res = await graph.request("GET", f"{base}/items/{item_id}/analytics/allTime")
            act = (res.get("access") or {})
            return {"vistas": act.get("actionCount"), "usuarios": act.get("actorCount")}
        except Exception as e:
            return {"disponible": False, "motivo": str(e)[:160]}

    # ------------------------- Gestión avanzada ------------------------- #
    @mcp.tool(annotations={"openWorldHint": True})
    async def files_update_permission(
        item_id: Annotated[str, Field(description="Id del item")],
        permission_id: Annotated[str, Field(description="Id del permiso (de files_list_permissions)")],
        role: Annotated[Optional[str], Field(None, description="read | write (para cambiar el nivel)")] = None,
        expiration_date: Annotated[Optional[str], Field(None, description="Nueva caducidad ISO (para ampliar/reducir)")] = None,
    ) -> dict:
        """Modifica un permiso existente: cambia lectura/edición o su caducidad."""
        helpers.guard_write()
        body: dict = {}
        if role:
            body["roles"] = [role]
        if expiration_date:
            body["expirationDateTime"] = expiration_date
        res = await graph.request("PATCH", f"/me/drive/items/{item_id}/permissions/{permission_id}", json=body)
        return {"id": res.get("id"), "roles": res.get("roles"), "expira": res.get("expirationDateTime")}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_from_template(
        template_item_id: Annotated[str, Field(description="Id del documento plantilla")],
        dest_folder_id: Annotated[str, Field(description="Carpeta destino")],
        new_name: Annotated[str, Field(description="Nombre del nuevo documento")],
    ) -> dict:
        """Crea un documento nuevo a partir de una PLANTILLA (copia y renombra)."""
        helpers.guard_write()
        await graph.request("POST", f"/me/drive/items/{template_item_id}/copy",
                            json={"parentReference": {"id": dest_folder_id}, "name": new_name})
        return {"status": "creado_desde_plantilla", "nombre": new_name}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_upload_from_url(
        url: Annotated[str, Field(description="URL pública del fichero a traer")],
        dest_folder: Annotated[str, Field("", description="Carpeta destino en OneDrive (vacío = raíz)")] = "",
        name: Annotated[Optional[str], Field(None, description="Nombre en OneDrive; por defecto, el del enlace")] = None,
    ) -> dict:
        """Trae un fichero desde una URL y lo sube a OneDrive (funciona también desde el VPS)."""
        helpers.guard_write()
        fname, ctype, data = await graph.fetch_attachment(url)
        target = name or fname or "descarga"
        dest = f"{dest_folder.strip('/')}/{target}" if dest_folder.strip("/") else target
        res = await graph.onedrive_upload(dest, data, ctype)
        return _fmt(res)

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_ensure_path(
        path: Annotated[str, Field(description="Ruta de carpetas a garantizar, p.ej. 'Clientes/Fulano/2026'")],
        drive_id: Annotated[Optional[str], Field(None, description="Drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Crea la ruta de carpetas completa si no existe (como 'mkdir -p')."""
        helpers.guard_write()
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        parent = None
        acc = ""
        for seg in [s for s in path.strip("/").split("/") if s]:
            acc = f"{acc}/{seg}" if acc else seg
            try:
                found = await graph.request("GET", f"{base}/root:/{acc}")
                parent = found.get("id")
            except Exception:
                tail = f"items/{parent}/children" if parent else "root/children"
                created = await graph.request("POST", f"{base}/{tail}", json={
                    "name": seg, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
                parent = created.get("id")
        return {"folder_id": parent, "path": path}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_batch_delete(
        item_ids: Annotated[list[str], Field(description="Ids de items a enviar a papelera (máx. 20)")],
    ) -> dict:
        """Envía VARIOS ficheros/carpetas a la papelera de una vez (lote)."""
        helpers.guard_write()
        reqs = [{"id": str(n), "method": "DELETE", "url": f"/me/drive/items/{i}"}
                for n, i in enumerate(item_ids[:20])]
        res = await graph.request("POST", "/$batch", json={"requests": reqs})
        ok = sum(1 for r in res.get("responses", []) if 200 <= r.get("status", 500) < 300)
        return {"enviados_a_papelera": ok, "total": len(reqs)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_batch_move(
        item_ids: Annotated[list[str], Field(description="Ids de items a mover (máx. 20)")],
        dest_folder_id: Annotated[str, Field(description="Carpeta destino")],
    ) -> dict:
        """Mueve VARIOS ficheros a una carpeta de una vez (lote)."""
        helpers.guard_write()
        reqs = [{"id": str(n), "method": "PATCH", "url": f"/me/drive/items/{i}",
                 "headers": {"Content-Type": "application/json"},
                 "body": {"parentReference": {"id": dest_folder_id}}}
                for n, i in enumerate(item_ids[:20])]
        res = await graph.request("POST", "/$batch", json={"requests": reqs})
        ok = sum(1 for r in res.get("responses", []) if 200 <= r.get("status", 500) < 300)
        return {"movidos": ok, "total": len(reqs)}

    @mcp.tool(annotations={"openWorldHint": True})
    async def files_copy_to_drive(
        item_id: Annotated[str, Field(description="Id del item de origen (tu OneDrive)")],
        dest_drive_id: Annotated[str, Field(description="Id del drive/biblioteca destino (SharePoint)")],
        dest_folder_id: Annotated[str, Field(description="Id de la carpeta destino en ese drive")],
        new_name: Annotated[Optional[str], Field(None, description="Nombre en destino (opcional)")] = None,
    ) -> dict:
        """Copia un documento ENTRE drives (p.ej. de tu OneDrive a una biblioteca de SharePoint)."""
        helpers.guard_write()
        body = {"parentReference": {"driveId": dest_drive_id, "id": dest_folder_id}}
        if new_name:
            body["name"] = new_name
        await graph.request("POST", f"/me/drive/items/{item_id}/copy", json=body)
        return {"status": "copia_en_proceso", "destino_drive": dest_drive_id}
