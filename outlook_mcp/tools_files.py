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
