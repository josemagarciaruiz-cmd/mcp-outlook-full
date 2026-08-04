"""SharePoint avanzado: LISTAS (base de datos ligera del despacho: trackers de
asuntos, plazos, mini-CRM), METADATOS de documentos (etiquetar por cliente/materia/
estado y filtrar) y check-out/check-in. Scope: Sites.ReadWrite.All (ya lo hay).
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph


def register(mcp) -> None:
    # ------------------------- Listas de SharePoint ------------------------- #
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sp_lists(
        site_id: Annotated[str, Field(description="Id del sitio (de sharepoint_find_sites)")],
    ) -> dict:
        """Lista las LISTAS de un sitio de SharePoint (trackers, tablas de datos, no ficheros)."""
        res = await graph.request("GET", f"/sites/{site_id}/lists",
                                  params={"$select": "id,name,displayName,webUrl"})
        ls = res.get("value", []) if isinstance(res, dict) else []
        return {"listas": [{"id": l.get("id"), "name": l.get("displayName") or l.get("name"),
                            "webUrl": l.get("webUrl")} for l in ls]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sp_list_columns(
        site_id: Annotated[str, Field(description="Id del sitio")],
        list_id: Annotated[str, Field(description="Id de la lista")],
    ) -> dict:
        """Muestra las columnas (campos) de una lista, con su nombre interno para escribir."""
        res = await graph.request("GET", f"/sites/{site_id}/lists/{list_id}/columns")
        cols = res.get("value", []) if isinstance(res, dict) else []
        return {"columnas": [{"name": c.get("name"), "displayName": c.get("displayName"),
                              "readOnly": c.get("readOnly")} for c in cols if not c.get("hidden")]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sp_list_items(
        site_id: Annotated[str, Field(description="Id del sitio")],
        list_id: Annotated[str, Field(description="Id de la lista")],
        top: Annotated[int, Field(50, ge=1, le=200)] = 50,
    ) -> dict:
        """Lee los elementos (filas) de una lista con sus campos."""
        res = await graph.get_paged(f"/sites/{site_id}/lists/{list_id}/items",
                                    params={"$expand": "fields", "$top": top}, max_items=top)
        out = []
        for it in res["items"]:
            f = it.get("fields", {}) if isinstance(it, dict) else {}
            f = {k: v for k, v in f.items() if not k.startswith("@") and k not in ("id",)}
            out.append({"id": it.get("id"), "fields": f})
        return {"items": out, "next": res["next"]}

    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_list_item_create(
        site_id: Annotated[str, Field(description="Id del sitio")],
        list_id: Annotated[str, Field(description="Id de la lista")],
        fields: Annotated[dict, Field(description="Campos {nombre_interno: valor}, p.ej. {\"Title\":\"Asunto X\",\"Estado\":\"Abierto\"}")],
    ) -> dict:
        """Crea un elemento (fila) en una lista."""
        res = await graph.request("POST", f"/sites/{site_id}/lists/{list_id}/items", json={"fields": fields})
        return {"id": res.get("id"), "creado": True}

    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_list_item_update(
        site_id: Annotated[str, Field(description="Id del sitio")],
        list_id: Annotated[str, Field(description="Id de la lista")],
        item_id: Annotated[str, Field(description="Id del elemento")],
        fields: Annotated[dict, Field(description="Campos a cambiar {nombre_interno: valor}")],
    ) -> dict:
        """Actualiza los campos de un elemento de lista."""
        await graph.request("PATCH", f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields", json=fields)
        return {"id": item_id, "actualizado": True}

    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_list_item_delete(
        site_id: Annotated[str, Field(description="Id del sitio")],
        list_id: Annotated[str, Field(description="Id de la lista")],
        item_id: Annotated[str, Field(description="Id del elemento")],
    ) -> dict:
        """Borra un elemento de una lista."""
        await graph.request("DELETE", f"/sites/{site_id}/lists/{list_id}/items/{item_id}")
        return {"id": item_id, "borrado": True}

    # ------------------ Metadatos de documentos (columnas) ------------------ #
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def sp_file_fields_get(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        item_id: Annotated[str, Field(description="Id del documento")],
    ) -> dict:
        """Lee los METADATOS (columnas: cliente, materia, estado...) de un documento de SharePoint."""
        res = await graph.request("GET", f"/drives/{drive_id}/items/{item_id}/listItem/fields")
        f = {k: v for k, v in (res or {}).items() if not k.startswith("@")}
        return {"metadatos": f}

    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_file_fields_set(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        item_id: Annotated[str, Field(description="Id del documento")],
        fields: Annotated[dict, Field(description="Columnas a fijar {nombre_interno: valor}, p.ej. {\"Cliente\":\"Fulano\",\"Estado\":\"Presentado\"}")],
    ) -> dict:
        """Etiqueta un documento: fija sus columnas (cliente/materia/estado) para poder filtrar luego."""
        await graph.request("PATCH", f"/drives/{drive_id}/items/{item_id}/listItem/fields", json=fields)
        return {"id": item_id, "etiquetado": True}

    # ------------------------- Check-out / Check-in ------------------------- #
    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_checkout(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        item_id: Annotated[str, Field(description="Id del documento")],
    ) -> dict:
        """Bloquea un documento para edición (check-out): evita que otros lo pisen."""
        await graph.request("POST", f"/drives/{drive_id}/items/{item_id}/checkout")
        return {"id": item_id, "estado": "checked_out"}

    @mcp.tool(annotations={"openWorldHint": True})
    async def sp_checkin(
        drive_id: Annotated[str, Field(description="Id de la biblioteca (drive) de SharePoint")],
        item_id: Annotated[str, Field(description="Id del documento")],
        comment: Annotated[str, Field("", description="Comentario de la versión")] = "",
    ) -> dict:
        """Libera un documento (check-in) publicando los cambios, con comentario de versión."""
        await graph.request("POST", f"/drives/{drive_id}/items/{item_id}/checkin",
                            json={"comment": comment, "checkInAs": "published"})
        return {"id": item_id, "estado": "checked_in"}
