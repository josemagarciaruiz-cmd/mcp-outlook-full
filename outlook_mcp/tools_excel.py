"""Herramientas de Excel (workbook API de Graph): leer/escribir celdas, rangos y
tablas de una hoja guardada en OneDrive o SharePoint. Scope: Files.ReadWrite (ya lo hay).

Gobierna por voz la hoja de facturación, el listado de clientes, el cuadro de plazos...
Todas aceptan drive_id (vacío = tu OneDrive) para trabajar sobre SharePoint.
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph


def _base(drive_id: Optional[str]) -> str:
    return f"/drives/{drive_id}" if drive_id else "/me/drive"


def register(mcp) -> None:
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def excel_worksheets(
        item_id: Annotated[str, Field(description="Id del fichero .xlsx")],
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Lista las hojas de un libro de Excel."""
        res = await graph.request("GET", f"{_base(drive_id)}/items/{item_id}/workbook/worksheets")
        ws = res.get("value", []) if isinstance(res, dict) else []
        return {"hojas": [{"name": w.get("name"), "position": w.get("position"), "visibility": w.get("visibility")} for w in ws]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def excel_read(
        item_id: Annotated[str, Field(description="Id del fichero .xlsx")],
        sheet: Annotated[str, Field(description="Nombre de la hoja")],
        address: Annotated[Optional[str], Field(None, description="Rango A1, p.ej. 'A1:D20'; vacío = rango usado")] = None,
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Lee valores de un rango de una hoja (o del rango usado si no se indica)."""
        base = f"{_base(drive_id)}/items/{item_id}/workbook/worksheets('{sheet}')"
        path = f"{base}/range(address='{address}')" if address else f"{base}/usedRange"
        res = await graph.request("GET", path, params={"$select": "address,values,text"})
        return {"address": res.get("address"), "values": res.get("values")}

    @mcp.tool(annotations={"openWorldHint": True})
    async def excel_write(
        item_id: Annotated[str, Field(description="Id del fichero .xlsx")],
        sheet: Annotated[str, Field(description="Nombre de la hoja")],
        address: Annotated[str, Field(description="Rango A1 a escribir, p.ej. 'A2:C2'")],
        values: Annotated[list[list], Field(description="Matriz de valores (filas x columnas) del tamaño del rango")],
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Escribe valores en un rango de una hoja (sobrescribe)."""
        base = f"{_base(drive_id)}/items/{item_id}/workbook/worksheets('{sheet}')"
        res = await graph.request("PATCH", f"{base}/range(address='{address}')", json={"values": values})
        return {"address": res.get("address"), "escrito": True}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def excel_tables(
        item_id: Annotated[str, Field(description="Id del fichero .xlsx")],
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Lista las tablas (con nombre) del libro."""
        res = await graph.request("GET", f"{_base(drive_id)}/items/{item_id}/workbook/tables")
        ts = res.get("value", []) if isinstance(res, dict) else []
        return {"tablas": [{"name": t.get("name"), "id": t.get("id")} for t in ts]}

    @mcp.tool(annotations={"openWorldHint": True})
    async def excel_add_row(
        item_id: Annotated[str, Field(description="Id del fichero .xlsx")],
        table: Annotated[str, Field(description="Nombre de la tabla")],
        values: Annotated[list[list], Field(description="Filas a añadir, p.ej. [[\"Cliente\", 100, \"Pendiente\"]]")],
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Añade una o varias filas al final de una tabla (ideal para facturación, registros)."""
        base = f"{_base(drive_id)}/items/{item_id}/workbook/tables('{table}')/rows"
        res = await graph.request("POST", base, json={"values": values})
        return {"añadidas": len(values), "index": res.get("index")}
