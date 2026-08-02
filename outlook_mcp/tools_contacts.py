"""Herramientas de Contactos de Outlook. Scope: Contacts.ReadWrite."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph, helpers

_SELECT = "id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle"


def _fmt(c: dict) -> dict:
    return {
        "id": c.get("id"),
        "name": c.get("displayName"),
        "emails": [e.get("address") for e in (c.get("emailAddresses") or []) if e.get("address")],
        "phones": (c.get("businessPhones") or []) + ([c["mobilePhone"]] if c.get("mobilePhone") else []),
        "company": c.get("companyName"),
        "jobTitle": c.get("jobTitle"),
    }


def register(mcp) -> None:
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def contacts_list_folders() -> dict:
        """Lista tus carpetas de contactos (si organizas los contactos en carpetas)."""
        res = await graph.request("GET", "/me/contactFolders", params={"$select": "id,displayName,parentFolderId"})
        return {"folders": [{"id": f.get("id"), "name": f.get("displayName")}
                            for f in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def contacts_list(
        top: Annotated[int, Field(50, ge=1, le=100)] = 50,
        folder: Annotated[Optional[str], Field(None, description="Id de carpeta de contactos (de contacts_list_folders); por defecto todos")] = None,
    ) -> dict:
        """Lista tus contactos (nombre, emails, teléfonos, empresa)."""
        path = f"/me/contactFolders/{folder}/contacts" if folder else "/me/contacts"
        res = await graph.get_paged(
            path, params={"$top": top, "$orderby": "displayName", "$select": _SELECT}, max_items=top
        )
        return {"contacts": [_fmt(c) for c in res["items"]], "next": res["next"]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def contacts_search(
        query: Annotated[str, Field(description="Texto (nombre, email, empresa…)")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Busca contactos por nombre, email o empresa."""
        res = await graph.request(
            "GET", "/me/contacts",
            params={"$search": f'"{query}"', "$top": top, "$select": _SELECT},
            headers={"ConsistencyLevel": "eventual"},
        )
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"contacts": [_fmt(c) for c in items]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def contacts_get(
        contact_id: Annotated[str, Field(description="Id del contacto")],
    ) -> dict:
        """Devuelve un contacto completo."""
        res = await graph.request("GET", f"/me/contacts/{contact_id}", params={"$select": _SELECT})
        return _fmt(res)

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def contacts_create(
        given_name: Annotated[str, Field(description="Nombre")],
        surname: Annotated[Optional[str], Field(None, description="Apellidos")] = None,
        emails: Annotated[Optional[list[str]], Field(None, description="Direcciones de correo")] = None,
        mobile_phone: Annotated[Optional[str], Field(None)] = None,
        business_phones: Annotated[Optional[list[str]], Field(None)] = None,
        company: Annotated[Optional[str], Field(None, description="Empresa")] = None,
        job_title: Annotated[Optional[str], Field(None, description="Cargo")] = None,
    ) -> dict:
        """Crea un contacto nuevo."""
        helpers.guard_write()
        c: dict = {"givenName": given_name}
        if surname:
            c["surname"] = surname
        if emails:
            c["emailAddresses"] = [{"address": e, "name": e} for e in emails]
        if mobile_phone:
            c["mobilePhone"] = mobile_phone
        if business_phones:
            c["businessPhones"] = business_phones
        if company:
            c["companyName"] = company
        if job_title:
            c["jobTitle"] = job_title
        res = await graph.request("POST", "/me/contacts", json=c)
        return _fmt(res)

    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True, "openWorldHint": True})
    async def contacts_update(
        contact_id: Annotated[str, Field(description="Id del contacto")],
        given_name: Annotated[Optional[str], Field(None)] = None,
        surname: Annotated[Optional[str], Field(None)] = None,
        emails: Annotated[Optional[list[str]], Field(None, description="Reemplaza las direcciones de correo")] = None,
        mobile_phone: Annotated[Optional[str], Field(None)] = None,
        company: Annotated[Optional[str], Field(None)] = None,
        job_title: Annotated[Optional[str], Field(None)] = None,
    ) -> dict:
        """Modifica un contacto (solo los campos indicados)."""
        helpers.guard_write()
        patch: dict = {}
        if given_name is not None:
            patch["givenName"] = given_name
        if surname is not None:
            patch["surname"] = surname
        if emails is not None:
            patch["emailAddresses"] = [{"address": e, "name": e} for e in emails]
        if mobile_phone is not None:
            patch["mobilePhone"] = mobile_phone
        if company is not None:
            patch["companyName"] = company
        if job_title is not None:
            patch["jobTitle"] = job_title
        if not patch:
            raise ValueError("Indica al menos un campo a modificar.")
        res = await graph.request("PATCH", f"/me/contacts/{contact_id}", json=patch)
        return _fmt(res)

    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def contacts_delete(
        contact_id: Annotated[str, Field(description="Id del contacto")],
    ) -> dict:
        """Elimina un contacto."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/contacts/{contact_id}")
        return {"status": "deleted", "contact_id": contact_id}
