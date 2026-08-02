"""Búsqueda de personas: contactos frecuentes (/me/people) y directorio de la
organización (/users). Scopes: People.Read y User.ReadBasic.All.
"""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from . import graph


def register(mcp) -> None:
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def people_search(
        query: Annotated[str, Field(description="Nombre, email o parte de él de la persona que buscas")],
        top: Annotated[int, Field(10, ge=1, le=50)] = 10,
    ) -> dict:
        """Busca entre las personas con las que MÁS te relacionas (correo/reuniones).

        Es la vía más rápida para encontrar el email de un cliente o colaborador
        con el que ya te escribes, aunque no lo tengas como contacto formal.
        """
        res = await graph.request(
            "GET", "/me/people",
            params={"$search": f'"{query}"', "$top": top,
                    "$select": "displayName,scoredEmailAddresses,jobTitle,companyName,personType"},
        )
        out = []
        for p in (res.get("value", []) if isinstance(res, dict) else []):
            emails = [e.get("address") for e in (p.get("scoredEmailAddresses") or []) if e.get("address")]
            out.append({
                "name": p.get("displayName"),
                "emails": emails,
                "jobTitle": p.get("jobTitle"),
                "company": p.get("companyName"),
                "type": (p.get("personType") or {}).get("subclass"),
            })
        return {"people": out}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def directory_search(
        query: Annotated[str, Field(description="Nombre o email a buscar en el directorio del despacho/organización")],
        top: Annotated[int, Field(15, ge=1, le=50)] = 15,
    ) -> dict:
        """Busca personas en el DIRECTORIO de la organización (compañeros del despacho)."""
        res = await graph.request(
            "GET", "/users",
            params={"$search": f'"displayName:{query}" OR "mail:{query}"', "$top": top,
                    "$select": "displayName,mail,userPrincipalName,jobTitle,department,mobilePhone"},
            headers={"ConsistencyLevel": "eventual"},
        )
        out = []
        for u in (res.get("value", []) if isinstance(res, dict) else []):
            out.append({
                "name": u.get("displayName"),
                "email": u.get("mail") or u.get("userPrincipalName"),
                "jobTitle": u.get("jobTitle"),
                "department": u.get("department"),
                "mobile": u.get("mobilePhone"),
            })
        return {"users": out}
