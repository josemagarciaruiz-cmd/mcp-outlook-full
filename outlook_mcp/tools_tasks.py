"""Herramientas de Microsoft To-Do (tareas y plazos). Scope: Tasks.ReadWrite."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph, helpers


def _fmt_task(t: dict) -> dict:
    due = (t.get("dueDateTime") or {}).get("dateTime")
    return {
        "id": t.get("id"),
        "title": t.get("title"),
        "status": t.get("status"),  # notStarted, inProgress, completed...
        "due": due,
        "importance": t.get("importance"),
        "reminder": (t.get("reminderDateTime") or {}).get("dateTime"),
        "hasBody": bool((t.get("body") or {}).get("content")),
    }


def register(mcp) -> None:
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def tasks_list_lists() -> dict:
        """Lista tus listas de tareas de Microsoft To-Do (id y nombre)."""
        res = await graph.request("GET", "/me/todo/lists")
        return {"lists": [{"id": l.get("id"), "name": l.get("displayName"),
                           "isDefault": l.get("wellknownListName") == "defaultList"}
                          for l in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def tasks_list(
        list_id: Annotated[str, Field(description="Id de la lista (de tasks_list_lists)")],
        top: Annotated[int, Field(50, ge=1, le=100)] = 50,
        include_completed: Annotated[bool, Field(False, description="Incluir también las completadas")] = False,
    ) -> dict:
        """Lista las tareas de una lista (por defecto solo las pendientes)."""
        params = {"$top": top, "$orderby": "createdDateTime desc"}
        if not include_completed:
            params["$filter"] = "status ne 'completed'"
        res = await graph.request("GET", f"/me/todo/lists/{list_id}/tasks", params=params)
        return {"tasks": [_fmt_task(t) for t in (res.get("value", []) if isinstance(res, dict) else [])]}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def tasks_create(
        list_id: Annotated[str, Field(description="Id de la lista donde crear la tarea")],
        title: Annotated[str, Field(description="Título de la tarea")],
        due: Annotated[Optional[str], Field(None, description="Vencimiento ISO 8601 (p.ej. 2026-08-05T00:00:00)")] = None,
        body: Annotated[Optional[str], Field(None, description="Notas/detalle de la tarea")] = None,
        importance: Annotated[Optional[str], Field(None, description="low | normal | high")] = None,
        reminder: Annotated[Optional[str], Field(None, description="Recordatorio ISO 8601")] = None,
    ) -> dict:
        """Crea una tarea (con vencimiento y recordatorio opcionales)."""
        helpers.guard_write()
        task: dict = {"title": title}
        if due:
            task["dueDateTime"] = helpers.date_time_zone(due)
        if body:
            task["body"] = {"content": body, "contentType": "text"}
        if importance:
            task["importance"] = importance
        if reminder:
            task["reminderDateTime"] = helpers.date_time_zone(reminder)
            task["isReminderOn"] = True
        res = await graph.request("POST", f"/me/todo/lists/{list_id}/tasks", json=task)
        return _fmt_task(res)

    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True, "openWorldHint": True})
    async def tasks_update(
        list_id: Annotated[str, Field(description="Id de la lista")],
        task_id: Annotated[str, Field(description="Id de la tarea")],
        title: Annotated[Optional[str], Field(None)] = None,
        due: Annotated[Optional[str], Field(None, description="Nuevo vencimiento ISO 8601")] = None,
        body: Annotated[Optional[str], Field(None)] = None,
        importance: Annotated[Optional[str], Field(None, description="low | normal | high")] = None,
        status: Annotated[Optional[str], Field(None, description="notStarted | inProgress | completed")] = None,
    ) -> dict:
        """Modifica una tarea (título, vencimiento, notas, importancia o estado)."""
        helpers.guard_write()
        patch: dict = {}
        if title is not None:
            patch["title"] = title
        if due is not None:
            patch["dueDateTime"] = helpers.date_time_zone(due)
        if body is not None:
            patch["body"] = {"content": body, "contentType": "text"}
        if importance is not None:
            patch["importance"] = importance
        if status is not None:
            patch["status"] = status
        if not patch:
            raise ValueError("Indica al menos un campo a modificar.")
        res = await graph.request("PATCH", f"/me/todo/lists/{list_id}/tasks/{task_id}", json=patch)
        return _fmt_task(res)

    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True, "openWorldHint": True})
    async def tasks_complete(
        list_id: Annotated[str, Field(description="Id de la lista")],
        task_id: Annotated[str, Field(description="Id de la tarea")],
    ) -> dict:
        """Marca una tarea como completada."""
        helpers.guard_write()
        res = await graph.request("PATCH", f"/me/todo/lists/{list_id}/tasks/{task_id}", json={"status": "completed"})
        return _fmt_task(res)

    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def tasks_delete(
        list_id: Annotated[str, Field(description="Id de la lista")],
        task_id: Annotated[str, Field(description="Id de la tarea")],
    ) -> dict:
        """Elimina una tarea."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/todo/lists/{list_id}/tasks/{task_id}")
        return {"status": "deleted", "task_id": task_id}
