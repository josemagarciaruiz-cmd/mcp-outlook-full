"""Herramientas de CALENDARIO de Outlook (Microsoft Graph).

Cobertura completa: calendarios, ver agenda por rango, crear/editar/borrar
eventos, reuniones de Teams, recurrencias, responder invitaciones, reenviar,
cancelar, buscar huecos (findMeetingTimes) y disponibilidad (getSchedule).
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import config
from . import format as fmt
from . import graph, helpers

_EVENT_SELECT = (
    "id,subject,start,end,isAllDay,location,organizer,attendees,showAs,"
    "isCancelled,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,responseStatus,"
    "seriesMasterId,type,webLink"
)
_EVENT_FULL = _EVENT_SELECT + ",body,categories,importance,sensitivity,recurrence,reminderMinutesBeforeStart,hasAttachments"


def register(mcp) -> None:
    # ------------------------------------------------------------- calendarios
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_list_calendars() -> dict:
        """Lista tus calendarios (principal y secundarios) con id y color."""
        res = await graph.request(
            "GET", "/me/calendars", params={"$select": "id,name,color,isDefaultCalendar,canEdit,owner"}
        )
        items = res.get("value", []) if isinstance(res, dict) else []
        return {"calendars": [fmt.fmt_calendar(c) for c in items]}

    # ------------------------------------------------------------ ver agenda
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_list_events(
        start: Annotated[str, Field(description="Inicio del rango ISO 8601, p.ej. 2026-07-29T00:00:00")],
        end: Annotated[str, Field(description="Fin del rango ISO 8601, p.ej. 2026-08-05T23:59:59")],
        calendar_id: Annotated[Optional[str], Field(None, description="Id de calendario; por defecto el principal")] = None,
        user: Annotated[Optional[str], Field(None, description="Email de otra persona/buzón cuyo calendario COMPARTIDO ver; por defecto el tuyo")] = None,
        top: Annotated[int, Field(50, ge=1, le=100)] = 50,
    ) -> dict:
        """Muestra la agenda en un rango de fechas (expande series recurrentes).

        Usa calendarView: devuelve TODAS las ocurrencias que caen en la ventana,
        incluidas las de eventos recurrentes. Con 'user' ve el calendario compartido
        de otra persona. Ordenado por hora de inicio.
        """
        root = f"/users/{user}" if user else "/me"
        base = f"{root}/calendars/{calendar_id}/calendarView" if calendar_id else f"{root}/calendarView"
        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$select": _EVENT_SELECT,
            "$orderby": "start/dateTime",
            "$top": top,
        }
        headers = graph.prefer_headers()
        res = await graph.get_paged(base, params=params, headers=headers, max_items=top)
        return {"events": [fmt.fmt_event(e) for e in res["items"]], "next": res["next"], "timezone": config.TIMEZONE}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_get_event(
        event_id: Annotated[str, Field(description="Id del evento")],
        body_as_text: Annotated[bool, Field(True)] = True,
    ) -> dict:
        """Lee un evento completo: cuerpo, asistentes, recurrencia, recordatorio."""
        headers = graph.prefer_headers(text_body=body_as_text)
        res = await graph.request("GET", f"/me/events/{event_id}", params={"$select": _EVENT_FULL}, headers=headers)
        return fmt.fmt_event(res, full=True)

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_list_instances(
        event_id: Annotated[str, Field(description="Id del evento maestro de una serie recurrente")],
        start: Annotated[str, Field(description="Inicio del rango ISO 8601")],
        end: Annotated[str, Field(description="Fin del rango ISO 8601")],
        top: Annotated[int, Field(50, ge=1, le=100)] = 50,
    ) -> dict:
        """Lista las ocurrencias concretas de un evento recurrente en un rango."""
        params = {"startDateTime": start, "endDateTime": end, "$select": _EVENT_SELECT, "$top": top}
        res = await graph.get_paged(
            f"/me/events/{event_id}/instances", params=params, headers=graph.prefer_headers(), max_items=top
        )
        return {"instances": [fmt.fmt_event(e) for e in res["items"]], "next": res["next"]}

    # ----------------------------------------------------------- crear/editar
    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def calendar_create_event(
        subject: Annotated[str, Field(description="Título del evento")],
        start: Annotated[str, Field(description="Inicio ISO 8601 (sin zona; se usa OUTLOOK_TIMEZONE), p.ej. 2026-07-30T10:00:00")],
        end: Annotated[str, Field(description="Fin ISO 8601")],
        timezone: Annotated[Optional[str], Field(None, description="Zona IANA; por defecto la del servidor (Europe/Madrid)")] = None,
        body: Annotated[Optional[str], Field(None, description="Descripción/cuerpo")] = None,
        body_type: Annotated[str, Field("HTML")] = "HTML",
        location: Annotated[Optional[str], Field(None, description="Ubicación (texto)")] = None,
        attendees: Annotated[Optional[list[str]], Field(None, description="Invitados: 'a@x.com' o 'Nombre <a@x.com>'")] = None,
        optional_attendees: Annotated[Optional[list[str]], Field(None, description="Invitados opcionales")] = None,
        is_online_meeting: Annotated[bool, Field(False, description="Crear reunión de Teams con enlace")] = False,
        is_all_day: Annotated[bool, Field(False, description="Evento de todo el día (start/end a medianoche)")] = False,
        reminder_minutes: Annotated[Optional[int], Field(None, description="Minutos de aviso antes del inicio")] = None,
        categories: Annotated[Optional[list[str]], Field(None)] = None,
        show_as: Annotated[Optional[str], Field(None, description="free, tentative, busy, oof, workingElsewhere")] = None,
        importance: Annotated[Optional[str], Field(None, description="low, normal, high")] = None,
        recurrence: Annotated[Optional[dict], Field(None, description="Objeto recurrence de Graph {pattern, range}. Ver README.")] = None,
        calendar_id: Annotated[Optional[str], Field(None, description="Crear el evento en un calendario concreto (de calendar_list_calendars); por defecto el principal")] = None,
    ) -> dict:
        """Crea un evento o reunión. Si hay asistentes, envía las invitaciones.

        Soporta Teams (is_online_meeting), todo el día, recordatorio, categorías
        y recurrencia (pasa el objeto 'recurrence' de Graph para series).
        """
        helpers.guard_write()
        ev: dict = {
            "subject": subject,
            "start": helpers.date_time_zone(start, timezone),
            "end": helpers.date_time_zone(end, timezone),
        }
        if body is not None:
            ev["body"] = helpers.body(body, body_type)
        if location:
            ev["location"] = {"displayName": location}
        att = []
        for a in helpers.recipients(attendees):
            att.append({**a, "type": "required"})
        for a in helpers.recipients(optional_attendees):
            att.append({**a, "type": "optional"})
        if att:
            ev["attendees"] = att
        if is_online_meeting:
            ev["isOnlineMeeting"] = True
            ev["onlineMeetingProvider"] = "teamsForBusiness"
        if is_all_day:
            ev["isAllDay"] = True
        if reminder_minutes is not None:
            ev["reminderMinutesBeforeStart"] = reminder_minutes
            ev["isReminderOn"] = True
        if categories:
            ev["categories"] = categories
        if show_as:
            ev["showAs"] = show_as
        if importance:
            ev["importance"] = importance
        if recurrence:
            ev["recurrence"] = recurrence
        ev_path = f"/me/calendars/{calendar_id}/events" if calendar_id else "/me/events"
        res = await graph.request("POST", ev_path, json=ev, headers=graph.prefer_headers())
        return fmt.fmt_event(res, full=True)

    @mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True, "openWorldHint": True})
    async def calendar_update_event(
        event_id: Annotated[str, Field(description="Id del evento")],
        subject: Annotated[Optional[str], Field(None)] = None,
        start: Annotated[Optional[str], Field(None, description="Nuevo inicio ISO 8601")] = None,
        end: Annotated[Optional[str], Field(None, description="Nuevo fin ISO 8601")] = None,
        timezone: Annotated[Optional[str], Field(None)] = None,
        body: Annotated[Optional[str], Field(None)] = None,
        body_type: Annotated[str, Field("HTML")] = "HTML",
        location: Annotated[Optional[str], Field(None)] = None,
        attendees: Annotated[Optional[list[str]], Field(None, description="Reemplaza la lista de asistentes requeridos")] = None,
        reminder_minutes: Annotated[Optional[int], Field(None)] = None,
        categories: Annotated[Optional[list[str]], Field(None)] = None,
        show_as: Annotated[Optional[str], Field(None)] = None,
    ) -> dict:
        """Modifica un evento existente (solo los campos que indiques).

        Si el evento tiene asistentes, Outlook envía la actualización a todos.
        """
        helpers.guard_write()
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if start is not None:
            patch["start"] = helpers.date_time_zone(start, timezone)
        if end is not None:
            patch["end"] = helpers.date_time_zone(end, timezone)
        if body is not None:
            patch["body"] = helpers.body(body, body_type)
        if location is not None:
            patch["location"] = {"displayName": location}
        if attendees is not None:
            patch["attendees"] = [{**a, "type": "required"} for a in helpers.recipients(attendees)]
        if reminder_minutes is not None:
            patch["reminderMinutesBeforeStart"] = reminder_minutes
        if categories is not None:
            patch["categories"] = categories
        if show_as is not None:
            patch["showAs"] = show_as
        if not patch:
            raise ValueError("Indica al menos un campo a modificar.")
        res = await graph.request("PATCH", f"/me/events/{event_id}", json=patch, headers=graph.prefer_headers())
        return fmt.fmt_event(res, full=True)

    # ------------------------------------------------------- responder/reenviar
    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def calendar_respond_event(
        event_id: Annotated[str, Field(description="Id del evento/invitación")],
        response: Annotated[str, Field(description="accept, decline o tentative")],
        comment: Annotated[str, Field("", description="Comentario opcional para el organizador")] = "",
        send_response: Annotated[bool, Field(True, description="Enviar la respuesta al organizador")] = True,
        proposed_new_start: Annotated[Optional[str], Field(None, description="Proponer nueva hora de inicio ISO (solo tentative/decline)")] = None,
        proposed_new_end: Annotated[Optional[str], Field(None, description="Proponer nueva hora de fin ISO")] = None,
    ) -> dict:
        """Responde a una invitación: aceptar, rechazar o provisional (tentative).

        Opcionalmente propone un nuevo horario (allowNewTimeProposals).
        """
        helpers.guard_write()
        action = {"accept": "accept", "decline": "decline", "tentative": "tentativelyAccept"}.get(response.lower())
        if not action:
            raise ValueError("response debe ser accept, decline o tentative.")
        payload: dict = {"comment": comment, "sendResponse": send_response}
        if proposed_new_start and proposed_new_end:
            payload["proposedNewTime"] = {
                "start": helpers.date_time_zone(proposed_new_start),
                "end": helpers.date_time_zone(proposed_new_end),
            }
        await graph.request("POST", f"/me/events/{event_id}/{action}", json=payload)
        return {"status": f"responded_{response.lower()}", "event_id": event_id}

    @mcp.tool(annotations={"readOnlyHint": False, "openWorldHint": True})
    async def calendar_forward_event(
        event_id: Annotated[str, Field(description="Id del evento")],
        to: Annotated[list[str], Field(description="Destinatarios a los que reenviar la invitación")],
        comment: Annotated[str, Field("")] = "",
    ) -> dict:
        """Reenvía una invitación de reunión a otras personas."""
        helpers.guard_write()
        await graph.request(
            "POST", f"/me/events/{event_id}/forward",
            json={"comment": comment, "toRecipients": helpers.recipients(to)},
        )
        return {"status": "forwarded", "event_id": event_id, "to": to}

    # --------------------------------------------------------- cancelar/borrar
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def calendar_cancel_event(
        event_id: Annotated[str, Field(description="Id del evento (debes ser el organizador)")],
        comment: Annotated[str, Field("", description="Mensaje de cancelación a los asistentes")] = "",
    ) -> dict:
        """Cancela una reunión que organizas TÚ y avisa a los asistentes.

        Para eventos sin asistentes o de los que no eres organizador, usa
        calendar_delete_event.
        """
        helpers.guard_write()
        await graph.request("POST", f"/me/events/{event_id}/cancel", json={"comment": comment})
        return {"status": "cancelled", "event_id": event_id}

    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "openWorldHint": True})
    async def calendar_delete_event(
        event_id: Annotated[str, Field(description="Id del evento")],
    ) -> dict:
        """Elimina un evento de tu calendario (si eres organizador con asistentes,
        preferible calendar_cancel_event para que se les notifique)."""
        helpers.guard_write()
        await graph.request("DELETE", f"/me/events/{event_id}")
        return {"status": "deleted", "event_id": event_id}

    # ------------------------------------------------ disponibilidad / huecos
    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_find_meeting_times(
        attendees: Annotated[list[str], Field(description="Asistentes requeridos (emails)")],
        duration_minutes: Annotated[int, Field(30, ge=15, le=480)] = 30,
        start: Annotated[Optional[str], Field(None, description="Inicio de la ventana de búsqueda ISO 8601")] = None,
        end: Annotated[Optional[str], Field(None, description="Fin de la ventana ISO 8601")] = None,
        max_candidates: Annotated[int, Field(10, ge=1, le=100)] = 10,
    ) -> dict:
        """Sugiere huecos libres para reunir a varias personas (findMeetingTimes)."""
        payload: dict = {
            "attendees": [{"emailAddress": r["emailAddress"], "type": "required"} for r in helpers.recipients(attendees)],
            "meetingDuration": f"PT{duration_minutes}M",
            "maxCandidates": max_candidates,
            "isOrganizerOptional": False,
            "returnSuggestionReasons": True,
            "minimumAttendeePercentage": 100,
        }
        if start and end:
            payload["timeConstraint"] = {
                "activityDomain": "work",
                "timeSlots": [{
                    "start": helpers.date_time_zone(start),
                    "end": helpers.date_time_zone(end),
                }],
            }
        res = await graph.request(
            "POST", "/me/findMeetingTimes", json=payload,
            headers=graph.prefer_headers(),
        )
        suggestions = (res or {}).get("meetingTimeSuggestions", [])
        out = []
        for s in suggestions:
            slot = s.get("meetingTimeSlot", {})
            out.append({
                "start": (slot.get("start") or {}).get("dateTime"),
                "end": (slot.get("end") or {}).get("dateTime"),
                "confidence": s.get("confidence"),
                "availability": s.get("suggestionReason"),
            })
        return {"suggestions": out, "emptyReason": (res or {}).get("emptySuggestionsReason")}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def calendar_get_schedule(
        schedules: Annotated[list[str], Field(description="Emails cuyos libre/ocupado consultar (incluye el tuyo si quieres)")],
        start: Annotated[str, Field(description="Inicio ISO 8601")],
        end: Annotated[str, Field(description="Fin ISO 8601")],
        interval_minutes: Annotated[int, Field(30, ge=5, le=1440)] = 30,
    ) -> dict:
        """Consulta la disponibilidad (libre/ocupado) de varias personas."""
        payload = {
            "schedules": schedules,
            "startTime": helpers.date_time_zone(start),
            "endTime": helpers.date_time_zone(end),
            "availabilityViewInterval": interval_minutes,
        }
        res = await graph.request("POST", "/me/calendar/getSchedule", json=payload, headers=graph.prefer_headers())
        items = (res or {}).get("value", [])
        out = []
        for s in items:
            out.append({
                "email": s.get("scheduleId"),
                "availabilityView": s.get("availabilityView"),  # cadena: 0=libre,1=tentativo,2=ocupado,3=fuera,4=trabajando fuera
                "busy": [
                    {"start": (i.get("start") or {}).get("dateTime"),
                     "end": (i.get("end") or {}).get("dateTime"),
                     "status": i.get("status")}
                    for i in (s.get("scheduleItems") or [])
                ],
            })
        return {"schedules": out, "legend": "availabilityView: 0=libre 1=provisional 2=ocupado 3=fuera 4=trabajando-fuera"}
