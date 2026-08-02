"""Formateadores: convierten objetos de Graph en dicts compactos.

Objetivo: respuestas pequeñas para no saturar el contexto del agente. El
cuerpo completo se pide aparte con get_message / get_event.
"""

from __future__ import annotations

from typing import Any


def _addr(recipient: dict | None) -> str | None:
    if not recipient:
        return None
    ea = recipient.get("emailAddress", {})
    name, addr = ea.get("name"), ea.get("address")
    if name and addr and name != addr:
        return f"{name} <{addr}>"
    return addr or name


def _addrs(recipients: list | None) -> list[str]:
    return [a for a in (_addr(r) for r in (recipients or [])) if a]


def fmt_message(m: dict, *, full: bool = False) -> dict[str, Any]:
    out = {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "from": _addr(m.get("from")),
        "to": _addrs(m.get("toRecipients")),
        "received": m.get("receivedDateTime"),
        "isRead": m.get("isRead"),
        "hasAttachments": m.get("hasAttachments"),
        "importance": m.get("importance"),
        "flag": (m.get("flag") or {}).get("flagStatus"),
        "categories": m.get("categories") or [],
        "webLink": m.get("webLink"),
    }
    if m.get("conversationId"):
        out["conversationId"] = m["conversationId"]
    if full:
        out["cc"] = _addrs(m.get("ccRecipients"))
        out["bcc"] = _addrs(m.get("bccRecipients"))
        out["replyTo"] = _addrs(m.get("replyTo"))
        out["sent"] = m.get("sentDateTime")
        out["parentFolderId"] = m.get("parentFolderId")
        body = m.get("body", {})
        out["bodyType"] = body.get("contentType")
        out["body"] = body.get("content")
    else:
        out["preview"] = m.get("bodyPreview")
    return out


def fmt_folder(f: dict) -> dict[str, Any]:
    return {
        "id": f.get("id"),
        "name": f.get("displayName"),
        "unread": f.get("unreadItemCount"),
        "total": f.get("totalItemCount"),
        "parentFolderId": f.get("parentFolderId"),
        "childFolderCount": f.get("childFolderCount"),
    }


def fmt_attachment(a: dict, *, with_content: bool = False) -> dict[str, Any]:
    out = {
        "id": a.get("id"),
        "name": a.get("name"),
        "contentType": a.get("contentType"),
        "size": a.get("size"),
        "isInline": a.get("isInline"),
        "type": (a.get("@odata.type") or "").split(".")[-1],
    }
    if with_content and a.get("contentBytes"):
        out["contentBytes"] = a["contentBytes"]  # base64
    return out


def fmt_event(e: dict, *, full: bool = False) -> dict[str, Any]:
    def _dt(x):
        return x.get("dateTime") if isinstance(x, dict) else x

    out = {
        "id": e.get("id"),
        "subject": e.get("subject"),
        "start": _dt(e.get("start")),
        "end": _dt(e.get("end")),
        "timezone": (e.get("start") or {}).get("timeZone"),
        "isAllDay": e.get("isAllDay"),
        "location": (e.get("location") or {}).get("displayName"),
        "organizer": _addr(e.get("organizer")),
        "showAs": e.get("showAs"),
        "isCancelled": e.get("isCancelled"),
        "isOnlineMeeting": e.get("isOnlineMeeting"),
        "onlineMeetingUrl": (e.get("onlineMeeting") or {}).get("joinUrl")
        or e.get("onlineMeetingUrl"),
        "responseStatus": (e.get("responseStatus") or {}).get("response"),
        "seriesMasterId": e.get("seriesMasterId"),
        "type": e.get("type"),
        "webLink": e.get("webLink"),
    }
    attendees = e.get("attendees") or []
    out["attendees"] = [
        {
            "name": _addr(a),
            "type": a.get("type"),
            "response": (a.get("status") or {}).get("response"),
        }
        for a in attendees
    ]
    if full:
        body = e.get("body", {})
        out["bodyType"] = body.get("contentType")
        out["body"] = body.get("content")
        out["categories"] = e.get("categories") or []
        out["importance"] = e.get("importance")
        out["sensitivity"] = e.get("sensitivity")
        out["recurrence"] = e.get("recurrence")
        out["reminderMinutesBeforeStart"] = e.get("reminderMinutesBeforeStart")
        out["hasAttachments"] = e.get("hasAttachments")
    return out


def fmt_calendar(c: dict) -> dict[str, Any]:
    return {
        "id": c.get("id"),
        "name": c.get("name"),
        "color": c.get("color"),
        "isDefault": c.get("isDefaultCalendar"),
        "canEdit": c.get("canEdit"),
        "owner": _addr(c.get("owner")) if c.get("owner") else None,
    }
