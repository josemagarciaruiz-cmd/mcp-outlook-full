"""Herramientas de Microsoft Teams (reuniones online, asistencia, presencia y
recuperación de grabaciones/transcripciones).

Scopes: OnlineMeetings.ReadWrite (reuniones + asistencia), Presence.Read (presencia).
Grabaciones/transcripciones: en un conector DELEGADO (device code) no se pueden bajar
por la API de artefactos de Teams (esa exige permiso de APLICACIÓN + admin). La vía
real es cogerlas como FICHERO en OneDrive/SharePoint, donde Teams las guarda: para eso
sirven teams_find_recording + teams_transcript_text (y las herramientas files_*/
sharepoint_*). Reuniones de canal → SharePoint; reuniones normales → OneDrive.
"""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from . import graph


def register(mcp) -> None:
    @mcp.tool(annotations={"openWorldHint": True})
    async def teams_create_meeting(
        subject: Annotated[str, Field(description="Título de la reunión")],
        start: Annotated[str, Field(description="Inicio ISO 8601 con zona, p.ej. 2026-08-05T10:00:00Z")],
        end: Annotated[str, Field(description="Fin ISO 8601 con zona")],
    ) -> dict:
        """Crea una reunión de Teams (sin evento de calendario) y devuelve el enlace de unión.

        Para crear la reunión Y ponerla en tu agenda con invitados, usa
        calendar_create_event(is_online_meeting=True). Esta es la vía directa a un enlace.
        """
        res = await graph.request(
            "POST", "/me/onlineMeetings",
            json={"subject": subject, "startDateTime": start, "endDateTime": end},
        )
        return {
            "id": res.get("id"),
            "joinUrl": res.get("joinWebUrl") or res.get("joinUrl"),
            "subject": res.get("subject"),
            "start": res.get("startDateTime"),
            "end": res.get("endDateTime"),
        }

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_get_meeting(
        meeting_id: Annotated[str, Field(description="Id de la reunión (onlineMeeting)")],
    ) -> dict:
        """Detalles de una reunión de Teams (enlace de unión, ajustes)."""
        res = await graph.request("GET", f"/me/onlineMeetings/{meeting_id}")
        return {
            "id": res.get("id"),
            "joinUrl": res.get("joinWebUrl") or res.get("joinUrl"),
            "subject": res.get("subject"),
            "start": res.get("startDateTime"),
            "end": res.get("endDateTime"),
            "recordAutomatically": res.get("recordAutomatically"),
            "allowedPresenters": res.get("allowedPresenters"),
        }

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_attendance_reports(
        meeting_id: Annotated[str, Field(description="Id de la reunión (onlineMeeting)")],
    ) -> dict:
        """Lista los informes de asistencia de una reunión (uno por sesión celebrada)."""
        res = await graph.request("GET", f"/me/onlineMeetings/{meeting_id}/attendanceReports")
        reports = res.get("value", []) if isinstance(res, dict) else []
        return {"reports": [
            {"id": r.get("id"), "totalParticipantCount": r.get("totalParticipantCount"),
             "start": r.get("meetingStartDateTime"), "end": r.get("meetingEndDateTime")}
            for r in reports
        ]}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_attendance_detail(
        meeting_id: Annotated[str, Field(description="Id de la reunión")],
        report_id: Annotated[str, Field(description="Id del informe (de teams_attendance_reports)")],
    ) -> dict:
        """Detalle de asistencia: quién asistió, su rol y los MINUTOS que estuvo (para facturar/constancia)."""
        res = await graph.request(
            "GET", f"/me/onlineMeetings/{meeting_id}/attendanceReports/{report_id}",
            params={"$expand": "attendanceRecords"},
        )
        recs = res.get("attendanceRecords", []) if isinstance(res, dict) else []
        out = []
        for r in recs:
            ident = (r.get("identity") or {})
            out.append({
                "name": ident.get("displayName") or r.get("emailAddress"),
                "email": r.get("emailAddress"),
                "role": r.get("role"),
                "minutos": round((r.get("totalAttendanceInSeconds") or 0) / 60, 1),
            })
        return {"total": res.get("totalParticipantCount"), "asistentes": out}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_presence(
        user: Annotated[Optional[str], Field(None, description="Email/id del usuario; vacío = tú")] = None,
    ) -> dict:
        """Presencia en Teams: disponible / ocupado / en una llamada / ausente."""
        path = f"/users/{user}/presence" if user else "/me/presence"
        res = await graph.request("GET", path)
        return {"availability": res.get("availability"), "activity": res.get("activity")}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_find_recording(
        query: Annotated[str, Field(description="Texto para localizar (p.ej. el título de la reunión o del cliente)")],
        top: Annotated[int, Field(25, ge=1, le=100)] = 25,
    ) -> dict:
        """Localiza en TU OneDrive la grabación (.mp4) o la transcripción (.vtt/.docx) de una reunión.

        Teams guarda las grabaciones de reuniones normales en tu OneDrive (carpeta
        'Recordings'). Para reuniones de CANAL, están en SharePoint: usa sharepoint_search.
        """
        q = query.replace("'", "''")
        res = await graph.request("GET", f"/me/drive/root/search(q='{q}')", params={"$top": top})
        items = res.get("value", []) if isinstance(res, dict) else []
        media = []
        for i in items:
            name = (i.get("name") or "").lower()
            if name.endswith((".mp4", ".vtt", ".docx")) or "recording" in name or "transcript" in name:
                media.append({
                    "id": i.get("id"), "name": i.get("name"), "size": i.get("size"),
                    "webUrl": i.get("webUrl"),
                    "downloadUrl": i.get("@microsoft.graph.downloadUrl"),
                    "tipo": "grabación" if name.endswith(".mp4") else "transcripción/otro",
                })
        return {"encontrados": media, "aviso": "Si es reunión de canal, mira en SharePoint (sharepoint_search)."}

    @mcp.tool(annotations={"readOnlyHint": True, "openWorldHint": True})
    async def teams_transcript_text(
        item_id: Annotated[str, Field(description="Id del fichero de transcripción (.vtt) en el drive")],
        drive_id: Annotated[Optional[str], Field(None, description="Id del drive de SharePoint; vacío = tu OneDrive")] = None,
    ) -> dict:
        """Devuelve el TEXTO de una transcripción (.vtt) para que la IA redacte el acta.

        Funciona con ficheros .vtt (texto). Para .docx conviene convertir antes
        (files_save_pdf o el conector markitdown).
        """
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        data = await graph.get_content(f"{base}/items/{item_id}/content")
        try:
            text = data.decode("utf-8", errors="replace")
        except Exception:
            text = ""
        # Limpieza ligera de WebVTT: quitar cabecera y marcas de tiempo
        lines = []
        for ln in text.splitlines():
            s = ln.strip()
            if not s or s == "WEBVTT" or "-->" in s or s.isdigit():
                continue
            lines.append(s)
        limpio = "\n".join(lines)
        return {"caracteres": len(limpio), "texto": limpio[:60000],
                "truncado": len(limpio) > 60000}
