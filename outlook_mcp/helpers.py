"""Constructores de payloads de Graph y guardas compartidas."""

from __future__ import annotations

import base64
import mimetypes
from pathlib import Path

from . import config

# Adjuntos por debajo de este tamaño van embebidos en el propio mensaje
# (1 sola llamada). Por encima, hay que subirlos por upload session a un
# borrador. Graph limita la vía embebida a ~3 MB.
LARGE_ATTACHMENT_THRESHOLD = 3_000_000  # bytes


def _source_bytes(a: dict) -> tuple[str, str, bytes] | None:
    """Obtiene (name, contentType, bytes) de un adjunto dado por RUTA o base64.

    - {"path": "/ruta/fichero.pdf"}  -> el SERVIDOR lee el fichero (el base64 NO
      pasa por el modelo; así no se corrompe). name/contentType se infieren.
    - {"contentBytes": "<base64>", ...} -> bytes ya provistos (p.ej. de otra tool).
    """
    path = a.get("path")
    if path:
        p = Path(path).expanduser()
        if not p.is_file():
            raise FileNotFoundError(
                f"El servidor no encuentra el fichero '{path}'. Si el conector corre en "
                f"el VPS no ve el disco de tu ordenador: usa el conector LOCAL para adjuntar "
                f"ficheros del equipo, o pasa el contenido en 'contentBytes'."
            )
        raw = p.read_bytes()
        name = a.get("name") or p.name
        ctype = a.get("contentType") or mimetypes.guess_type(p.name)[0] or "application/octet-stream"
        return name, ctype, raw
    cb = a.get("contentBytes")
    if cb:
        raw = base64.b64decode(cb)
        return a.get("name", "adjunto"), a.get("contentType", "application/octet-stream"), raw
    return None


def resolve_local(specs: list[dict] | None) -> list[tuple[str, str, bytes]]:
    """Resuelve adjuntos por ruta/base64 (síncrono) a [(name, contentType, bytes)]."""
    out: list[tuple[str, str, bytes]] = []
    for a in specs or []:
        src = _source_bytes(a)
        if src:
            out.append(src)
    return out


def classify(resolved: list[tuple[str, str, bytes]]) -> tuple[list[dict], list[dict]]:
    """Clasifica [(name, ctype, bytes)] en (pequeños embebibles, grandes por upload).

    - pequeños: {name, contentType, contentBytes(base64)} listos para incrustar.
    - grandes: {name, contentType, data(bytes)} para subir por trozos.
    """
    small: list[dict] = []
    large: list[dict] = []
    for name, ctype, raw in resolved:
        if len(raw) > LARGE_ATTACHMENT_THRESHOLD:
            large.append({"name": name, "contentType": ctype, "data": raw})
        else:
            small.append({
                "name": name,
                "contentType": ctype,
                "contentBytes": base64.b64encode(raw).decode("ascii"),
            })
    return small, large


def split_attachments(attachments: list[dict] | None) -> tuple[list[dict], list[dict]]:
    """Compat: resuelve adjuntos por ruta/base64 y los clasifica (small, large)."""
    return classify(resolve_local(attachments))


def mbox(mailbox: str | None) -> str:
    """Base de la API: buzon propio (/me) o uno compartido (/users/{email})."""
    return f"/users/{mailbox}" if mailbox else "/me"


def guard_write() -> None:
    if config.READ_ONLY:
        raise RuntimeError(
            "Servidor en modo solo lectura (OUTLOOK_READ_ONLY). "
            "Operación de escritura/envío/borrado bloqueada."
        )


def recipients(addresses: list[str] | None) -> list[dict]:
    """['a@x.com', 'Nombre <b@y.com>'] -> lista de recipients de Graph."""
    out = []
    for a in addresses or []:
        a = a.strip()
        if not a:
            continue
        if "<" in a and ">" in a:
            name = a.split("<")[0].strip().strip('"')
            addr = a.split("<")[1].split(">")[0].strip()
            out.append({"emailAddress": {"address": addr, "name": name or addr}})
        else:
            out.append({"emailAddress": {"address": a}})
    return out


def body(content: str | None, content_type: str = "HTML") -> dict | None:
    if content is None:
        return None
    ct = "HTML" if content_type.lower() == "html" else "Text"
    return {"contentType": ct, "content": content}


def date_time_zone(dt: str, tz: str | None = None) -> dict:
    """ISO 8601 sin zona -> objeto dateTimeTimeZone de Graph."""
    return {"dateTime": dt, "timeZone": tz or config.TIMEZONE}
