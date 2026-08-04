"""Cliente asíncrono de Microsoft Graph con manejo de errores y paginación.

Todas las tools de correo y calendario pasan por aquí. Inyecta el bearer de
MSAL, normaliza errores de Graph a mensajes accionables y ofrece helpers de
paginación (@odata.nextLink) y de subida de adjuntos grandes.
"""

from __future__ import annotations

import base64
from typing import Any

import httpx

from . import auth, config

_client: httpx.AsyncClient | None = None


class GraphError(RuntimeError):
    """Error de Graph con código y mensaje legibles para el agente."""


def _get_client() -> httpx.AsyncClient:
    global _client
    if _client is None or _client.is_closed:
        _client = httpx.AsyncClient(base_url=config.GRAPH_BASE, timeout=60.0)
    return _client


def _auth_headers(extra: dict | None = None) -> dict:
    headers = {"Authorization": f"Bearer {auth.get_access_token()}"}
    if extra:
        headers.update(extra)
    return headers


def _raise_for_graph(resp: httpx.Response) -> None:
    if resp.is_success:
        return
    detail = ""
    try:
        body = resp.json()
        err = body.get("error", {}) if isinstance(body, dict) else {}
        code = err.get("code", "")
        msg = err.get("message", "")
        detail = f"{code}: {msg}".strip(": ")
    except Exception:
        detail = resp.text[:500]
    hint = ""
    if resp.status_code == 401:
        hint = " (token inválido/expirado; reintenta o vuelve a hacer login)"
    elif resp.status_code == 403:
        hint = " (permiso insuficiente; revisa los scopes/consentimiento del registro de Azure)"
    elif resp.status_code == 429:
        ra = resp.headers.get("Retry-After")
        hint = f" (throttling de Graph; reintenta tras {ra or 'unos'} segundos)"
    raise GraphError(f"Graph {resp.status_code} {detail}{hint}")


async def request(
    method: str,
    path: str,
    *,
    params: dict | None = None,
    json: Any | None = None,
    headers: dict | None = None,
) -> Any:
    """Petición genérica a Graph. Devuelve JSON (dict/list) o None si 204/202."""
    client = _get_client()
    resp = await client.request(
        method,
        path,
        params=params,
        json=json,
        headers=_auth_headers(headers),
    )
    _raise_for_graph(resp)
    if resp.status_code in (202, 204) or not resp.content:
        return None
    ctype = resp.headers.get("content-type", "")
    if "application/json" in ctype:
        return resp.json()
    return resp.content


async def get_paged(
    path: str,
    *,
    params: dict | None = None,
    headers: dict | None = None,
    max_items: int = 50,
) -> dict:
    """Recupera hasta max_items siguiendo @odata.nextLink.

    Devuelve {"items": [...], "next": <url|None>} para que el agente pueda
    seguir paginando si quiere (pasando la url a `get_next`).
    """
    items: list = []
    data = await request("GET", path, params=params, headers=headers)
    while True:
        page = data.get("value", []) if isinstance(data, dict) else []
        for it in page:
            items.append(it)
            if len(items) >= max_items:
                return {"items": items, "next": data.get("@odata.nextLink")}
        nxt = data.get("@odata.nextLink") if isinstance(data, dict) else None
        if not nxt:
            return {"items": items, "next": None}
        # nextLink es una URL absoluta: la usamos tal cual.
        client = _get_client()
        resp = await client.get(nxt, headers=_auth_headers(headers))
        _raise_for_graph(resp)
        data = resp.json()


async def get_next(next_link: str, *, headers: dict | None = None, max_items: int = 50) -> dict:
    """Continúa una paginación a partir de un @odata.nextLink absoluto."""
    client = _get_client()
    resp = await client.get(next_link, headers=_auth_headers(headers))
    _raise_for_graph(resp)
    data = resp.json()
    items = data.get("value", [])[:max_items]
    return {"items": items, "next": data.get("@odata.nextLink")}


async def _put_chunks(upload_url: str, data: bytes) -> None:
    """Sube bytes por rangos a un uploadUrl de sesión (OneDrive/adjuntos)."""
    client = _get_client()
    size = len(data)
    start = 0
    while start < size:
        end = min(start + _UPLOAD_CHUNK, size)
        chunk = data[start:end]
        headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {start}-{end - 1}/{size}",
        }
        resp = await client.put(upload_url, content=chunk, headers=headers)
        if resp.status_code not in (200, 201, 202):
            _raise_for_graph(resp)
        start = end


async def onedrive_upload(dest_path: str, data: bytes, content_type: str = "application/octet-stream") -> dict:
    """Sube bytes a OneDrive en dest_path (p.ej. 'Documentos/archivo.pdf')."""
    client = _get_client()
    dest_path = dest_path.lstrip("/")
    if len(data) <= 4 * 1024 * 1024:
        resp = await client.put(
            f"/me/drive/root:/{dest_path}:/content",
            content=data,
            headers=_auth_headers({"Content-Type": content_type}),
        )
        _raise_for_graph(resp)
        return resp.json()
    session = await request(
        "POST", f"/me/drive/root:/{dest_path}:/createUploadSession",
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
    )
    await _put_chunks(session["uploadUrl"], data)
    return await request("GET", f"/me/drive/root:/{dest_path}")


async def download_item(item_id: str) -> bytes:
    """Descarga el contenido de un item de OneDrive (sigue la redirección)."""
    return await get_content(f"/me/drive/items/{item_id}/content")


async def get_json(url: str, params: dict | None = None) -> dict:
    """GET que devuelve JSON, admitiendo ruta relativa o URL absoluta (p.ej. un
    deltaLink/nextLink de Graph). Conserva claves @odata.* (deltaLink, nextLink)."""
    client = _get_client()
    resp = await client.get(url, headers=_auth_headers(), params=params, follow_redirects=True)
    _raise_for_graph(resp)
    return resp.json()


async def get_content(path: str) -> bytes:
    """GET binario de un endpoint /content (OneDrive o SharePoint), siguiendo redirecciones.

    Sirve para descargar cualquier drive (/me/drive, /drives/{id}, /sites/.../drive)
    y para conversión (p.ej. .../content?format=pdf).
    """
    client = _get_client()
    resp = await client.get(path, headers=_auth_headers(), follow_redirects=True)
    _raise_for_graph(resp)
    return resp.content


def prefer_headers(text_body: bool = False, timezone: str | None = None) -> dict:
    """Cabeceras Prefer para pedir cuerpo en texto plano y zona horaria."""
    prefs = []
    if text_body:
        prefs.append('outlook.body-content-type="text"')
    tz = timezone or config.TIMEZONE
    if tz:
        prefs.append(f'outlook.timezone="{tz}"')
    return {"Prefer": ", ".join(prefs)} if prefs else {}


def b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


# Tamaño de trozo para upload sessions: múltiplo de 320 KiB (requisito de Graph)
# y por debajo de 4 MB para no reventar el body. 10 * 320 KiB = 3,125 MiB.
_UPLOAD_CHUNK = 10 * 320 * 1024


async def fetch_attachment(url: str, *, max_bytes: int = 150 * 1024 * 1024) -> tuple[str, str, bytes]:
    """Descarga un fichero desde una URL http(s) para adjuntarlo.

    Devuelve (name, contentType, bytes). NO envía la cabecera Authorization de
    Graph (el token no debe salir a hosts externos). Sigue redirecciones.
    """
    import mimetypes
    import urllib.parse

    if not url.lower().startswith(("http://", "https://")):
        raise GraphError(f"URL de adjunto no válida (usa http/https): {url}")
    client = _get_client()
    resp = await client.get(url, follow_redirects=True, timeout=120.0)
    if not resp.is_success:
        raise GraphError(f"No se pudo descargar el adjunto de {url}: HTTP {resp.status_code}")
    data = resp.content
    if len(data) > max_bytes:
        raise GraphError(f"El fichero de {url} supera el máximo permitido (~150 MB).")
    name = None
    cd = resp.headers.get("content-disposition", "")
    if "filename=" in cd:
        name = cd.split("filename=")[-1].split(";")[0].strip().strip('"')
    if not name:
        path = urllib.parse.urlparse(url).path
        name = urllib.parse.unquote(path.rsplit("/", 1)[-1]) or "adjunto"
    ctype = (resp.headers.get("content-type", "").split(";")[0].strip()
             or mimetypes.guess_type(name)[0] or "application/octet-stream")
    return name, ctype, data


async def add_attachments(message_id: str, small: list[dict], large: list[dict], base: str = "/me") -> None:
    """Añade adjuntos ya clasificados a un mensaje/borrador existente."""
    for s in small:
        await request(
            "POST",
            f"{base}/messages/{message_id}/attachments",
            json={"@odata.type": "#microsoft.graph.fileAttachment", **s},
        )
    for l in large:
        await upload_large_attachment(message_id, l["name"], l["contentType"], l["data"], base=base)


async def upload_large_attachment(
    message_id: str, name: str, content_type: str, data: bytes, base: str = "/me"
) -> None:
    """Sube un adjunto grande a un borrador mediante upload session (por trozos).

    Graph no admite adjuntos >~3 MB embebidos en el mensaje; hay que crear una
    sesión de subida y enviar los bytes en rangos (Content-Range). El uploadUrl
    ya viene pre-autenticado, así que se hace PUT sin cabecera Authorization.
    """
    size = len(data)
    session = await request(
        "POST",
        f"{base}/messages/{message_id}/attachments/createUploadSession",
        json={
            "AttachmentItem": {
                "attachmentType": "file",
                "name": name,
                "size": size,
                "contentType": content_type,
            }
        },
    )
    upload_url = session["uploadUrl"]
    client = _get_client()
    start = 0
    while start < size:
        end = min(start + _UPLOAD_CHUNK, size)
        chunk = data[start:end]
        headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {start}-{end - 1}/{size}",
        }
        resp = await client.put(upload_url, content=chunk, headers=headers)
        if resp.status_code not in (200, 201, 202):
            _raise_for_graph(resp)
        start = end


async def aclose() -> None:
    global _client
    if _client is not None and not _client.is_closed:
        await _client.aclose()
