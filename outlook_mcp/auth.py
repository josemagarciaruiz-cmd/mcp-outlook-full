"""Autenticación con Microsoft Entra ID (MSAL, flujo de código de dispositivo).

Modelo: permisos DELEGADOS + device code flow. Se hace login UNA vez por
máquina (cada Mac; una vez en el VPS). El token, con su refresh token
(offline_access), queda cacheado y se renueva solo de forma silenciosa.
"""

from __future__ import annotations

import atexit
import sys
import threading

import msal

from . import config

_lock = threading.Lock()
_cache: msal.SerializableTokenCache | None = None
_app: msal.PublicClientApplication | None = None


def _load_cache() -> msal.SerializableTokenCache:
    global _cache
    if _cache is not None:
        return _cache
    cache = msal.SerializableTokenCache()
    path = config.TOKEN_CACHE_PATH
    if path.exists():
        try:
            cache.deserialize(path.read_text())
        except Exception:  # cache corrupto -> se regenera con el próximo login
            pass

    def _persist() -> None:
        if cache.has_state_changed:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(cache.serialize())
            try:
                path.chmod(0o600)
            except OSError:
                pass

    atexit.register(_persist)
    _cache = cache
    return cache


def _persist_now() -> None:
    if _cache is not None and _cache.has_state_changed:
        path = config.TOKEN_CACHE_PATH
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(_cache.serialize())
        try:
            path.chmod(0o600)
        except OSError:
            pass


def _get_app() -> msal.PublicClientApplication:
    global _app
    if _app is not None:
        return _app
    client_id = config.require_client_id()
    _app = msal.PublicClientApplication(
        client_id,
        authority=config.AUTHORITY,
        token_cache=_load_cache(),
    )
    return _app


def device_login() -> dict:
    """Ejecuta el flujo de código de dispositivo de forma interactiva.

    Imprime el código y la URL por stderr para que el usuario los introduzca en
    el navegador. Bloquea hasta que se completa. Devuelve la info de la cuenta.
    Pensado para ejecutarse desde la CLI: `python -m outlook_mcp login`.
    """
    app = _get_app()
    flow = app.initiate_device_flow(scopes=config.SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            f"No se pudo iniciar el device flow: {flow.get('error_description', flow)}"
        )
    # message ya trae un texto listo del estilo: "To sign in, use a web browser
    # to open the page https://microsoft.com/devicelogin and enter the code XXXX".
    print("\n" + flow["message"] + "\n", file=sys.stderr, flush=True)
    result = app.acquire_token_by_device_flow(flow)  # bloquea hasta completar
    _persist_now()
    if "access_token" not in result:
        raise RuntimeError(
            f"Login fallido: {result.get('error')}: {result.get('error_description')}"
        )
    claims = result.get("id_token_claims", {})
    return {
        "account": claims.get("preferred_username") or claims.get("email"),
        "name": claims.get("name"),
        "tenant": claims.get("tid"),
        "scopes": result.get("scope"),
    }


def get_access_token() -> str:
    """Devuelve un access token válido de Graph, renovándolo en silencio.

    No es interactivo: si no hay cuenta cacheada lanza un error accionable
    indicando que hay que hacer login primero.
    """
    with _lock:
        app = _get_app()
        accounts = app.get_accounts()
        if not accounts:
            raise RuntimeError(
                "No hay ninguna cuenta autenticada. Ejecuta primero el login:\n"
                "  python -m outlook_mcp login\n"
                "(en el VPS, hazlo una vez dentro del contenedor o con el volumen "
                "de token montado)."
            )
        result = app.acquire_token_silent(config.SCOPES, account=accounts[0])
        _persist_now()
        if not result or "access_token" not in result:
            err = (result or {}).get("error_description", "sesión expirada")
            raise RuntimeError(
                f"No se pudo renovar el token en silencio ({err}). "
                "Vuelve a ejecutar: python -m outlook_mcp login"
            )
        return result["access_token"]


def current_account() -> dict | None:
    app = _get_app()
    accounts = app.get_accounts()
    if not accounts:
        return None
    a = accounts[0]
    return {"account": a.get("username"), "home_account_id": a.get("home_account_id")}


def logout() -> bool:
    app = _get_app()
    accounts = app.get_accounts()
    for a in accounts:
        app.remove_account(a)
    _persist_now()
    return bool(accounts)
