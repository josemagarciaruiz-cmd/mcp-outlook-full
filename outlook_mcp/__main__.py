"""CLI del servidor MCP de Outlook.

Uso:
  python -m outlook_mcp login     # login interactivo (device code), una vez por máquina
  python -m outlook_mcp whoami    # verifica la cuenta y el acceso a Graph
  python -m outlook_mcp logout
  python -m outlook_mcp stdio     # transporte stdio (local, Claude Desktop) [por defecto]
  python -m outlook_mcp http      # transporte Streamable HTTP (VPS)
"""

from __future__ import annotations

import argparse
import asyncio
import sys

from . import auth, config


def _cmd_login() -> int:
    config.require_client_id()
    info = auth.device_login()
    print(f"\n✓ Sesión iniciada como {info.get('name')} <{info.get('account')}> "
          f"(tenant {info.get('tenant')})", file=sys.stderr)
    return 0


def _cmd_logout() -> int:
    removed = auth.logout()
    print("✓ Sesión cerrada." if removed else "No había sesión activa.", file=sys.stderr)
    return 0


def _cmd_whoami() -> int:
    from . import graph
    from .server import build_server  # asegura que la app existe

    build_server()

    async def _run():
        try:
            me = await graph.request("GET", "/me", params={"$select": "displayName,userPrincipalName,mail"})
            print(f"✓ {me.get('displayName')} <{me.get('mail') or me.get('userPrincipalName')}>", file=sys.stderr)
            return 0
        except Exception as e:  # noqa: BLE001
            print(f"✗ {e}", file=sys.stderr)
            return 1
        finally:
            await graph.aclose()

    return asyncio.run(_run())


def _cmd_stdio() -> int:
    from .server import build_server

    build_server().run(transport="stdio")
    return 0


def _cmd_http() -> int:
    import uvicorn
    from mcp.server.transport_security import TransportSecuritySettings

    from .server import build_server

    mcp = build_server()
    # Detrás de Traefik el Host es tu dominio; la protección anti DNS-rebinding
    # (activa por defecto) lo bloquearía. El acceso ya está protegido por la
    # ruta-token secreta de Traefik y por el Bearer, así que la desactivamos.
    app = mcp.streamable_http_app(
        host=config.HTTP_HOST,
        transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
    )

    if config.HTTP_BEARER_TOKEN:
        app = _wrap_bearer(app, config.HTTP_BEARER_TOKEN)

    print(
        f"Outlook MCP (Streamable HTTP) en http://{config.HTTP_HOST}:{config.HTTP_PORT}/mcp"
        + ("  [Bearer requerido]" if config.HTTP_BEARER_TOKEN else ""),
        file=sys.stderr,
    )
    uvicorn.run(app, host=config.HTTP_HOST, port=config.HTTP_PORT)
    return 0


def _wrap_bearer(app, token: str):
    """Middleware ASGI mínimo: exige Authorization: Bearer <token>.

    Segunda línea de defensa además de la ruta-token secreta de Traefik.
    /health queda abierto para healthchecks.
    """
    expected = f"Bearer {token}"

    async def middleware(scope, receive, send):
        if scope["type"] == "http" and scope.get("path") != "/health":
            headers = dict(scope.get("headers") or [])
            if headers.get(b"authorization", b"").decode() != expected:
                await send({"type": "http.response.start", "status": 401,
                            "headers": [(b"content-type", b"text/plain")]})
                await send({"type": "http.response.body", "body": b"Unauthorized"})
                return
        await app(scope, receive, send)

    return middleware


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="outlook_mcp", description=__doc__)
    parser.add_argument(
        "command",
        nargs="?",
        default="stdio",
        choices=["login", "logout", "whoami", "stdio", "http"],
    )
    args = parser.parse_args(argv)
    return {
        "login": _cmd_login,
        "logout": _cmd_logout,
        "whoami": _cmd_whoami,
        "stdio": _cmd_stdio,
        "http": _cmd_http,
    }[args.command]()


if __name__ == "__main__":
    raise SystemExit(main())
