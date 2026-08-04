#!/usr/bin/env python3
"""
Registra en Claude Desktop los conectores locales de Outlook.

Por defecto (paquete completo de cliente):
  - 'outlook'        -> python -m outlook_mcp stdio   (conector COMPLETO)
  - 'agente-outlook' -> agent.py                      (puente con el disco)

Con la bandera --solo-agente (o --agent-only): registra SOLO 'agente-outlook'
y no exige el paquete del conector. Útil para instalar únicamente el puente.

El conector y el agente se ejecutan DESDE EL ÁRBOL DE ORIGEN de esta carpeta
(vía PYTHONPATH), sin instalación editable: así el paquete funciona offline sin
necesitar backend de compilación.

IMPORTANTE: ejecutar con Claude COMPLETAMENTE CERRADO (si esta abierto, al
cerrarse sobrescribe el cambio). Se niega si Claude esta abierto (--force lo salta).
"""
import json, os, sys, shutil, time, platform, subprocess

HERE = os.path.dirname(os.path.realpath(__file__))
ENVP = os.path.join(HERE, ".env")
SOLO = ("--solo-agente" in sys.argv) or ("--agent-only" in sys.argv)


def config_path():
    home = os.path.expanduser("~")
    s = platform.system()
    if s == "Darwin":
        return os.path.join(home, "Library/Application Support/Claude/claude_desktop_config.json")
    if s == "Windows":
        return os.path.join(os.environ.get("APPDATA", home), "Claude", "claude_desktop_config.json")
    return os.path.join(home, ".config/Claude/claude_desktop_config.json")


def claude_procs():
    procs = []
    try:
        sysname = platform.system()
        if sysname == "Windows":
            out = subprocess.run(["tasklist"], capture_output=True, text=True).stdout
            return ["Claude.exe"] if "Claude.exe" in out else []
        out = subprocess.run(["ps", "-Ao", "command="], capture_output=True, text=True).stdout
        for line in out.splitlines():
            l = line.strip()
            if sysname == "Darwin":
                if "/Applications/Claude.app" not in l:
                    continue
            else:
                if "claude" not in l.lower() or ".app" not in l.lower():
                    continue
            if "chrome-native-host" in l or "crashpad" in l or "conectar_claude" in l:
                continue
            procs.append(l[:70])
    except Exception:
        pass
    return procs


def load_env(p):
    if not os.path.exists(p):
        sys.exit("ERROR: no encuentro el archivo .env en " + p)
    ev = {}
    for line in open(p, encoding="utf-8"):
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        ev[k.strip()] = v.strip().strip('"').strip("'")
    return ev


def main():
    procs = claude_procs()
    if procs and "--force" not in sys.argv:
        muestra = "\n".join("   - " + x for x in procs[:6])
        sys.exit("ATENCION: Claude (app de escritorio) sigue ABIERTO. Cierralo del todo\n"
                 "con Cmd+Q (o cerrar del todo en Windows), espera 3 segundos y repite.\n"
                 "Procesos detectados:\n" + muestra)

    ev = load_env(ENVP)
    faltan = [k for k in ("OUTLOOK_CLIENT_ID", "OUTLOOK_TENANT_ID") if not ev.get(k)]
    if faltan:
        sys.exit("ERROR: faltan datos en el .env: " + ", ".join(faltan) +
                 "\nRellena OUTLOOK_CLIENT_ID y OUTLOOK_TENANT_ID (ver GUIA DETALLADA / README).")

    if platform.system() == "Windows":
        py = os.path.join(HERE, ".venv/Scripts/python.exe")
    else:
        py = os.path.join(HERE, ".venv/bin/python")
    agent = os.path.join(HERE, "agent.py")
    requeridos = [py, agent]
    if not SOLO:
        # el conector completo se ejecuta desde el árbol de origen (outlook_mcp/)
        requeridos.append(os.path.join(HERE, "outlook_mcp", "__init__.py"))
    for f in requeridos:
        if not os.path.exists(f):
            sys.exit("ERROR: falta " + f + " (ejecuta antes el instalador).")

    raw = ev.get("OUTLOOK_ALLOWED_DIRS", "~") or "~"
    sep = ";" if platform.system() == "Windows" else ":"
    partes = [os.path.realpath(os.path.expanduser(os.path.expandvars(x)))
              for x in raw.replace(";", ":").split(":") if x.strip()]
    allowed = sep.join(partes) if partes else os.path.expanduser("~")

    tz = ev.get("OUTLOOK_TIMEZONE", "Europe/Madrid")
    # PYTHONPATH = esta carpeta -> permite 'python -m outlook_mcp' y 'import outlook_mcp'
    # sin instalación editable (imprescindible para el paquete offline).
    base_env = {
        "OUTLOOK_CLIENT_ID": ev["OUTLOOK_CLIENT_ID"],
        "OUTLOOK_TENANT_ID": ev["OUTLOOK_TENANT_ID"],
        "OUTLOOK_TIMEZONE": tz,
        "PYTHONPATH": HERE,
    }
    if ev.get("OUTLOOK_TOKEN_CACHE"):
        base_env["OUTLOOK_TOKEN_CACHE"] = os.path.expanduser(ev["OUTLOOK_TOKEN_CACHE"])

    cfg = config_path()
    os.makedirs(os.path.dirname(cfg), exist_ok=True)
    if os.path.exists(cfg):
        shutil.copy2(cfg, cfg + ".bak-" + time.strftime("%Y%m%d-%H%M%S"))
        d = json.load(open(cfg, encoding="utf-8"))
    else:
        d = {}
    d.setdefault("mcpServers", {})

    # Conector COMPLETO — solo si NO estamos en modo --solo-agente
    if not SOLO:
        d["mcpServers"]["outlook"] = {
            "command": py, "args": ["-m", "outlook_mcp", "stdio"],
            "env": dict(base_env, MCP_TRANSPORT="stdio"),
        }
    # Agente local (puente con el disco) — siempre
    d["mcpServers"]["agente-outlook"] = {
        "command": py, "args": [agent],
        "env": dict(base_env, MCP_TRANSPORT="stdio", OUTLOOK_ALLOWED_DIRS=allowed),
    }
    json.dump(d, open(cfg, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    if SOLO:
        print("OK. Registrado SOLO 'agente-outlook' (puente con tu disco).")
    else:
        print("OK. Registrados 'outlook' (completo) y 'agente-outlook' (puente con tu disco).")
    print("  Config:     " + cfg)
    print("  Carpetas:   " + allowed)
    print("  Conectores: " + ", ".join(d["mcpServers"].keys()))


if __name__ == "__main__":
    main()
