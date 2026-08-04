# ─────────────────────────────────────────────────────────────────────────
# Instalador COMPLETO de MCP Outlook (conector + agente) para Windows.
# Ejecutar con Claude CERRADO (el .bat lo cierra por ti).
#
# Control de errores explícito (NO usamos $ErrorActionPreference='Stop', que
# convertía avisos inofensivos —p. ej. taskkill sin Claude abierto— en abortos).
# ─────────────────────────────────────────────────────────────────────────
Set-Location -Path $PSScriptRoot
$log = Join-Path $PSScriptRoot "install_log.txt"
try { Start-Transcript -Path $log -Force | Out-Null } catch {}

function Fail($msg) {
    Write-Host ""
    Write-Host "ERROR: $msg" -ForegroundColor Red
    Write-Host "Detalle en install_log.txt (en esta misma carpeta)."
    try { Stop-Transcript | Out-Null } catch {}
    Read-Host "Pulsa ENTER para salir"
    exit 1
}

# --- Defecto A: elegir intérprete concreto con el lanzador 'py' (3.13 -> 3.12 -> 3.11)
Write-Host "==> 1/4  Buscando Python 3.13 / 3.12 / 3.11..."
$PYARG = $null
foreach ($v in @("3.13", "3.12", "3.11")) {
    & py "-$v" -c "import sys" 2>$null
    if ($LASTEXITCODE -eq 0) { $PYARG = "-$v"; break }
}
if (-not $PYARG) {
    Fail "No hay Python 3.11, 3.12 o 3.13. Instala 3.13 desde https://www.python.org/downloads/ marcando 'Add python.exe to PATH' y repite."
}
Write-Host "     Usando Python $PYARG"

& py $PYARG -m venv .venv
if ($LASTEXITCODE -ne 0) { Fail "No se pudo crear el entorno virtual (.venv)." }
$py = ".\.venv\Scripts\python.exe"

Write-Host "==> 2/4  Instalando dependencias..."
& $py -m pip install -q --upgrade pip 2>$null
if (Test-Path ".\wheels") {
    Write-Host "     (sin internet: usando la carpeta wheels)"
    & $py -m pip install -q --no-index --find-links wheels -r requirements.txt
    if ($LASTEXITCODE -ne 0) { Fail "Fallo instalando dependencias desde 'wheels' (offline). ¿Coinciden con tu versión de Python?" }
} else {
    Write-Host "     (por internet)"
    & $py -m pip install -q -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "     Reintento con hosts de confianza (proxy/antivirus)..."
        & $py -m pip install -q --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
        if ($LASTEXITCODE -ne 0) { Fail "No se pudieron instalar las dependencias." }
    }
}

# El conector se ejecuta desde el árbol de origen (sin instalación editable).
$env:PYTHONPATH = $PSScriptRoot

Write-Host "==> 3/4  Comprobando que todo carga (conector + agente)..."
& $py -c "import mcp, msal, httpx; from mcp.server.mcpserver import MCPServer; from outlook_mcp.server import build_server; import agent; print('OK imports')"
if ($LASTEXITCODE -ne 0) { Fail "Las librerias o el codigo no cargan (instalacion incompleta)." }

if (-not (Test-Path ".env")) { Copy-Item ".env.example" ".env" }
Get-Content ".env" | ForEach-Object {
    $l = $_.Trim()
    if ($l -and -not $l.StartsWith("#") -and $l.Contains("=")) {
        $k, $v = $l.Split("=", 2)
        Set-Item -Path ("Env:" + $k.Trim()) -Value ($v.Trim().Trim('"').Trim("'"))
    }
}
if (-not $env:OUTLOOK_CLIENT_ID -or -not $env:OUTLOOK_TENANT_ID) {
    Write-Host ""
    Write-Host "!! Falta configurar el .env: rellena OUTLOOK_CLIENT_ID y OUTLOOK_TENANT_ID."
    Write-Host "   (Como obtenerlos: GUIA DETALLADA, seccion 'Variables de Azure'.)"
    try { Stop-Transcript | Out-Null } catch {}
    Read-Host "Pulsa ENTER para salir"
    exit 1
}

Write-Host "==> 4/4  Iniciando sesion en Microsoft 365 y registrando en Claude..."
Write-Host "     Se abrira el navegador con un codigo: TECLEALO e inicia sesion con TU cuenta."
& $py -m outlook_mcp login
if ($LASTEXITCODE -ne 0) { Fail "El inicio de sesion en Microsoft 365 fallo." }

# taskkill inofensivo: si Claude no esta abierto devuelve error; lo ignoramos a proposito.
taskkill /IM Claude.exe /F /T 2>$null | Out-Null
Start-Sleep -Seconds 3
& $py conectar_claude.py
if ($LASTEXITCODE -ne 0) { Fail "No se pudieron registrar los conectores en Claude." }

Write-Host ""
Write-Host "LISTO. Abre Claude: tendras 'outlook' (conector completo) y 'agente-outlook' (puente con disco)." -ForegroundColor Green
try { Stop-Transcript | Out-Null } catch {}
Read-Host "Pulsa ENTER para salir"
