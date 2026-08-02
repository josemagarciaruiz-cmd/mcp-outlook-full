# Instalador COMPLETO de MCP Outlook (conector 63 tools + agente) para Windows.
# Ejecutar con Claude CERRADO (el .bat lo cierra por ti).
$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "==> 1/3  Entorno y dependencias (conector + agente en un solo venv)..."
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -q --upgrade pip
& .\.venv\Scripts\pip.exe install -q -e .

if (-not (Test-Path ".env")) { Copy-Item ".env.example" ".env" }

# Cargar variables del .env
Get-Content ".env" | ForEach-Object {
    $l = $_.Trim()
    if ($l -and -not $l.StartsWith("#") -and $l.Contains("=")) {
        $k, $v = $l.Split("=", 2)
        $v = $v.Trim().Trim('"').Trim("'")
        Set-Item -Path ("Env:" + $k.Trim()) -Value $v
    }
}

if (-not $env:OUTLOOK_CLIENT_ID -or -not $env:OUTLOOK_TENANT_ID) {
    Write-Host ""
    Write-Host "!! Falta configurar el .env."
    Write-Host "   Abre  $PSScriptRoot\.env  y rellena OUTLOOK_CLIENT_ID y OUTLOOK_TENANT_ID."
    Write-Host "   (Como obtenerlos: GUIA DETALLADA / README, seccion 'Variables de Azure'.)"
    Read-Host "Pulsa ENTER para salir"
    exit 1
}

Write-Host ""
Write-Host "==> 2/3  Iniciando sesion en Microsoft 365 (una sola vez)."
Write-Host "         Se abrira el navegador con un codigo: TECLEALO e inicia sesion con TU cuenta."
& .\.venv\Scripts\python.exe -m outlook_mcp login

Write-Host ""
Write-Host "==> 3/3  Cerrando Claude y registrando los conectores..."
taskkill /IM Claude.exe /F /T 2>$null | Out-Null
Start-Sleep -Seconds 3
& .\.venv\Scripts\python.exe conectar_claude.py

Write-Host ""
Write-Host "LISTO. Abre Claude. Tendras 'outlook' (63 herramientas) y 'agente-outlook' (puente con disco)."
