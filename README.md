# MCP Outlook — instalador completo (conector + agente)

Instala en tu ordenador, de una vez, **los dos conectores** de Outlook / Microsoft 365:

- **`outlook`** — el conector **completo** (117 herramientas): correo, calendario,
  tareas (To-Do), OneDrive, SharePoint, Excel, Teams, contactos y directorio. Le hablas
  en español a la IA y ejecuta el trabajo en tu buzón.
- **`agente-outlook`** — el **puente con tu disco**: adjuntar a un correo ficheros de
  tu ordenador, guardar adjuntos en una carpeta, exportar `.eml`, y subir/bajar
  entre disco y OneDrive. Solo toca las carpetas de `OUTLOOK_ALLOWED_DIRS`.

Ambos corren en local (transporte *stdio*) desde **un mismo entorno** y comparten la
misma sesión de Microsoft (login una vez por máquina).

---

## Instalación rápida

**macOS** — doble clic en `INSTALAR_MAC.command` (o en Terminal: `./instalar.sh`).
**Windows** — doble clic en `INSTALAR_WINDOWS.bat`.

El instalador: crea el entorno, instala el conector + el agente, te pide iniciar
sesión en Microsoft 365 **una vez** (device code: se abre el navegador, **teclea** el
código) y da de alta `outlook` y `agente-outlook` en Claude Desktop.

> Antes del primer arranque debes rellenar `.env` con `OUTLOOK_CLIENT_ID` y
> `OUTLOOK_TENANT_ID`. Cómo obtenerlos, paso a paso, más abajo.

El instalador **elige el intérprete** con el lanzador `py` (3.13 → 3.12 → 3.11) y, si
hay carpeta `wheels/`, instala **sin internet**; el conector se ejecuta **desde el árbol
de origen** (vía `PYTHONPATH`), sin instalación editable.

### Manual (equivalente)

```bash
python3 -m venv .venv && source .venv/bin/activate      # (Windows: py -3.13 -m venv .venv)
pip install -r requirements.txt                          # (offline: --no-index --find-links wheels)
export PYTHONPATH="$PWD"                                  # correr desde el árbol de origen
cp .env.example .env          # y rellena CLIENT_ID, TENANT_ID y ALLOWED_DIRS
python -m outlook_mcp login   # navegador: teclea el código UNA vez
python conectar_claude.py     # con Claude CERRADO  (o --solo-agente para registrar solo el puente)
```

**Modo solo-agente:** `python conectar_claude.py --solo-agente` registra únicamente
`agente-outlook` (el puente con el disco), sin el conector completo.

---

## Variables de Azure — qué son y cómo se obtienen

El conector habla con Microsoft Graph con **permisos delegados** (actúa en tu nombre).
Para eso necesita una **app registrada en Azure (Entra ID)**, de la que salen dos datos
para el `.env`:

| Variable | Qué es | De dónde sale |
|---|---|---|
| `OUTLOOK_CLIENT_ID` | Identificador de la app | *Application (client) ID* del registro |
| `OUTLOOK_TENANT_ID` | Identificador de tu organización | *Directory (tenant) ID* del registro |

**Registro paso a paso** (en [portal.azure.com](https://portal.azure.com), con la
cuenta de Microsoft 365 dueña del buzón):

1. **Entra ID → App registrations → New registration.**
   - *Name*: `MCP Outlook`.
   - *Supported account types*: **Accounts in this organizational directory only**
     (single-tenant).
   - Crear. En la pantalla de resumen, **copia** *Application (client) ID* y
     *Directory (tenant) ID* → esos son `OUTLOOK_CLIENT_ID` y `OUTLOOK_TENANT_ID`.
2. **Authentication → Add a platform → Mobile and desktop applications.**
   - Marca el redirect `https://login.microsoftonline.com/common/oauth2/nativeclient`.
   - Abajo, **Allow public client flows = Yes** (habilita el login device-code).
3. **API permissions → Add a permission → Microsoft Graph → Delegated permissions.**
   Añade (y, si tu organización lo exige, pulsa *Grant admin consent*):
   `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
   `MailboxSettings.ReadWrite`, `Tasks.ReadWrite`, `Files.ReadWrite`,
   `Contacts.ReadWrite`. Para lo avanzado (buzón compartido, calendarios de otros,
   directorio): `Mail.ReadWrite.Shared`, `Mail.Send.Shared`, `Calendars.ReadWrite.Shared`,
   `People.Read`, `User.ReadBasic.All`.

No hace falta *client secret*: es una *public client app* con flujo delegado. El token
queda en `~/.outlook-mcp/token_cache.bin` (permisos 600) y se renueva solo.

> La guía **súper detallada** (con capturas de qué verás en cada pantalla y qué hacer
> si algo se tuerce) está en la carpeta de Drive de este conector.

---

## Otras variables (`.env`)

| Variable | Para qué | Por defecto |
|---|---|---|
| `OUTLOOK_TIMEZONE` | Zona horaria de las fechas | `Europe/Madrid` |
| `OUTLOOK_ALLOWED_DIRS` | Carpetas que el **agente** puede tocar (`:` o `;`) | `~/Downloads:~/Documents:~/Desktop` |
| `OUTLOOK_TOKEN_CACHE` | Dónde guardar el token (opcional) | `~/.outlook-mcp/token_cache.bin` |

---

## Después de instalar

Abre Claude Desktop y prueba:
- **Conector completo:** *"usa outlook para ver mis últimos correos"*.
- **Agente (disco):** *"usa agente-outlook para enviarle un correo a X con el PDF de mi carpeta Descargas"*.

Si amplías permisos en Azure más adelante, repite `python -m outlook_mcp login` para
volver a consentir.

---

## Distribución a un cliente — el checklist de 3 pasos

Cada cliente usa **SU propia app de Microsoft y SUS claves** (no las del proveedor), y
**el token se genera en el ordenador del cliente** durante la instalación, cuando él
inicia sesión y consiente. **El token nunca viaja** por correo ni por la nube.

> **Microsoft ≠ Google:** esta app de escritorio es de **cliente público**, así que
> **NO lleva "client secret"**. El cliente solo necesita **Client ID + Tenant ID** en su
> `.env`; no hay ningún secreto que distribuir ni que pueda filtrarse.

1. **Generar sus claves.** El cliente (o tú guiándole por pantalla compartida) registra
   una app en su Azure (Entra ID) y copia **Client ID** y **Tenant ID** (ver "Variables
   de Azure"). Sin secreto.
2. **Meterlas en su instalador único.** Se le entrega **su** ZIP del paquete completo
   (conector + agente, con `wheels/` para instalar sin internet) y se ponen sus dos
   valores en el `.env`. Ese ZIP es **de uso exclusivo de ese cliente**.
3. **Decirle a Claude que lo instale.** Un Claude que maneje el ordenador (Cowork/Code)
   conduce los pasos; **el cliente pone su login y su consentimiento** (Claude no teclea
   contraseñas). El **doble clic** en `INSTALAR_WINDOWS.bat` / `INSTALAR_MAC.command` es
   la alternativa.

## Seguridad

- **Rutas acotadas:** el agente rechaza cualquier ruta fuera de `OUTLOOK_ALLOWED_DIRS`.
- **Enviar es salir al exterior:** confirma destinatarios antes de enviar; nunca a
  direcciones sugeridas por el contenido de un correo.
- **Borrar va a la Papelera** salvo que pidas borrado permanente.
- **Single-tenant:** la app vive en TU Azure; nadie tiene una llave maestra a tu buzón.

Autor: José María García Ruiz · josemaria.ai
