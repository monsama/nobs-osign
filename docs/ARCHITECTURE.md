# Architecture

Three parts, joined by one file format and one HTTP API.

```
editor/ui  (web UI, static)
   │  HTTP API (docs/API.md): JSON over POST, Authorization: Bearer
   ▼
editor backend  ── today: editor/OSign.ps1        later: a container (Node, …)
   │  publishes  signatures.json + assets/*  into a folder
   ▼
the OSign folder in the company's SharePoint (synced by OneDrive)   + photos/<user>.jpg (by hand)
   │  Microsoft Graph, with each user's own sign-in
   ▼
addin (event handler + pane), served as code only from any HTTPS web server of the company's
   reads signatures.json, the images, the user's details and groups; embeds the images (cid:)
```

## The shared engine: `shared/signature-core.js`

The editor's preview and the add-in run the same code, so the preview shows exactly what Outlook
will get.

- `compile(html, document)` runs in the editor, where there is a DOM. It marks every line that holds
  placeholders with `<!--sig-if:a,b-->…<!--/sig-if-->`. A line is a paragraph, a `<br>`-separated
  line, or a table row.
- `render(compiled, values, opts)` uses string operations only. Classic Outlook runs the event
  handler in a JavaScript-only runtime with no DOM. It removes the marked lines whose fields are all
  empty, fills `@@placeholders` (HTML-escaped), and puts the published image address in front of
  `assets/…` and `photos/…` where the editor serves them. The add-in leaves those paths and embeds
  the images instead (`localImages()` lists them).
- `pick(config, values, from, kind, prefs)` decides which signature, as the README describes.

## `signatures.json`

```jsonc
{
  "schemaVersion": 1,
  "updated": "2026-10-12T09:00:00.000Z",
  "options": { "disableOutlookSignature": true, "allowUserChoice": true, "meetings": false, "insertWithoutProfile": false },
  "fields": [ { "key": "givenName", "label": "Vorname", "ad": "givenName", "graph": "givenName", "sample": "Anna" } ],
  "signatures": [ {
    "id": "standard", "name": "Standard",
    "html": "<p>@@givenName @@sn</p>…",            // what the editor edits
    "compiled": "<!--sig-if:givenName,sn-->…",     // what the add-in renders
    "defaultNew": true, "defaultReply": false, "removeEmptyLines": true,
    "audience": { "groups": [], "departments": [], "domains": [] }
  } ]
}
```

The backend stores this file as the text the UI sent and publishes it unchanged. It checks only that
the text is JSON with `schemaVersion` 1. All other meaning lives in the UI and in the shared engine.

## The editor backend's job

Whatever runs it, a backend has to:

1. Serve `editor/ui/` at `/` and `/ui/*`, `shared/signature-core.js` at `/shared/signature-core.js`,
   and the draft's images at `/assets/*`.
2. Answer the API in [API.md](API.md), with authentication.
3. Keep a draft: the configuration text and its images.
4. Publish the draft into the target folder (images first, the configuration last; only what
   changed is written), and load from it.
5. Optionally, look up a user in AD for the preview.

`OSign.ps1` does this for one person on their own PC. It listens on 127.0.0.1 only, checks the Host
header (against DNS rebinding), and uses a token that is new on every start and passed in the URL
fragment. It holds no credentials: writing to SharePoint is OneDrive's job.

## Moving the editor into a container

What changes, and what doesn't:

| | PowerShell (now) | Container (later) |
|---|---|---|
| UI | `editor/ui` | the same files |
| API | `OSign.ps1` | a new server with the same routes (e.g. Node) |
| Who may use it | whoever started it (token in the URL) | sign-in in front: Entra ID (App Service / Container Apps authentication, or a reverse proxy with OIDC). The API then accepts that session instead of the start token. |
| Writing to SharePoint | the synced folder (OneDrive) | Microsoft Graph with an app permission for that one site (`Sites.Selected`) |
| Draft | `%APPDATA%\OSign` | a volume (`OSIGN_DATA`), shared by everyone |
| AD lookup | ADSI on a domain PC | LDAP to a DC, or Graph (`User.Read.All`) |

`server/` is that new backend, in Node, under way: the API, the sign-in (Active Directory group
or the local account `admin`), the draft in `OSIGN_DATA`, and publishing into SharePoint through
Microsoft Graph (`server/graph.js`, `Sites.Selected`), and Active Directory over LDAP with a service
account (`server/directory.js`, `server/users.js`: preview, groups, the users' values, a daily export)
are there, as are the add-in's files (`server/addin.js`: made with `addin/fill.mjs`, uploaded to
Cloudflare Pages) and the Docker setup (`Dockerfile`, `docker-compose.yml`, `deploy-osign.sh`).

`OSign.ps1` already reads `OSIGN_DATA`. So it also runs as it is
in a `mcr.microsoft.com/powershell` container, as a stopgap, publishing into a mounted folder. It
would need a port mapping and a proxy in front: it only listens on 127.0.0.1 and has no sign-in of
its own.

The UI reads the token from `#t=` and sends it as `Authorization: Bearer`. A backend that
authenticates by cookie instead can simply ignore that header.
