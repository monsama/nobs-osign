# Architecture

Three parts, joined by one file format and one HTTP API.

```
editor/ui  (web UI, static)
   │  HTTP API (docs/API.md): JSON over POST, Authorization: Bearer
   ▼
editor backend  ── today: editor/OSign.ps1        later: a container (Node, …)
   │  publishes  <path>/signatures.json + <path>/assets/*
   ▼
target  ── GitHub repository (Pages serves it)   or   a folder on a web server
   │  https
   ▼
addin (event handler + pane)  ── reads signatures.json, the user's details from Graph
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
  `assets/…`.
- `pick(config, values, from, kind, prefs)` decides which signature, as the README describes.

## `signatures.json`

```jsonc
{
  "schemaVersion": 1,
  "updated": "2026-10-12T09:00:00.000Z",
  "assetBaseUrl": "https://monsama.github.io/osign/config/",   // images: <this>assets/<name>
  "auth": { "clientId": "…", "tenantId": "…" },               // Entra app for Graph
  "options": { "disableOutlookSignature": true, "allowUserChoice": true, "meetings": false, "insertWithoutProfile": false },
  "fields": [ { "key": "givenName", "label": "Vorname", "ad": "givenName", "graph": "givenName", "sample": "Anna" } ],
  "signatures": [ {
    "id": "standard", "name": "Standard",
    "html": "<p>@@givenName @@sn</p>…",            // what the editor edits
    "compiled": "<!--sig-if:givenName,sn-->…",     // what the add-in renders
    "defaultNew": true, "defaultReply": false, "removeEmptyLines": true,
    "audience": { "departments": [], "domains": [] }
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
4. Publish the draft to a target, and load from it. For GitHub that is one commit through the Git
   Data API, never forced. For a folder, the images are copied first and the configuration is
   written last.
5. Optionally, look up a user in AD for the preview.

`OSign.ps1` does this for one person on their own PC. It listens on 127.0.0.1 only, checks the Host
header (against DNS rebinding), and uses a token that is new on every start and passed in the URL
fragment. It stores the GitHub token with DPAPI.

## Moving the editor into a container

What changes, and what doesn't:

| | PowerShell (now) | Container (later) |
|---|---|---|
| UI | `editor/ui` | the same files |
| API | `OSign.ps1` | a new server with the same routes (e.g. Node) |
| Who may use it | whoever started it (token in the URL) | sign-in in front: Entra ID (App Service / Container Apps authentication, or a reverse proxy with OIDC). The API then accepts that session instead of the start token. |
| GitHub token | DPAPI, per user | a secret of the container (`GITHUB_TOKEN`), or a GitHub App |
| Draft | `%APPDATA%\OSign` | a volume (`OSIGN_DATA`), shared by everyone |
| AD lookup | ADSI on a domain PC | LDAP to a DC, or Graph (`User.Read.All`) |
| Publishing | GitHub or a folder | the same, or straight into the container's own web root, so it can serve the add-in too |

`OSign.ps1` already reads `OSIGN_DATA` and, without Windows, `GITHUB_TOKEN`. So it also runs as it is
in a `mcr.microsoft.com/powershell` container, as a stopgap. It would need a port mapping and a
proxy in front: it only listens on 127.0.0.1 and has no sign-in of its own.

The UI reads the token from `#t=` and sends it as `Authorization: Bearer`. A backend that
authenticates by cookie instead can simply ignore that header.
