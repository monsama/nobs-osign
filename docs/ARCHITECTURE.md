# Architecture

Three parts, joined by one file format and one HTTP API.

```
editor/ui  (web UI, static)
   │  HTTP API (docs/API.md): JSON over POST, a session cookie (AD group or the local admin)
   ▼
server/  (Node, in Docker, behind the shared proxy)
   │  publishes  signatures.json + assets/* (+ users/*) through Microsoft Graph (Sites.Selected)
   ▼
the OSign folder in one SharePoint site   + photos/<user>.jpg (by hand)
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

## The server

`server/` is the editor's backend, in Node:

1. It serves `editor/ui/` (only after the sign-in, the sign-in page and `/health` aside),
   `shared/signature-core.js`, the draft's images at `/assets/*` and the photos at `/photos/*`.
2. It answers the API in [API.md](API.md), with a session cookie (HttpOnly, SameSite=Strict) and
   same-site checks. Who may sign in: members of an AD group, and the local account `admin`.
3. It keeps a draft in the data volume (`OSIGN_DATA`): the configuration text, its images, the
   settings, the history of what was published.
4. It publishes into the SharePoint folder through Microsoft Graph with its own app registration,
   which has write access to that one site (`Sites.Selected`): images first, the configuration
   last, only what changed (`server/graph.js`, `server/store.js`).
5. It reads Active Directory over LDAP with a read-only service account: the preview, group names,
   and the users' values written into `users/` on Publish and every day (`server/directory.js`,
   `server/users.js`); and, when allowed, a person from Entra ID for the preview.
6. It makes the add-in's files with `addin/fill.mjs` and uploads them to Cloudflare Pages
   (`server/addin.js`).
7. Its connections (AD, Graph, Cloudflare) are set in the editor by the local account `admin`,
   sealed with the server's key (`server/connections.js`); the setup script for Microsoft 365
   makes both app registrations and sends its result back (`server/setup.js`,
   `server/setup-template.ps1`).

The Docker setup is `Dockerfile`, `docker-compose.yml` and `deploy-nobs-osign.sh`; updates,
backups and restores are `scripts/*.sh`.
