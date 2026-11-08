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
- `withCampaign(config, sig, ctx)` puts the banner of the campaign running that day into it:
  where it has `{{campaign}}`, else below (or above) it; none running, `{{campaign}}` goes.
  `campaignFor()` picks the campaign (its days, signatures, audience, new mail only, the person's
  language), `withUtm()` adds the link tracking, `campaignsEnding()` finds the ones about to end.

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
    "audience": { "groups": [], "departments": [], "domains": [] },
    "validFrom": "", "validUntil": ""              // a campaign's days (YYYY-MM-DD), or always
  } ],
  "colors": ["#1d4ed8"],                           // the house colours (the editor's palette)
  "campaigns": [ {                                 // banner campaigns (09-campaigns.js)
    "id": "c-xmas", "name": "Weihnachten 2026", "validFrom": "2026-12-01", "validUntil": "2026-12-31",
    "html": "<div>…</div>", "compiled": "…",        // the banner, made from design
    "design": { "tpl": "image", "img": "xmas.png", "width": 400, "link": "https://…" },
    "utm": true, "signatures": [],                 // none: all signatures
    "position": "below", "newOnly": false,         // where without {{campaign}}; not in replies
    "days": [4],                                   // only on these weekdays (0 Sunday … 6 Saturday)
    "more": [ { "html": "…", "compiled": "…", "design": {} } ],  // banners taking turns, mail by mail
    "audience": { "groups": [], "departments": [], "domains": [], "everyone": true },
    "langs": { "en": { "html": "…", "compiled": "…", "design": {} } }
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
   `server/setup-template.ps1`). The server signs in to Entra ID with its own certificate: it
   makes the key and a self-signed certificate (`server/cert.js`), the script puts only the
   certificate on the app, and each token request carries a signed client assertion.
8. It keeps an audit log (`server/audit.js`), a backup locked with a passphrase
   (`server/backup.js`), four eyes before publishing (`approval.json`), what changed between
   versions (`server/diff.js`), and asks GitHub twice a day whether a newer release is out
   (`server/updates.js`, `OSIGN_UPDATE_CHECK`).
9. It makes backups by itself every day (`server/autobackup.js`), publishes at a time set
   beforehand (`server/routes-release.js`), tells a webhook (Teams, Power Automate) when something
   needs doing (`server/notify.js`: outgoing only, https, the address sealed), and lists who is
   signed in for the local account `admin` (`server/auth.js`, sessions by a hash reference only).

`server/app.js` makes the parts and answers every request; the API's routes are in modules by
topic, each given what it needs from it: `routes-publish.js` (publishing, four eyes),
`routes-release.js` (a time set), `routes-connections.js` (the connections, the setup script),
`routes-addin.js` (the add-in's files), `routes-history.js`, `routes-photos.js`,
`routes-directory.js` (AD and Entra ID) and `routes-backup.js`.

## The editor

`editor/ui/app/NN-*.js` are one script by topic, put together in order and served as
`/ui/app.js` (one IIFE: `01-base.js` opens it, `11-start.js` closes it): the state, saving and the
list (`01`), the design editor and its undo (`02-design.js`), PNGs with a palette (`02-png.js`), its tools: house colours, spacing, the bars for a link, a cell and an image's place (`02-tools.js`), the preview (`03`), photos (`04`), signatures
and what is published (`05`), import (`06`), fields, options and users (`07`), building blocks and
Settings (`08`), banner campaigns (`09-campaigns.js`), the setup assistant (`09-wizard.js`), classic Outlook's room and the ways to fit it (`10-classic.js`),
the draft's images: used, made lighter, put in another's place (`10-images.js`), publishing, search and the dialogs (`10-publish.js`),
replace in all signatures (`10-replace.js`), undo for the whole draft and ⟳ (`10-undo.js`), and
the start (`11`). `editor/ui/i18n.js` holds the German; the page is written in English.

The add-in's pane and notice speak German (`addin/src/taskpane.js`, `addin/src/common.js`),
French and Italian (`addin/src/lang.js`), else English, as Outlook's display language says.

The Docker setup is `Dockerfile`, `docker-compose.yml` and `deploy-nobs-osign.sh`; updates,
backups and restores are `scripts/*.sh`.
