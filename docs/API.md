# Editor API

The contract between `editor/ui` and its backend, `server/` (the sections below the server's own are
the routes it shares with the editor; `pick-folder`, `sp/resolve`, `sp/local` and `quit` were the
Windows editor's and answer `ok: false`).

- Every call is `POST /api/<name>` with `Content-Type: application/json` and
  `Authorization: Bearer <token>`. The body is a JSON object, `{}` when there is nothing to send.
- Every answer is JSON: `{ "ok": true, … }` or `{ "ok": false, "error": "a sentence for the user" }`.
  A wrong or missing token gets HTTP 401. A request from another origin, another Host or another
  method gets 403.
- `GET /assets/<name>` serves a draft image, `GET /photos/<name>` a photo from the photos folder
  (no token: an `<img>` cannot send one).

## The server version (`server/`)

The container's backend answers the same API, behind a sign-in instead of a start token. It ignores
the `Authorization` header; a session cookie (`osign_session`, HttpOnly, SameSite=Strict) counts
instead. Every call except the three below needs one and gets HTTP 401 without it. A request whose
`Origin` is not this site, or whose `Sec-Fetch-Site` says another site, gets 403. `GET /` without a
session redirects to `/login`; `GET /health` answers `ok` to anyone, for health checks.

- **login** `{ login, password }` → `{ ok, user: { login, name } }` and the cookie. `admin` is the
  local account (its password from `OSIGN_ADMIN_PASSWORD_FILE`); any other name is checked against
  Active Directory, and only members of `OSIGN_LDAP_GROUP` get in (403 otherwise). A wrong password
  gets 401; five for one name within 15 minutes, 429; a directory that does not answer, 503.
- **login-methods** → `{ ok, local: bool, ad: bool }`: what the sign-in page offers.
- **logout** → `{ ok }`, and the cookie removed.

`info` answers `backend: "node"`, `user: { login, name }`, `features.signOut: true`, and where
publishing goes: `features.publishTarget` is `"sharepoint"` (through Microsoft Graph, with the
server's own app registration `features.graphClientId`) or `"folder"` (the server's own folder,
when Graph is not set up; `features.graphProblem` says what is missing).

Publishing goes to the SharePoint site and folder of the add-in (`spSite`, `spFolder`): `publish`,
`load-published` and `GET /photos/<name>` use it. `folder` in the settings answers where that is,
in words (`SharePoint contoso.sharepoint.com/sites/IT › OSign`); `folder`, `photosFolder` and
`addinFolder` sent are ignored. `settings/get` also answers `folderError` when SharePoint cannot be
asked.

- **ad/lookup**, **ad/groups**, **ad/read**, **ad/snapshot**, **ad/export** work as above, with the
  server's service account over LDAP (`features.adLookup`; `features.adProblem` says what is
  missing). An empty `query` in `ad/lookup` means the one signed in. `users/` is in the same
  SharePoint folder; the server keeps a copy of what it holds (`users-mirror.json`), so a read asks
  SharePoint only for the files that changed since. `features.adExport` is `{ at, last }` when the
  server exports by itself every day (`OSIGN_AD_EXPORT_AT`); `last` is `{ at, ok, result | error }`.
- **addin/build** makes the add-in's files on the server (into the data volume, with
  `addin/fill.mjs`); `folder` is empty. **addin/check** compares what the add-in's address serves with
  them. **addin/upload** → `{ ok, id, url, uploaded, files }` uploads them to Cloudflare Pages
  (Direct Upload; `features.pagesUpload`, `features.pagesProblem`). `GET /addin.zip` gives them as a
  ZIP, to upload by hand.
- **connections/get** → `{ ok, keyOk, fields, secrets, ca, passwordPinned, problems }`: the
  server's connections (server/connections.js), for the local account `admin` only (anyone else:
  `ok: false`). `fields.<name>` is `{ value, env }`: what is saved, and what `.env` gives (it counts
  when `value` is empty). `secrets.<name>` is `{ saved, readable, file, fileName }`, never the value.
  `problems` says per connection (`signIn`, `directory`, `graph`, `pages`) what is missing, or `''`.
- **connections/save** `{ values: { <name>: text }, secrets: { <name>: text | null }, ca: base64 | null }`
  → like `connections/get`. A value `''` goes back to `.env`; a secret `null` is removed, one not
  sent stays; `ca` is a CA certificate file (.cer or PEM). Takes effect at once.
- **setup/start** `{ origin, groups }` → `{ ok, id, until }`, for the local account `admin`: a setup
  script for the folder and the add-in's address as saved, valid for an hour; `GET
  /setup/nobs-osign-setup.ps1?id=<id>` downloads it. **setup/status** → `{ state: none | waiting |
  failed | done, problem, result }` (never the secret). **setup/finish** is the script's: without a
  session, with its one-time code; its result goes into the settings and the server connections.
- **entra/lookup** `{ query, fields: [{ key, graph }] }` → `{ ok, values, displayName, groups }`: a real
  person from Entra ID for the preview, by UPN or mail, through the server's app (it needs
  User.Read.All and GroupMember.Read.All, which the setup script gives when asked to).
- **history/list** → `{ ok, versions: [{ id, at, by, message, signatures, images }] }`: what was
  published, newest first (each publish that changed something, with its images; the last 50).
  **history/restore** `{ id }` puts that version and exactly its images into the draft.
  **history/diff** `{ id, against?: 'draft' }` → `{ ok, changes }`: what changed in that version
  against the one before it (or the draft against it). Each change is `{ kind: added | removed |
  renamed | html | rules | fields | options, name, detail?, lines? }`; `lines` are the HTML's lines
  that differ, `{ op: '-' | '+', text }`.
- **Four eyes** (setting `requireApproval: 'on'`, the local `admin`'s; `features.requireApproval`):
  `publish` answers `ok: false`; **approval/request** `{ configJson, message }` saves the draft and
  asks for approval; **approval/get** → `{ ok, required, pending: { by, at, message, current, mine,
  mayApprove, changes } | null }` (`changes` against what is published; `current: false` once the
  draft changed since); **approval/approve** publishes it, by someone other than who asked, and only
  while the draft is as it was; **approval/cancel** withdraws or rejects it.
- **audit/list** `{ limit?, what?, failed? }` → `{ ok, entries: [{ at, who, ip, what, detail, ok }] }`,
  newest first, for the local `admin`: sign-ins (and failed ones), publishing, approvals, restores,
  connections, settings, the setup script, the add-in's upload, backups. Never a password, a token
  or a signature's content (server/audit.js, `audit.log` in the data volume).
- **backup/export** `{ passphrase }` → `{ ok, name, data }` (base64 of the file) and
  **backup/restore** `{ data, passphrase }` → `{ ok, at, version, files }`, for the local `admin`:
  the data volume (not the audit log) in one file locked with the passphrase (scrypt, AES-256-GCM).
  The server's key goes with it, so a server with another key seals the passwords again with its own.
- `features.update` is `{ version, url, at }` when a newer release is out (`OSIGN_UPDATE_CHECK=1`,
  twice a day one request to GitHub), else `null`. `features.graphCredential` is `certificate` (the
  server's own, server/cert.js, put on the app by the setup script) or `secret`; `features.secretExpires`
  and `features.secretDaysLeft` are for whichever it is. `connections/get` answers `cert: { thumbprint,
  expires } | null`; `connections/save` with `removeCert: true` goes back to a client secret. **connections/cert** → `{ ok, der, thumbprint, expires }`: the server's certificate (base64 DER, made and put in use when there is none), to upload to its app in Entra by hand.
- **connections/detect** → `{ ok, domain, base, host }`: the domain the saved domain controller
  serves, from its RootDSE (no sign-in needed), for the local account `admin`.
  `connections/get` and `connections/save` also answer `derived`: what the empty fields follow from
  the domain (`ldapBind`, `ldapBase`, `ldapDirectoryBase`, `ldapServiceDn`).
- **connections/test** `{ what: ad | signIn | directory | graph | pages }` → `{ ok, message }`: that part
  of the server's connections tried for real, with what is saved (for the local account `admin`).
- **sp/share** `{ url }` → `{ ok, webUrl, name }`: where a sharing link of SharePoint or OneDrive
  (`https://…/:f:/…`) leads, through Graph (`/shares`), as the folder's own address; a link to a
  file, one that is gone, or one the app may not open answers `ok: false` with why.
- **sp/check** `{ spSite, spFolder }` → `{ ok, found, photos, where }`: whether the server reaches
  that folder, before it is saved (the setup assistant). `found: false` means the site is reached
  and the folder is made on the first publish. A site the app has no access to answers `ok: false`. What needs Windows or OneDrive (`pick-folder`, `sp/resolve`, `sp/local`, `quit`)
answers `ok: false` with a sentence saying so.

## ping

Keeps the server from stopping when idle. → `{ ok }`

## info

→ `{ ok, backend: "powershell", version, dataDir, features: { adLookup: bool, addinFiles: bool, pickFolder: bool } }`

`addinFiles`: the add-in's template is there, so `addin/build` can work. `pickFolder`: `pick-folder`
can show the Windows folder dialog.

## draft/get

→ `{ ok, configJson: string|null, assets: [{ name, size }], dirty: bool, publishedAt }`

`publishedAt` is milliseconds since 1970 (0: never).

`configJson` is the configuration as text, exactly as last saved. `dirty` means the draft differs
from what was last published or loaded.

## draft/save

`{ configJson }` → `{ ok, dirty }`. It is refused if the text is not JSON with `schemaVersion` 1, or
is larger than 1.5 MB.

## assets/upload

`{ name, data }`, where `data` is base64 → `{ ok, path: "assets/<name>", name, size }`

PNG, JPG or GIF only, checked by content, at most 1 MB. The name is cut down to `[A-Za-z0-9_-]` and
the extension. If an image with the same name but other content exists, the upload gets a new name
(`logo-2.png`), so signatures using the old one keep it.

## assets/delete

`{ name }` → `{ ok }`

## settings/get

→ `{ ok, settings: { folder, photosFolder, addinUrl, clientId, tenantId, spSite, spFolder, addinFolder, addinHost, setupDone }, folderFound: bool, photosFound: bool }`

`folder`: where Publish writes (the synced OSign folder). `photosFolder`: where the preview finds
the photos; empty means `<folder>\photos`. The rest is for `addin/build`: the address the add-in is
served from, the app registration (client ID, tenant ID or empty), the SharePoint site (`root` or
`<host>.sharepoint.com:/sites/<name>:`, `…:/teams/<name>:`, or a OneDrive,
`<host>-my.sharepoint.com:/personal/<name>:`) and folder the add-in reads, where its files are written,
and the kind of web server (`cloudflare`, `azure`, `iis`, `apache`, `other` or empty), whose caching
file goes with them. `setupDone` is `"1"` once the setup assistant was finished; until then the UI
opens it by itself while `folder` or `clientId` is empty.

## settings/save

`{ settings: { … } }` → like `settings/get`. Only the settings sent change; the others stay as they are.
Paths have to be full paths; a folder that is not there (yet) is accepted, and Publish then says so.
The add-in's settings are checked as they are (an https address, GUIDs) and may be empty.

## addin/build

`{}` → `{ ok, folder, files, version, manifestUrl, redirectUri }`

Writes the Outlook add-in's files for the saved settings into `addinFolder`: the template
(`addin` next to the script, `dist/template` in the repository) with its blanks filled in, as
`addin/fill.mjs` does it. Files already in that folder are overwritten, others left alone. Needs
the address, the client ID and a folder that is not the publish folder. With `addinHost`, that web
server's caching file comes along (`_headers`, `staticwebapp.config.json`, `web.config` or
`.htaccess`, from `hosting/<host>/` in the template), so a new version reaches every Outlook within
minutes.

## sp/resolve

`{ path }` → `{ ok, found, location: { site, siteName, host, library, folder, url, defaultLibrary, uncertain } }`

Where the local folder `path` is in SharePoint, from what OneDrive notes about the libraries it
syncs to this PC (`HKCU\Software\SyncEngines\Providers\OneDrive`: each library's local folder
and address). `site` and `folder` are what the add-in's settings take; `defaultLibrary` says
whether it is the site's main library (the only one the add-in reads). `found: false` when no
synced library holds the folder.

## sp/local

`{ url }` → `{ ok, found, path, exists }`

Where the SharePoint folder at `url` (`https://<host>/sites/<site>/<library>/<folder>`) is on this
PC: in the place OneDrive syncs it to. Shortcuts (*Add shortcut to My files*) are left out: OneDrive
notes only their library, not which folder of it they are. `sp/resolve` answers
`location.uncertain: true` for a folder in one.

## pick-folder

`{ start, title }` → `{ ok, path }`

Shows the Windows folder dialog on this PC and answers the folder chosen, or `""` when cancelled.
Windows only (`features.pickFolder`).

## addin/check

`{}` → `{ ok, checks: [{ what, ok, detail }], allOk }`

Loads `manifest.xml`, `launchevent.js`, `taskpane.js` and `taskpane.html` from the add-in's address
and compares each, byte for byte, with the files in `addinFolder`; also whether the manifest comes
with `Cache-Control: no-cache`. `allOk` leaves the caching header out.

## publish

`{ configJson }` → `{ ok, changed, folder }`

Saves the draft, then writes `assets\*` and, last, `signatures.json` into the folder: only files that
differ, so OneDrive uploads nothing that did not change. Images no longer in the draft are removed
there; `photos\` and anything that is not an image are left alone. `changed: false` means nothing
differed.

## load-published

→ like `draft/get`. Replaces the draft and its images with what is in the folder. The previous
draft is kept as `draft.json.bak`.

## ad/lookup

`{ query, fields: [{ key, ad }] }` → `{ ok, values: { <key>: string }, displayName, groups: [string] }`

Finds one user by sAMAccountName, mail or UPN (an empty query means the current user). It returns
the listed AD attributes under the field keys, and the names (name and sAMAccountName) of every
group the user is in, nested ones included. A backend that cannot do this answers `ok: false` and
reports `features.adLookup: false`.

## ad/groups

`{ query }` → `{ ok, groups: [string] }`: up to 30 group names (sAMAccountName) starting with
`query`, for picking who gets a signature.

## ad/read

`{ fields: [{ key, ad }], groups: [string] }` → `{ ok, readAt, written, counts: { new, changed, unchanged, removed }, users: [{ mail, status, values, groups, old }], removed: [mail] }`

For signatures whose values come from AD. Reads every active user with a mail address and works out
what an export would do, compared with `users\` in the publish folder; `old` holds the values now
there for a changed user. Nothing is written to the folder: what was read is kept by the backend,
for **ad/export**.

## ad/snapshot

`{ fields, groups }` → `{ ok, snapshot: null | <like ad/read>, matches }`: the last read, and whether
it was read for these fields and groups.

## ad/export

`{ fields, groups }` → `{ ok, users, written, unchanged, removed }`

Writes exactly the last read into `users\<mail>.json`: `{ "values": { <key>: string }, "groups":
[<the named groups the user is in, nested ones included, lower case>] }`, and removes the files of
users no longer found. Refused when nothing was read, when it was read for other fields or groups,
and when it was written already (or a scheduled `-ExportUsers` ran since): an old read never goes
over newer values. The file name is the mail address in lower case, characters other than
`a-z 0-9 @ . _ -` as `_` (`userFileName()` in the shared engine).

## quit

Stops the server. → `{ ok }`
