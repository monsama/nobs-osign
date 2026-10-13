# OSign

Central email signatures for **classic and new Outlook**, managed in one place.

- **The editor** (`editor/OSign.ps1`): a single PowerShell script, no installation needed. It
  starts a small web server on 127.0.0.1 and opens a web interface (German or English) where you design the signatures, set who gets which
  one, and publish them.
- **The Outlook add-in** (`addin/`): it inserts the right signature by itself in every new message,
  reply and forward, and switches signatures when the From address changes. Its pane lets users pick
  another signature or choose their own defaults. It fills the placeholders (`@@givenName`,
  `@@mobile`, and so on) with each user's details from Microsoft Graph.

![The editor](docs/editor.png)

**It runs without GitHub.** GitHub only holds the source code and the releases. The add-in's
files are made by the editor and served from a web server of the company's own choosing (its own
web server, Azure Static Web Apps, an Azure storage account…). The signatures, images and photos
stay in the company's own Microsoft 365: in a folder `OSign` in the SharePoint home site. The
add-in reads that folder with each user's own sign-in, and embeds the images in the mail.

```
 OSign.ps1 (editor, admin PC)        Company SharePoint (home site)       Outlook (every user)
 ┌──────────────────────────┐ write  ┌───────────────────────────┐ Graph  ┌───────────────────────────┐
 │ web UI on 127.0.0.1      │ ─────> │ OSign\signatures.json     │ ─────> │ OSign add-in              │
 │ publishes into the       │ (One-  │ OSign\assets\*.png        │ user's │  new mail / reply → sign  │
 │ synced OSign folder      │ Drive) │ OSign\photos\<user>.jpg   │ sign-in│  details, groups ← Graph  │
 └──────────────────────────┘        └───────────────────────────┘        │  images embedded (cid:)   │
                                                                          └───────────────────────────┘
 The company's HTTPS web server: the add-in's code (no data), made by the editor ──────┘
```

## Where it works

| Outlook | Automatic signature | Pane |
|---|---|---|
| New Outlook for Windows | yes | yes |
| Outlook on the web | yes | yes |
| Classic Outlook for Windows (Microsoft 365, current channel) | yes | yes |
| Outlook for Mac (new Outlook) | yes | yes |
| Outlook for iOS / Android | not yet | - |

All of these need Exchange Online (Microsoft 365) mailboxes.

## Fields

These are the placeholders the signatures used before, mapped to where the add-in reads each value.

| Field | Placeholder | AD attribute | Microsoft Graph |
|---|---|---|---|
| Vorname | `@@givenName` | givenName | `givenName` |
| Name | `@@sn` | sn | `surname` |
| Beschreibung 1 | `@@description` | description | *(none, see below)* |
| Beschreibung 2 | `@@department` | department | `department` |
| Beschreibung 3 | `@@title` | title | `jobTitle` |
| Handy | `@@mobile` | mobile | `mobilePhone` |
| Direkte Nr. ext. | `@@homePhone` | homePhone | *(none, see below)* |
| Abwesenheit | `@@info` | info | *(none, see below)* |
| Mail | `@@mail` | mail | `mail` |

**Where the values come from** is a choice in the editor (*Add-in*): **Entra ID** (Microsoft Graph,
each field by its Graph source), **Active Directory** (the editor exports each user's fields into
`OSign\users\<mail>.json`; `OSign.ps1 -ExportUsers` keeps them current from the Task Scheduler), or
**both**: Graph where a field has a Graph source, AD for the rest.

Graph has no `description`, `homePhone` or `info`. To use them, sync them with Entra Connect,
either into `onPremisesExtensionAttributes.extensionAttribute1`…`15` or as directory extensions
(`extension_<appId>_description`). Then enter that path as the field's Graph source in the editor.
Until then those fields stay empty in Outlook. The editor's preview can still show them, because it
reads them from AD.

`@@sAMAccountName` (Benutzername, Graph `onPremisesSamAccountName`) is there as well. It also names
each person's photo: a template image `photos/@@sAMAccountName.jpg` becomes
`OSign\photos\<username>.jpg`, embedded in the mail; whoever has none gets the signature without.

A line whose fields all come out empty is left out, so nobody's signature shows "Mobile:" with no
number. This works for paragraphs, lines split by line breaks, and table rows.

## Getting started

[docs/SETUP.md](docs/SETUP.md) has every step. In short:

1. **SharePoint:** in the home site's document library, make a folder `OSign` (readable for all
   employees). Sync it to the admin PC (*Sync* or *Add shortcut to My files*). Copy the photos into
   `OSign\photos` (`<username>.jpg`).
2. **Entra ID:** register an app (single-page application, redirect URI
   `brk-multihub://<host of the add-in>`) with Graph `User.Read`, `Sites.Read.All` and, for groups,
   `GroupMember.Read.All`, delegated, with admin consent.
3. **Editor:** unpack the release zip on Windows and run
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\OSign.ps1
   ```
   In Settings, pick the synced `OSign` folder. Import or make the signatures and click **Publish**.
4. **The add-in's files:** in Settings → *Outlook add-in*, enter the address they will be served
   from and the app's client ID, click **Make the add-in files**, and copy that folder to the web
   server.
5. **Microsoft 365 admin center:** Settings → Integrated apps → Upload custom apps, with
   `https://<host of the add-in>/manifest.xml`, and assign it to the users.

## How a signature is chosen

Each signature can be the default for new mail, for replies and forwards, or both. It can also be
limited to **AD groups** (nested groups count), departments or sender domains. For each message:

1. If the editor allows user choice and the user picked a signature (or *none*) in the pane, that
   one is used.
2. Otherwise the default for that kind of message is used. When several defaults apply, the one for
   fewer people wins: sender domain before group before department before everyone.

Outlook's own signature is switched off for messages the add-in signs (an option in the editor), so
a mail never gets two.

## Importing existing templates

**⤓** next to *Signatures* imports `.htm` templates with their images: from CI-Sign (it keeps them
next to `CI-Sign.exe`, as UTF-16, with images in `images\`) or from classic Outlook. Pick the
folder, or the template files and images together. A per-person image (`images/@@sAMAccountName.jpg`)
becomes the photo from `OSign\photos`. Not imported (yet): CI-Sign's text blocks
(`{##Name}…{end ##Name}` in `__intern.htm`).

## Privacy

- GitHub holds only code; nothing at run time comes from it. The add-in's files on the company's
  web server are only code too, with the app registration ID filled in (it is not a secret).
- The signatures, images and photos are in the company's SharePoint, readable for its employees.
  The add-in reads them, and each user's details and groups, with that user's own sign-in
  (delegated permissions): it can never read more than the user can.
- Images go into the mail itself, so no recipient's mail program loads anything from anywhere.
- The add-in keeps a copy of the user's details in their own mailbox settings, read again after
  15 minutes and used when Graph cannot be reached; and the images in the browser's storage.

## Hosting the add-in elsewhere

Any HTTPS web server can serve the add-in's files; [docs/SETUP.md](docs/SETUP.md) step 3 lists
some. Moving them means: make the files again for the new address, add
`brk-multihub://<new host>` to the app registration, and upload the new `manifest.xml` in the admin
center. Without the editor, the build does the same:

```
npm ci
BASE_URL=https://intranet.example.com/osign/ OSIGN_CLIENT_ID=<app id> node addin/build.mjs
```

and serve `dist/site/`. `OSIGN_SP_SITE` and `OSIGN_SP_FOLDER` choose another SharePoint site or
folder for the signatures.

The editor is meant to move into a container later. The web UI talks to its backend only through
the HTTP API in [docs/API.md](docs/API.md), so a second backend can replace `OSign.ps1` without the
UI changing. [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) describes the parts.

## Development

```
npm ci
npx playwright install chromium   # once, for the browser tests
npm test                          # builds dist/template, then runs every test
```

The tests need `pwsh` (PowerShell 7). CI also runs the editor under Windows PowerShell 5.1.

| Path | What |
|---|---|
| `shared/signature-core.js` | Placeholders, empty lines, who gets which signature. The editor's preview and the add-in both run this one file. |
| `editor/OSign.ps1` | The editor's backend: local web server, draft, folder publishing, AD lookup, the add-in's files |
| `editor/ui/` | The editor's web UI |
| `addin/` | The Outlook add-in: event handler, pane, manifest template, build (`build.mjs`, `fill.mjs`) |
| `tests/` | Engine, add-in runtime (with a stand-in Graph and SharePoint), backend API, editor and pane in a browser |

## License

GPL-3.0-or-later, see [LICENSE](LICENSE).
