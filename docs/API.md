# Editor API

The contract between `editor/ui` and any backend (`editor/OSign.ps1` today).

- Every call is `POST /api/<name>` with `Content-Type: application/json` and
  `Authorization: Bearer <token>`. The body is a JSON object, `{}` when there is nothing to send.
- Every answer is JSON: `{ "ok": true, … }` or `{ "ok": false, "error": "a sentence for the user" }`.
  A wrong or missing token gets HTTP 401. A request from another origin, another Host or another
  method gets 403.
- `GET /assets/<name>` serves a draft image, `GET /photos/<name>` a photo from the photos folder
  (no token: an `<img>` cannot send one).

## ping

Keeps the server from stopping when idle. → `{ ok }`

## info

→ `{ ok, backend: "powershell", version, dataDir, features: { adLookup: bool } }`

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

→ `{ ok, settings: { folder, photosFolder }, folderFound: bool, photosFound: bool }`

`folder`: where Publish writes (the synced OSign folder). `photosFolder`: where the preview finds
the photos; empty means `<folder>\photos`.

## settings/save

`{ settings: { folder, photosFolder } }` → like `settings/get`. Paths have to be full paths; a folder
that is not there (yet) is accepted, and Publish then says so.

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

## ad/export

`{ fields: [{ key, ad }], groups: [string] }` → `{ ok, users, written, unchanged, removed }`

For signatures whose values come from AD. Writes `users\<mail>.json` into the publish folder for
every active user with a mail address: `{ "values": { <key>: string }, "groups": [<the named groups
the user is in, nested ones included, lower case>] }`. Only files that differ are written; files of
users no longer found are removed. The file name is the mail address in lower case, characters
other than `a-z 0-9 @ . _ -` as `_` (`userFileName()` in the shared engine).

## quit

Stops the server. → `{ ok }`
