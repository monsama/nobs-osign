# Editor API

The contract between `editor/ui` and any backend (`editor/OSign.ps1` today).

- Every call is `POST /api/<name>` with `Content-Type: application/json` and
  `Authorization: Bearer <token>`. The body is a JSON object, `{}` when there is nothing to send.
- Every answer is JSON: `{ "ok": true, … }` or `{ "ok": false, "error": "a sentence for the user" }`.
  A wrong or missing token gets HTTP 401. A request from another origin, another Host or another
  method gets 403.
- `GET /assets/<name>` serves a draft image (no token: an `<img>` cannot send one).

## ping

Keeps the server from stopping when idle. → `{ ok }`

## info

→ `{ ok, backend: "powershell", version, dataDir, features: { adLookup: bool, tokenStore: "dpapi"|"env" } }`

## draft/get

→ `{ ok, configJson: string|null, assets: [{ name, size }], dirty: bool, publishedAt, commitUrl }`

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

→ `{ ok, settings: { target: { kind: "github"|"folder", owner, repo, branch, path, folder }, publicBaseUrl }, tokenSet, tokenFromEnv }`

The token itself is never returned.

## settings/save

`{ settings, githubToken }` → like `settings/get`. `githubToken`: `null` keeps the stored token, `""`
removes it, and any other string replaces it.

## publish

`{ configJson, message }` → `{ ok, changed, commitUrl }`

Saves the draft, then writes `<path>/signatures.json` and `<path>/assets/*` to the target. Images no
longer in the draft are removed there. Other files are left alone. On GitHub this is one commit; if
the branch moved in the meantime it is refused, not forced. `changed: false` means nothing differed.

## load-published

→ like `draft/get`. Replaces the draft and its images with what is on the target. The previous
draft is kept as `draft.json.bak`.

## ad/lookup

`{ query, fields: [{ key, ad }] }` → `{ ok, values: { <key>: string }, displayName }`

Finds one user by sAMAccountName, mail or UPN (an empty query means the current user). It returns
the listed AD attributes under the field keys. A backend that cannot do this answers `ok: false` and
reports `features.adLookup: false`.

## quit

Stops the server. → `{ ok }`
