# Setting up NOBS OSign

You do this once. Afterwards, changing a signature only means editing it and clicking **Publish**.

**The quick way:** install the server (step 1), sign in as `admin`. The **setup assistant** opens
by itself and walks through steps 2–6, saving each one. It can be opened again any time with
**Setup assistant** at the top. This page has the same steps in full.

NOBS OSign does not need GitHub to run: the server is installed from a release once, and everything
else is the company's own. The editor runs on the company's server, the add-in's code is served
from a web server of its choosing, and the signatures, images and photos live in the company's own
SharePoint. Nothing on the server is reachable from the internet.

## 1. Install the server

The editor runs as a container (Docker Compose), on the same server as NOBS Inventory and behind
its shared proxy. Admins sign in there with their Windows account (members of an AD group); the
server writes into SharePoint itself through Microsoft Graph, reads AD with a service account, and
uploads the add-in's files to Cloudflare Pages.

**Install** (on the server, with NOBS Inventory and its proxy running):

```bash
mkdir -p /opt/nobs-osign && cd /opt/nobs-osign
curl -fsSLO https://github.com/monsama/nobs-osign/releases/download/vX.Y.Z/nobs-osign-server-X.Y.Z.tar.gz
tar -xzf nobs-osign-server-X.Y.Z.tar.gz --strip-components=1
sudo OSIGN_HOST=osign.example.ch ./deploy-nobs-osign.sh
```

`deploy-nobs-osign.sh` creates `.env` and `secrets/` (the password of the local account `admin`,
shown once, and the key for the passwords set in the editor, `secrets/app_key`), starts the release
image (by its digest, exactly the image that was scanned), adds `nobs-osign.caddy` to the shared
proxy's sites (with its default certificate; your own with NOBS Inventory's
`scripts/install-cert.sh nobs-osign <bundle.pem>`), and a daily backup at 03:25. Then open
`https://osign.example.ch` and sign in as `admin`.

| Task | Command |
|---|---|
| Update to the latest release | `sudo ./scripts/update.sh` (backs up first; switches back if the new one is not healthy) |
| Update to a release | `sudo ./scripts/update.sh 0.7.2` |
| Back up | `sudo ./scripts/backup.sh` → `backup/nobs-osign-<time>.tar.gz` (the draft, its history, the settings and secrets; with `BACKUP_PASSPHRASE_FILE` in `.env` encrypted) |
| Restore | `sudo ./scripts/restore.sh backup/<file>` (`RESTORE_SECRETS=yes` restores `.env` and the secrets too) |
| Logs | `docker logs -f nobs-osign` |

**Updates are checked.** `update.sh` checks the download's checksum, and, where the GitHub CLI is
installed and signed in (`gh auth login`), the build attestation of the server's files and of the
image: that both were built by this repository's release workflow. Without `gh` it says it checked
only the checksum; `OSIGN_REQUIRE_ATTESTATION=1` in `.env` makes that stop the update instead.

**Backups hold the secrets.** The backup contains `secrets/` (with `app_key`) and the data volume
(with the passwords set in the editor, sealed with that key): set `BACKUP_PASSPHRASE_FILE` in `.env`
so backups are encrypted, and keep them as private as the server.

## 2. The server's connections

**Settings → Server connections** (also the setup assistant's first step), for the local account
`admin` only: whoever sets the domain controller's address gets the passwords people sign in with.
Each tab has **Save and test**.

- **Active Directory:**
  1. **Domain controller**: `ldaps://dc01.example.ch` (`ldap://` only with StartTLS), and, when its
     certificate comes from your own CA, that **CA certificate** (.cer). **Detect** reads the
     domain from the domain controller.
  2. **Domain**: its DNS name (`example.ch`). People then sign in with their Windows login, as on
     a PC (`anna`, signed in as `anna@example.ch`); users are looked for in the whole domain.
  3. **Sign-in group**: the name of an AD group (`NOBS-OSign-Admins`); its members may use the
     editor, nested groups count.
  4. **Service account**: the login of an account that reads the directory (`svc-osign`; a plain
     domain user is enough) and its password: for the preview, the group list and the users'
     values. Optionally the time of the daily export.

  **Advanced** is only for a directory that differs from the usual: another sign-in name format
  (`DOMAIN\{login}`), an OU to let in only the users below it, another login attribute, another
  place to read users and groups. Empty fields there follow from the domain (shown in grey).
- **SharePoint:** **Download the setup script** right there (step 4 describes it); its result comes
  back by itself. *Enter by hand* is for the rare case without it: the tenant ID, the server's
  client ID, and the server's certificate to upload to its app (see *By hand* in step 4). An ID
  of all zeros (an example value left in `.env`) counts as not set. A client secret in
  `secrets/graph_client_secret` still works, for installs that have one.
- **Cloudflare Pages** (before SharePoint: its address is the add-in's), with the three steps
  shown in the tab: the **account ID** (dashboard, *Account home*, at the right), an **API token**
  (*My Profile → API Tokens → Create Token → Create Custom Token*, *Account → Cloudflare Pages →
  Edit* for your account), and a **project name** (`firma-osign`; made on the first upload). The
  add-in's address then is `https://<project>.pages.dev/`, set by itself unless one is set already.

What is saved there wins over `.env`; a field left empty takes the value from `.env` (the variables
at the end of this page, the secrets as files in `secrets/`). The passwords and the token set there
are kept encrypted in the data volume and never shown again. The service account's password in
`secrets/ldap_service_password` only goes to the domain controller `.env` names (`OSIGN_LDAP_URL`).
A domain controller set in the editor has to use an LDAP port (389, 636, 3268, 3269; others with
`OSIGN_LDAP_PORTS` in `.env`).

## 3. The OSign folder in SharePoint

The signatures, their images, the photos and (with values from AD) the users' values live in a
folder **`OSign`** in a SharePoint site. The server writes into it, the add-in reads it with each
user's own sign-in.

1. **Make it:** in a site's document library (*Documents*), a folder **OSign**, with a folder
   **photos** in it. A site of its own is fine (e.g. a communication site `/sites/osign`, kept out
   of search and navigation); the server is given write access to that one site only.
2. **Who may read it:** everyone who sends mail with a signature. Share the folder with *Everyone
   except external users*, **Can view**, as direct access (a sharing link is not enough: the add-in
   reads the files, it never opens a link).
3. **Photos:** one per person in `OSign/photos`, named by the username (`<sAMAccountName>.jpg`).
   Templates that show `photos/@@sAMAccountName.jpg` embed each sender's own photo; who has none
   gets the signature without one.
4. **Tell the setup assistant where it is** (step *Where the signatures are kept*): open the folder
   in SharePoint in the browser and paste the address from the address bar. A sharing link works
   too once Microsoft 365 is set up (the server asks Graph where it leads); otherwise the assistant
   opens it for you to copy the address from there. **Test** checks the folder again.

**A personal OneDrive is not the place for it.** Only its owner can give the server access (the
setup script, run by any other account, even a Global administrator, gets *403* and stops), nobody
else reads it, and it goes with that person. For a test it works only when the script runs as the
OneDrive's owner; for the rollout use a SharePoint site everyone can read.

## 4. Microsoft 365: the setup script

Two app registrations are needed: **NOBS OSign** for the add-in (it reads with each user's sign-in)
and **NOBS OSign server** for the server (it writes into the one site with its own).

**The setup script does both.** In the setup assistant's step *Register the app in Entra ID*,
**Download the setup script**, and run it on a Windows PC:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\nobs-osign-setup.ps1"
```

Signed in once in the browser as a SharePoint or Global administrator (through *Microsoft Graph
Command Line Tools*, Microsoft's own app for Graph scripts; leave *Consent on behalf of your
organization* unticked), it:

- makes *NOBS OSign*: single-page application, redirect URI `brk-multihub://<host of the add-in>`,
  delegated `User.Read`, `Sites.Read.All` (and `GroupMember.Read.All` where groups come from Entra
  ID), with admin consent;
- makes *NOBS OSign server*: application permission `Sites.Selected` with admin consent (and, when
  ticked, `User.Read.All` and `GroupMember.Read.All` for the preview of a real person from Entra
  ID), and the server's own **certificate**: the server makes the key itself and keeps it (sealed
  in its data volume); the script only puts the public half on the app. No secret to copy or to
  leak. Where the tenant refuses certificates on apps, it makes a client secret instead;
- gives *NOBS OSign server* write access to the site of the folder (step 3), and no other;
- sends the IDs (and, only without the certificate, the secret and when it runs out) straight to
  the server (with a one-time code, valid for an hour): nothing to copy. The assistant goes on by
  itself;
- takes back the consent it was given for this, for your account, and deletes itself.

It can run again: it uses the apps of its earlier run (as the server knows them, and owned by you)
and puts a new certificate on the server's app each time (client secrets of earlier runs are
removed then). The certificate holds for two years; the status bar warns 30 days before. Apps
somebody else registered under these names are never used.

### By hand

**The add-in's app** (*NOBS OSign*):

1. <https://entra.microsoft.com> → **App registrations → New registration**: name *NOBS OSign*,
   *Accounts in this organizational directory only*.
2. **Authentication → Add a platform → Single-page application**, redirect URI
   `brk-multihub://<host of the add-in>` (e.g. `brk-multihub://company-osign.pages.dev`).
3. **API permissions** → *Microsoft Graph → Delegated*: `User.Read`, `Sites.Read.All`, and
   `GroupMember.Read.All` only if signatures go to groups and the values come from Entra ID (with
   values from AD the groups come from AD too). Then **Grant admin consent**.
4. Copy the **Application (client) ID** and the **Directory (tenant) ID** into the assistant.

**The server's app** (*NOBS OSign server*):

1. **New registration**: name *NOBS OSign server*, no redirect URI.
2. **The server's certificate**: in Server connections → SharePoint → *Enter by hand*, **Download
   the certificate (.cer)** (the server makes it; its key stays on the server), then in Entra
   **Certificates & secrets → Certificates → Upload certificate**. It holds for two years; the
   editor warns 30 days before. No client secret.
3. **API permissions → Microsoft Graph → Application permissions → Sites.Selected → Grant admin
   consent**. This alone reaches no site yet.
4. **Give it the one site** (once, as a SharePoint or Global administrator). The setup assistant
   shows both requests with the right values. In Graph Explorer (<https://aka.ms/ge>), signed in:
   - `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/osign:?$select=id`
     (for the home site: `…/sites/root?$select=id`), and copy the `id`;
   - `POST https://graph.microsoft.com/v1.0/sites/<id>/permissions` with
     `{"roles":["write"],"grantedToIdentities":[{"application":{"id":"<client ID of NOBS OSign server>","displayName":"NOBS OSign server"}}]}`.
     Graph Explorer asks for consent to *Sites.FullControl.All* for this.

   Or let the setup assistant make both requests with **Sign in with Microsoft and grant**. That
   needs a third, small app registration *NOBS OSign setup*: platform **Single-page application**
   with the redirect URI the assistant shows (`https://<editor>/ui/signin.html`), the delegated
   permission `Sites.FullControl.All`, no client secret. Keep it narrow: no admin consent (the
   consent holds for your account only), *Assignment required?* = Yes with only you assigned, and
   *Enabled for users to sign-in?* = No (or delete it) once the site is granted.
5. Enter the tenant ID and the client ID in Server connections → SharePoint → *Enter by hand*, and
   **Save and test the connection**.

`Sites.ReadWrite.All` instead of `Sites.Selected` works too, but lets the server write to every site
of the tenant.

## 5. The add-in's files

The add-in is a few static files (HTML, JavaScript, icons, `manifest.xml`): only code, no data.
Outlook loads them over HTTPS, from every device where someone writes mail, so they have to be
reachable from wherever Outlook is used, with a certificate browsers trust and no sign-in in front.

**With Cloudflare Pages** (free, no credit card; set up in step 2), the address is
`https://<project>.pages.dev/` (the assistant fills it in). In the step *Create the add-in files*,
**Create** and **Upload to Cloudflare Pages**. Your own domain later: Cloudflare → *Workers & Pages
→ the project → Custom domains*; then enter the new address, create and upload the files again, and
upload the new `manifest.xml` in the admin center.

**Another web server:** enter its address, **Create**, then **Download as ZIP** and put its content
there. *Web server* in Settings chooses the caching file that goes with the files: `_headers`
(Cloudflare Pages), `staticwebapp.config.json` (Azure), `web.config` (IIS) or `.htaccess` (Apache).
It makes Outlook check for a newer version before using its copy, so updates arrive within minutes.
For another server, have it send `Cache-Control: no-cache`.

| Where | How the files get there | Address |
|---|---|---|
| **Your website's web host** (Apache, most providers) | A subdomain with HTTPS; upload by FTP/SFTP, `.htaccess` included. | e.g. `https://osign.contoso.com/` |
| **Azure Storage, static website** | Storage account → *Data management → Static website → Enabled*; upload into `$web`. | `https://<account>.z<n>.web.core.windows.net/` |
| **Azure Static Web Apps** (Free plan) | `npx @azure/static-web-apps-cli deploy <folder> --deployment-token <token> --env production` | `https://<name>.azurestaticapps.net/` |
| **Your own web server** (IIS, nginx, a container) | Copy the files into a site or virtual directory with HTTPS. | e.g. `https://intranet.contoso.com/osign/` |

The files only have to be made again for a new NOBS OSign version or when these settings change,
never for a new signature. After a new version, upload the new `manifest.xml` in the admin center
too (step 7).

## 6. Signatures and fields

1. **Signatures**: import the existing ones with **⤓** next to *Signatures* (the `.htm` files
   together with their images; CI-Sign templates in UTF-16 with `images\…` and classic Outlook ones
   work), or make new ones. Under *Who gets it*, enter the groups.
2. **Fields and where their values come from** (*Add-in → Where the users' values come from*):
   - **Entra ID**: each field by its Graph source. Fields without one (`description`, `homePhone`,
     `info`) stay empty, unless Entra Connect syncs them as directory extensions.
   - **Entra ID and Active Directory** (recommended without those extensions): fields with a Graph
     source from Entra ID, the others from AD.
   - **Active Directory**: every field from AD.

   With AD, the users' values go to `OSign/users/`: one file per person (`<mail>.json`) with their
   fields and groups. The add-in reads only its own user's file.

   **Check before you publish:** open **Users**, click **Read from AD**. You see every user with
   the values that would go out, new and changed ones marked, and who would be removed. Nothing is
   written yet. **Publish** then writes exactly what you read there, once. With an export time set
   (Server connections → Service account), the server also exports by itself every day; the
   Add-in view shows how its last run went.
3. **Preview**: sample values, all fields empty, or a real person: *User from Active Directory…*
   (with the service account), or *Person from Entra ID…* (when the setup script gave the server
   `User.Read.All`).
4. **Publish** (Ctrl+Enter). The dialog lists what would go wrong first. New messages use it a
   minute or two later. Every publish is kept under **History**, with its images, the last 50:
   *Compare* shows what changed (the HTML line by line, who gets it, fields and options), against
   the version before or the draft; *Load into the draft* brings an earlier version back.
5. **Campaigns**: under *When*, a signature gets its first and last day (e.g. December, for the
   holidays). On those days it is used instead of the usual default, for whoever it is open to;
   before and after, the usual one again, by itself. CI-Sign's text blocks with days
   (`{dats …}{date …}` in `__intern.htm`) become such campaigns when imported; text blocks
   without days go straight into the template.
6. **Four eyes** (Settings, the local `admin`): with *Publishing needs a second person's approval*
   on, *Publish…* asks for approval. A second person signed in with AD sees what would change
   and publishes it, or rejects it; a change to the draft after that withdraws the request.
7. **Dark mode**: the preview's *Dark mode* shows roughly what recipients with Outlook's dark mode
   see. Check that the logo still shows on a dark background.

## 7. Deploy the add-in

1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. App type *Office Add-in*, link to the manifest: `https://<address of the add-in>/manifest.xml`.
3. Assign it to everyone, or to a group to start with.

It can take a few hours to reach every Outlook. To try it yourself first, sideload it: in Outlook on
the web: open <https://aka.ms/olksideload>, then **My add-ins → Custom add-ins → Add a custom add-in →
Add from URL**.

**Outlook for iOS and Android** (preview): tick *Also in Outlook for iOS and Android* in the
settings, make the add-in's files again and upload the new `manifest.xml`. Phones then put the
signature in new messages, replies and forwards too (Outlook 4.2352 and later); there is no pane
there, and no change with the From address. Try it on one phone first.

## 8. What users see

- A new message, reply or forward gets its signature by itself, images included.
- **NOBS OSign → Signatures** in the message toolbar opens the pane: insert another signature, or
  choose the one for new mail and for replies (if allowed under *Add-in*).

## Running it

- **Audit log** (Settings, the local `admin`): who signed in (and who failed), published, approved,
  brought back a version, changed the connections or settings, ran the setup script or made a
  backup, with the address it came from. Never a password or a signature's content.
- **Backup** (Settings, the local `admin`): one file with the draft and its images, the settings,
  the connections with their passwords, the certificate and the history, locked with a passphrase
  of at least 12 characters. Keep it apart from the passphrase. *Bring it back* on this server or
  a new one: the passwords are sealed again with that server's key.
- **Updates**: the status bar says when a newer version is out (twice a day the server asks GitHub
  for the latest release; `OSIGN_UPDATE_CHECK=0` in `.env` switches that off). Then
  `sudo ./scripts/update.sh <version>`.

## Settings of the server

In Settings → Server connections, or as environment variables in `.env` with the secrets as files:

| Variable | |
|---|---|
| `OSIGN_HOST` | the name people type (`osign.example.ch`); only in `.env` |
| `OSIGN_GRAPH_TENANT_ID`, `OSIGN_GRAPH_CLIENT_ID` | the tenant, and the client ID of *NOBS OSign server* |
| `OSIGN_GRAPH_CLIENT_SECRET_FILE` | a file holding its client secret (only without the server's certificate) |
| `OSIGN_GRAPH_CLIENT_SECRET_EXPIRES` | when that secret runs out (`YYYY-MM-DD`) |
| `OSIGN_ADMIN_PASSWORD_FILE` | a file holding the password of the local account `admin` (only here) |
| `OSIGN_APP_KEY_FILE` | a file holding the key the passwords set in Settings are kept with (only here) |
| `OSIGN_LDAP_URL` | `ldaps://dc01.example.ch` (or `ldap://…` with `OSIGN_LDAP_STARTTLS=1`) |
| `OSIGN_AD_DOMAIN` | the domain's DNS name, `example.ch`; the two below follow from it |
| `OSIGN_LDAP_BIND`, `OSIGN_LDAP_BASE` | only when they differ: `{login}@example.ch`, `OU=Users,DC=example,DC=ch` |
| `OSIGN_LDAP_GROUP` | who may sign in: `NOBS-OSign-Admins` (its name, or a DN; nested groups count) |
| `OSIGN_LDAP_CA` | the CA certificate the domain controller's certificate comes from, if it is not a public one |
| `OSIGN_LDAP_SERVICE_DN` | the login of an account that reads the directory (`svc-osign`, or its DN): the preview, the group list and the users' values |
| `OSIGN_LDAP_SERVICE_PASSWORD_FILE` | a file holding its password |
| `OSIGN_LDAP_DIRECTORY_BASE` | where users and groups are searched; empty: the whole domain of `OSIGN_LDAP_BASE` |
| `OSIGN_AD_EXPORT_AT` | `06:00`: the users' values are exported every day at that time (`TZ`) |
| `OSIGN_CF_ACCOUNT_ID`, `OSIGN_CF_PROJECT`, `OSIGN_CF_API_TOKEN_FILE` | Cloudflare Pages |
| `OSIGN_REQUIRE_ATTESTATION` | `1`: an update whose attestation cannot be checked stops |
| `OSIGN_UPDATE_CHECK` | `0`: the server does not ask GitHub whether a newer version is out (default `1`) |
| `OSIGN_DATA` | the draft, its images, its history and the settings (`/data`) |

Without Microsoft Graph set up the server publishes into a folder of its own (`/data/publish`),
for testing only: Outlook reads from SharePoint.

## The web certificate

Out of the box the proxy serves a certificate of its own for the editor's address: the browser
warns, and the setup script cannot send its result (it then saves it as a file to upload). With a
certificate from your own CA, which the domain's PCs trust already, neither happens. Settings →
Server shows which one is served, and whether PCs trust it.

1. **Issue it** on a domain PC (as administrator), from the template *Web Server*, for the editor's
   address and its short name. `osign.inf`:

   ```
   [NewRequest]
   Subject = "CN=osign.example.ch"
   KeyLength = 2048
   Exportable = TRUE
   MachineKeySet = TRUE
   [Extensions]
   2.5.29.17 = "{text}"
   _continue_ = "dns=osign.example.ch&"
   _continue_ = "dns=osign"
   [RequestAttributes]
   CertificateTemplate = WebServer
   ```

   ```powershell
   certreq -new osign.inf osign.req
   certreq -submit -config "<CA server>\<CA name>" osign.req osign.cer
   certreq -accept osign.cer
   $c = Get-ChildItem Cert:\LocalMachine\My | Where-Object Subject -eq 'CN=osign.example.ch' | Select-Object -First 1
   Export-PfxCertificate -Cert $c -FilePath osign.pfx -Password (Read-Host -AsSecureString 'PFX password') -ChainOption BuildChain
   ```

2. **Install it** on the server, in the install folder:

   ```
   sudo ./scripts/install-cert.sh osign.pfx
   ```

   It asks for the PFX's password, checks that the key belongs to the certificate, that it holds
   now and that it is for `OSIGN_HOST`, puts it in the proxy (with its CA after it) and reloads the
   proxy; a PEM file with the key and the certificate works as well. Then delete `osign.pfx`.
   `sudo ./scripts/install-cert.sh --remove` goes back to the proxy's own. Updates keep it; before
   it runs out (the status bar warns 30 days before), issue and install the next one the same way.

## When something is off

| What | Why, and what to do |
|---|---|
| No signature at all | A notice above the message ("NOBS OSign inserted no signature; see Signatures.") gives the reason in short. Open the pane: it says what failed. Usually: no admin consent yet, the OSign folder not readable for that user, the redirect URI missing in the app registration, or no client ID when the add-in's files were made. |
| "No signatures are published yet" | `OSign/signatures.json` is not in SharePoint: publish (Settings shows where publishing goes). |
| Signature without name, title… | The Graph source of those fields is empty, or the user's Entra profile has no value there. With values from AD: the user has no file in `OSign/users/` (no mail address in AD, or not exported yet). |
| No photo | No `OSign/photos/<username>.jpg` for that user, or `@@sAMAccountName` has no Graph source. |
| Group signatures missing | `GroupMember.Read.All` without admin consent, or the group not synced to Entra ID. |
| Publishing says the server could not sign in | The certificate or client secret ran out (the status bar says so 30 days before), or the certificate was taken off the app in Entra: run the setup script again. |
| Setup script: *the server could not be reached* (SSL/TLS trust) | The PC does not trust the certificate of the editor's address (your own CA, or the proxy's). The script saves its result as `nobs-osign-setup-result.json` next to itself: **Upload the result** in Server connections → SharePoint (or the setup assistant), then delete the file. Or have the PC trust that CA. |
| Setup script: *403* on the site | Your account is not an owner of that site (for a personal OneDrive: not its owner). The script stops before it makes anything. Make your account an owner of the site, or choose a SharePoint site, and run it again. |
| Setup script: the rights could not be taken back | It tells which; take them back in Entra → Enterprise applications → *Microsoft Graph Command Line Tools* → Permissions → User consent. |
| Two signatures | Check *Add-in → Switch off Outlook's own signature*. In classic Outlook the user can also remove their default under File → Options → Mail → Signatures. |

## Security software (EDR / XDR)

The setup script is a plain PowerShell script made for one setup; it starts no other PowerShell
and runs no encoded or downloaded code. What it does that security software looks at:

- It listens on `127.0.0.1` / `::1`, on a free port, for Microsoft's sign-in to come back (at most
  5 minutes, 16 KB per request).
- It opens the browser for that sign-in, and calls Microsoft Graph and the NOBS OSign server.
- It deletes itself when it is done.

If it is blocked, look at the alert's rule. Running it from a folder the security software trusts,
or allowing it once by its hash, is cleaner than excluding a folder for good.
