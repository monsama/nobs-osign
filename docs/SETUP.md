# Setting up OSign

You do this once. Afterwards, changing a signature only means editing it and clicking **Publish**.

**The quick way:** start the editor (step 4.1–4.2). On the first start its **setup assistant** opens
and walks through steps 1–3 and 5, saving each one, with what to do in SharePoint, Entra ID, the web
host and the admin center. It can be opened again any time: **Settings → Setup assistant**. This
page has the same steps in full.

OSign does not need GitHub to run: you download the release once, and everything else is the
company's own. The add-in's code is served from your own web server; the signatures, images and
photos live in the company's own SharePoint.

## 1. The OSign folder in SharePoint

The add-in looks for its data in the **home site** of your tenant (`https://<tenant>.sharepoint.com`),
in its document library (*Documents* / *Dokumente*), folder **`OSign`**. Every Microsoft 365 tenant
has that site, so the add-in finds it without being told.

1. Open `https://<tenant>.sharepoint.com` → *Documents*, and make a folder **OSign**.
2. **Who may read it:** everyone who sends mail with a signature. Share the folder (or the library)
   with *Everyone except external users*, read only. Only the admins who publish need to edit.
3. **On the admin PC:** open the folder and choose **Sync** (or **Add shortcut to My files**).
   OneDrive then shows it in File Explorer, e.g.
   `C:\Users\<you>\<Company>\<Site> - Documents\OSign`. That path goes into the editor in step 4.
4. **Photos:** copy one photo per person into `OSign\photos`, named by the username
   (`<sAMAccountName>.jpg`), as the network share has them now. Templates that show
   `photos/@@sAMAccountName.jpg` embed each sender's own photo; who has none gets the signature
   without one.

**The setup assistant asks for the folder one of two ways** (step 1), and finds the other half
itself:

- **SharePoint link:** open the folder in SharePoint in the browser and paste the address from the
  address bar. OSign finds the folder on this PC that OneDrive syncs it to. The library has to be
  synced with *Sync* (OSign publishes through OneDrive). Sharing links (*Copy link*) usually do not
  say which folder they are; use the address bar.
- **An existing folder on this PC:** choose it with **Choose folder…**; OSign finds out where it is in
  SharePoint from OneDrive's own notes and shows it (e.g. *✓ In SharePoint: IT › Documents ›
  General › OSign*). For a folder added with *Add shortcut to My files* OneDrive does not note which
  folder it is: use the SharePoint link instead.

Your OneDrive is readable only by you: good for testing. For the rollout, take a site's
*Documents* and share the folder (read) with everyone.

Another site or folder by hand: set it in the editor (setup assistant step 1, or Settings →
*Outlook add-in*) before making the add-in's files (step 3). *SharePoint site*: paste the folder's address from
the browser or a sharing link to it (*Share → Copy link*); OSign takes the site from it, and the
folder where the address names it. Or write it as `<host>.sharepoint.com:/sites/<name>:` (also `/teams/<name>`). *Folder in
its library*: the path within *Documents*, e.g. `General/OSign` in a team's channel folder.

**For testing, your own OneDrive works too:** make the folder `OSign` in your OneDrive and paste
the OneDrive's address (`https://<tenant>-my.sharepoint.com/personal/<you>/…`); it becomes
`<tenant>-my.sharepoint.com:/personal/<you>:`. Nobody else can read it, so only you get signatures.
For the rollout, move to a site everyone can read.

## 2. App registration in Entra ID

The add-in signs in through Outlook itself (nested app authentication), so nobody is asked for a
password, and reads with each user's own rights.

1. Entra admin center → **Identity → Applications → App registrations → New registration**.
   - Name: `OSign`
   - Supported account types: *Accounts in any organizational directory* (then no tenant ID is
     needed). *This organizational directory only* works too; then enter the tenant ID in step 3.
2. **Authentication → Add a platform → Single-page application**. Redirect URI:
   `brk-multihub://<host>`, the host the add-in is served from (step 3), e.g.
   `brk-multihub://osign.contoso.com`. The editor shows it when it makes the files.
3. **API permissions**, *Microsoft Graph → Delegated*:
   - `User.Read` (there by default): the user's name, title, phone…
   - `Sites.Read.All`: the OSign folder
   - `GroupMember.Read.All`: only if signatures go to AD groups and the values come from Entra ID
     (with values from AD, the groups come from AD too)

   Then **Grant admin consent for …**.
4. Copy the **Application (client) ID** from Overview, for step 3.

## 3. Serve the add-in from your own web server

The add-in is a few static files (HTML, JavaScript, icons, `manifest.xml`): only code, no data.
Outlook loads them over HTTPS, from every device where someone writes mail.

**Pick where they go.** Any web server that serves static files over HTTPS, with a certificate
browsers trust, and without a sign-in in front of them. It has to be reachable from wherever Outlook
is used: Outlook on the web and the new Outlook load the files on the user's own device, so a
server only reachable inside the network means signatures only inside the network (or on VPN).

| Where | How the files get there | Address |
|---|---|---|
| **Cloudflare Pages** (free, no credit card) | dash.cloudflare.com → *Workers & Pages → Create → Pages → Upload assets*: drag in the folder's content, *Deploy*. Updates: *Create new deployment*. | `https://<project>.pages.dev/`, own domain under *Custom domains* |
| **Your website's web host** (Apache, most providers) | A subdomain with HTTPS; upload by FTP/SFTP, `.htaccess` included. | e.g. `https://osign.contoso.com/` |
| **Azure Storage, static website** (simplest in Azure; costs cents) | Storage account → *Data management → Static website → Enabled*. Upload the folder's content into the `$web` container (Azure Storage Explorer, or the portal's *Upload* with the `assets` folder kept). | `https://<account>.z<n>.web.core.windows.net/` (or your own domain via Azure Front Door) |
| **Azure Static Web Apps** (Free plan) | Create one with deployment source *Other*. Upload with `npx @azure/static-web-apps-cli deploy <folder> --deployment-token <token> --env production`. | `https://<name>.azurestaticapps.net/`, own domain possible |
| **Your own web server** (IIS, nginx, Apache, a container) | Copy the folder into a site or virtual directory with HTTPS. | e.g. `https://osign.contoso.com/` or `https://intranet.contoso.com/osign/` |

**Make the files.** Start the editor (step 4.1–4.2), then **Settings → Outlook add-in**:

- *Address of the add-in*: where the files will be served, e.g. `https://osign.contoso.com/`
- *Write the files to*: a folder of its own on this PC, e.g. `C:\OSign-addin`
- *Client ID* (and *Tenant ID*, if the app is for your organization only): from step 2
- *SharePoint site* and *Folder*: leave `root` and `OSign` for the home site's OSign folder

- *Web server*: the kind of host. Its caching file comes along with the files: `_headers`
  (Cloudflare Pages), `staticwebapp.config.json` (Azure), `web.config` (IIS) or `.htaccess` (Apache).
  It makes Outlook check for a newer version before using its copy, so updates arrive within
  minutes. For another server, have it send `Cache-Control: no-cache`.

Click **Make the add-in files**, then copy the folder's content to the web server. The setup
assistant's last step, **Check the address**, then confirms that the files there are the ones made
here.

The files only have to be made again for a new OSign version or when these settings change, never
for a new signature. After a new version, upload the new `manifest.xml` in the admin center too
(step 5).

## 4. The editor

1. Download `OSign-vX.Y.Z.zip` from the releases once and unpack it, e.g. to `C:\Tools\OSign`. It
   holds the editor and the add-in's files (`addin\`). From a clone of the repository, run
   `npm ci` and `npm run build` first, so the add-in's files are there.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\OSign.ps1          # from the release
   powershell -ExecutionPolicy Bypass -File .\editor\OSign.ps1   # from the repository
   ```
   It opens in an Edge or Chrome app window, in German or English as Windows is set (change it
   under Settings → Language). Closing that window stops it.
3. **Settings**: *Publish to* = the synced OSign folder from step 1. *Photos for the preview*: leave
   it empty to use `OSign\photos`, or point it at the network share the photos are on now.
4. **Signatures**: import the existing ones with **⤓** next to *Signatures*. Pick the folder, or the
   `.htm` files together with their images. CI-Sign templates (UTF-16, `images\…`) and classic
   Outlook ones (`%APPDATA%\Microsoft\Signatures`) work. Under *Who gets it*, enter the AD groups.
5. **Fields and where their values come from** (*Add-in → Where the users' values come from*):
   - **Entra ID**: each field by its Graph source. Fields without one (`description`, `homePhone`,
     `info`) stay empty, unless Entra Connect syncs them as directory extensions.
   - **Entra ID and Active Directory** (recommended without those extensions): fields with a Graph
     source from Entra ID, the others from AD.
   - **Active Directory**: every field from AD.

   With AD, the users' values go to `OSign\users\`: one file per person (`<mail>.json`) with
   their fields and groups. The add-in reads only its own user's file. Groups then come from AD
   too, so `GroupMember.Read.All` is not needed.

   **Check before you publish:** open **Users**, click **Read from AD**. You see every user with
   the values that would go out, new and changed ones marked (with the old value), and who would
   be removed. Nothing is written yet. **Publish** then writes exactly what you read there, once;
   without a read, or after the fields or groups changed, the users stay as they are.
6. **Publish**. The dialog lists what would go wrong first. OneDrive uploads the folder; new
   messages use it a minute or two later.

On a PC in the domain, the preview can show any user's real values and groups from Active Directory:
pick *User from Active Directory…* and enter a user name or address.

Several admins can use the editor, each on their own PC with the folder synced. **Load published**
takes the current state from the folder before you start.

### Keeping the AD values current

Only needed with *Entra ID and Active Directory* or *Active Directory*, and only if AD changes
(a new phone number) should reach the signatures without anyone publishing. The scheduled export
writes **without being checked** - leave it out if every change is to be looked at first under
**Users**. On a PC in the domain where the OSign folder is synced, Task Scheduler → **Create Task**:

- **General:** as the user whose OneDrive syncs the folder; *Run only when user is logged on*
  (OneDrive uploads only while that user is signed in).
- **Triggers:** daily, e.g. 12:00, or at log on.
- **Actions:** program `powershell.exe`, arguments
  `-NoProfile -ExecutionPolicy Bypass -File "C:\Tools\OSign\OSign.ps1" -ExportUsers`

It reads the settings and the published `signatures.json`, writes only the users whose values
changed, removes the files of users who are gone, and exits (code 0 when it worked).

## 5. Deploy the add-in

1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. App type *Office Add-in*, link to the manifest: `https://<address of the add-in>/manifest.xml`
   (step 3)
3. Assign it to everyone, or to a group to start with.

It can take a few hours to reach every Outlook. To try it yourself first, sideload it: in Outlook on
the web, **Get Add-ins → My add-ins → Add a custom add-in → Add from URL**.

## 6. What users see

- A new message, reply or forward gets its signature by itself, images included.
- **OSign → Signatures** in the message toolbar opens the pane: insert another signature, or choose
  the one for new mail and for replies (if allowed under *Add-in*).

## When something is off

| What | Why, and what to do |
|---|---|
| No signature at all | Open the pane: it says what failed. Usually: no admin consent yet, the OSign folder not readable for that user, the redirect URI missing in the app registration, or no client ID when the add-in's files were made. |
| "No signatures are published yet" | `OSign\signatures.json` is not in SharePoint: publish, and check that OneDrive has synced it. |
| Signature without name, title… | The Graph source of those fields is empty, or the user's Entra profile has no value there. With values from AD: the user has no file in `OSign\users\` (no mail address in AD, or not exported yet). |
| No photo | No `OSign\photos\<username>.jpg` for that user, or `@@sAMAccountName` has no Graph source. |
| Group signatures missing | `GroupMember.Read.All` without admin consent, or the group not synced to Entra ID. |
| The editor shows plain text, no layout | Its server had stopped right after the page: an OSign window of an earlier start was still open, and the browser opened the new one in it. Fixed in 0.4.3; with an older version, close every OSign window (and a leftover Edge in Task Manager), then start OSign again. |
| Two signatures | Check *Add-in → Switch off Outlook's own signature*. In classic Outlook the user can also remove their default under File → Options → Mail → Signatures. |

## Security software (EDR / XDR)

OSign is a plain PowerShell script; it starts no other PowerShell, and runs no encoded or downloaded
code. What it does that security software looks at:

- It listens on `127.0.0.1` (ports 17690–17697) for its own window; nothing from the network can
  reach it, and every request needs a random token.
- It opens Edge or Chrome as an app window, with a browser profile of its own in `%APPDATA%\OSign`
  (and turns translation off in that profile).
- With values from AD it reads users and groups from Active Directory (LDAP), as Users and Publish
  ask for.
- It reads OneDrive's list of synced libraries (`HKCU\Software\SyncEngines\Providers\OneDrive`) to
  show where a folder is in SharePoint.

If it is blocked, look at the alert's rule and the process it names. The clean way to allow it is
to sign `OSign.ps1` with the company's code-signing certificate and allow that signer (or the
file's hash) in the security software, rather than excluding a folder.

