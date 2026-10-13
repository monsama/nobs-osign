# Setting up OSign

You do this once. Afterwards, changing a signature only means editing it and clicking **Publish**.

Nothing of the company goes to GitHub: GitHub holds the code, GitHub Pages serves the add-in's code.
The signatures, images and photos live in the company's own SharePoint.

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

Another site or folder: build the add-in with `OSIGN_SP_SITE` (`<host>.sharepoint.com:/sites/<name>:`)
and `OSIGN_SP_FOLDER` (step 3).

## 2. App registration in Entra ID

The add-in signs in through Outlook itself (nested app authentication), so nobody is asked for a
password, and reads with each user's own rights.

1. Entra admin center → **Identity → Applications → App registrations → New registration**.
   - Name: `OSign`
   - Supported account types: *Accounts in any organizational directory* (then the build needs no
     tenant ID). *This organizational directory only* works too; then set `OSIGN_TENANT_ID` in step 3.
2. **Authentication → Add a platform → Single-page application**. Redirect URI:
   `brk-multihub://monsama.github.io` (the host the add-in is served from).
3. **API permissions**, *Microsoft Graph → Delegated*:
   - `User.Read` (there by default): the user's name, title, phone…
   - `Sites.Read.All`: the OSign folder
   - `GroupMember.Read.All`: only if signatures go to AD groups (synced to Entra ID; nested groups count)

   Then **Grant admin consent for …**.
4. Copy the **Application (client) ID** from Overview, for step 3.

## 3. Build and serve the add-in (GitHub Pages)

1. In the repository: **Settings → Secrets and variables → Actions → Variables → New repository
   variable**: `OSIGN_CLIENT_ID` = the Application (client) ID. Optionally `OSIGN_TENANT_ID`,
   `OSIGN_SP_SITE`, `OSIGN_SP_FOLDER`.
2. **Settings → Pages → Build and deployment → Source: GitHub Actions.**
3. **Actions → pages → Run workflow.** Then <https://monsama.github.io/osign/manifest.xml> serves
   the manifest.

The add-in only has to be built again when its code or these variables change, never for a new
signature.

## 4. The editor

1. Download `OSign-vX.Y.Z.zip` from the releases and unpack it, or clone the repository.
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
5. **Fields**: check the Graph sources (see the README).
6. **Publish**. The dialog lists what would go wrong first. OneDrive uploads the folder; new
   messages use it a minute or two later.

On a PC in the domain, the preview can show any user's real values and groups from Active Directory:
pick *User from Active Directory…* and enter a user name or address.

Several admins can use the editor, each on their own PC with the folder synced. **Load published**
takes the current state from the folder before you start.

## 5. Deploy the add-in

1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. App type *Office Add-in*, link to the manifest: `https://monsama.github.io/osign/manifest.xml`
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
| No signature at all | Open the pane: it says what failed. Usually: no admin consent yet, the OSign folder not readable for that user, or `OSIGN_CLIENT_ID` not set when the add-in was built. |
| "No signatures are published yet" | `OSign\signatures.json` is not in SharePoint: publish, and check that OneDrive has synced it. |
| Signature without name, title… | The Graph source of those fields is empty, or the user's Entra profile has no value there. |
| No photo | No `OSign\photos\<username>.jpg` for that user, or `@@sAMAccountName` has no Graph source. |
| Group signatures missing | `GroupMember.Read.All` without admin consent, or the group not synced to Entra ID. |
| Two signatures | Check *Add-in → Switch off Outlook's own signature*. In classic Outlook the user can also remove their default under File → Options → Mail → Signatures. |
