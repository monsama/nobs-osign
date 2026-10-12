# Setting up OSign

You do this once. Afterwards, changing a signature only means editing it and clicking **Publish**.

## 1. GitHub Pages

1. In `monsama/osign`, go to **Settings → Pages → Build and deployment → Source** and pick
   **GitHub Actions**.
2. Under **Actions**, run the **pages** workflow once (*Run workflow*). Then
   <https://monsama.github.io/osign/> shows the add-in's start page and
   <https://monsama.github.io/osign/manifest.xml> serves the manifest.

From then on every push to `main` that changes `config/` deploys again. A publish from the editor is
such a push, and it is live about a minute later.

GitHub Pages from a private repository needs a paid GitHub plan. Either way, the site itself is
public.

## 2. App registration in Entra ID

The add-in reads each user's details (name, title, phone…) from Microsoft Graph. It signs in through
Outlook itself (nested app authentication), so nobody is asked for a password.

1. Entra admin center → **Identity → Applications → App registrations → New registration**.
   - Name: `OSign`
   - Supported account types: *Accounts in this organizational directory only*
2. **Authentication → Add a platform → Single-page application**. Redirect URI:
   `brk-multihub://monsama.github.io`
   (if you host elsewhere, use `brk-multihub://<that host>`).
3. **API permissions**: *Microsoft Graph → Delegated → User.Read* is there by default. Click
   **Grant admin consent for …**.
4. From **Overview**, copy the *Application (client) ID* and the *Directory (tenant) ID*. They go into
   the editor in step 4.

## 3. A GitHub token for the editor

The editor publishes by making a commit in this repository.

1. GitHub → **Settings → Developer settings → Personal access tokens → Fine-grained tokens →
   Generate new token**.
2. Repository access: **Only select repositories → monsama/osign**.
3. Permissions: **Contents: Read and write**. Nothing else is needed.
4. Give it an expiry you will remember to renew.

## 4. The editor

1. Download `OSign-vX.Y.Z.zip` from the releases and unpack it, or clone the repository.
2. Run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\OSign.ps1          # from the release
   powershell -ExecutionPolicy Bypass -File .\editor\OSign.ps1   # from the repository
   ```
   It opens in an Edge or Chrome app window. Closing that window stops it.
3. **Settings**: the target is already *GitHub, monsama/osign, main, config*. Paste the token. It is
   stored encrypted for your Windows account (DPAPI) in `%APPDATA%\OSign`.
4. **Add-in**: paste the client ID and the tenant ID from step 2.
5. **Fields**: check the mapping, especially the Graph sources (see the README).
6. **Signatures**: make them. *Import .htm…* takes over an existing classic Outlook signature from
   `%APPDATA%\Microsoft\Signatures`. Its images have to be added again with the image button.
7. **Publish**. The dialog lists anything that would go wrong first.

On a PC in the domain, the preview can show any user's real values from Active Directory: pick
*User from Active Directory…* and enter a user name or address.

Several people can use the editor, each on their own PC with their own token. **Load published**
fetches the current state before you start. If someone published in the meantime, a publish is
refused instead of overwriting their work.

## 5. Deploy the add-in

1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. App type *Office Add-in*. Provide a link to the manifest:
   `https://monsama.github.io/osign/manifest.xml`
3. Assign it to everyone, or to a group to start with.

The add-in can take a few hours to reach every Outlook. In new Outlook and on the web, signing out
and in again speeds this up. In classic Outlook, restart it.

To try it yourself first, sideload it instead: in Outlook on the web, open **Get Add-ins → My
add-ins → Add a custom add-in → Add from URL** and enter the manifest URL.

## 6. What users see

- A new message, reply or forward gets its signature by itself.
- **OSign → Signatures** in the message toolbar opens the pane. There they can insert another
  signature, or choose which one they use for new mail and for replies (if you allow that under
  *Add-in*).

## When something is off

| What | Why, and what to do |
|---|---|
| No signature at all | The add-in is not deployed to that user yet, or the app registration is missing. Open the pane: it says what it could not do. |
| Signature without name, title… | The Graph source of those fields is empty, or the user's Entra profile has no value there. |
| Two signatures | Outlook's own signature was not switched off: check *Add-in → Switch off Outlook's own signature*. In classic Outlook, the user can also remove their default under File → Options → Mail → Signatures. |
| Old signature still used | Wait a minute after publishing (GitHub Pages), then start a new message. |
| Images missing in the received mail | The recipient's mail program blocks remote images until they allow them. Keep images small and give them alt text. |
