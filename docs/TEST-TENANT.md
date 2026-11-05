# Trying NOBS OSign in a test tenant

The automatic tests (`npm test`) run the editor, the server and the add-in against stand-ins for
Active Directory, Microsoft Graph and SharePoint. What only the real ones show is on this list: go
through it once in a Microsoft 365 test tenant (a developer tenant does) before rolling out, and
again after a release whose notes say the add-in or the Microsoft 365 parts changed.

Tick each line; note what differs.

## Set up

- [ ] Server installed from the release, `update.sh` verifies it (attestation).
- [ ] Setup assistant: AD connection *Test* green; sign-in with an AD account of the sign-in group.
- [ ] Setup script (`nobs-osign-setup.ps1`) run as a Global Admin: both app registrations made,
      site permission granted, the result arrives in the editor by itself.
- [ ] *Settings → Server* Status: every line green.
- [ ] Add-in files created and uploaded (Cloudflare Pages or own web server); `manifest.xml`
      uploaded in the Microsoft 365 admin center, assigned to a test group.

## Signatures

- [ ] A signature with fields, an image, a photo per person and a building block, published.
- [ ] **Outlook on the web**: new mail, reply, forward each get the right one; a field empty for the
      user leaves its line out; the image shows (also in the received mail, not as attachment).
- [ ] **New Outlook for Windows**: the same.
- [ ] **Classic Outlook (Windows)**: the same, also with Outlook just started (event before the pane
      was ever opened); a second From address switches the signature.
- [ ] **Outlook on a phone** (if *Also in Outlook for iOS and Android* is on).
- [ ] A signature limited to an AD group: only its members get it (groups synced to Entra ID).
- [ ] A language variant: a user with `preferredLanguage` en-US gets it, one with de-CH the other.
- [ ] A campaign with today's dates wins; one in the past does not.
- [ ] A banner campaign running today: its banner where the signature has `{{campaign}}` (else
      below), its image embedded, its link with `utm_campaign`; *new mail only*: not in a reply.
- [ ] Classic Outlook: the banner also when the signature goes in by itself (pane opened once).

## Publishing

- [ ] A set time a few minutes ahead: published by itself then (History shows it).
- [ ] Four eyes on: one user asks, another approves; the first cannot approve their own.
- [ ] With a publishing group set: a member of the sign-in group outside it can only ask for
      approval, and cannot change photos.

## Users

- [ ] The pane: picking another signature for new mail, an own variant with an editable field.
- [ ] *Read my details again* after changing a phone number in Entra ID.
- [ ] Values from AD: *Users → Read from AD*, publish, the add-in uses them.

## Afterwards

- [ ] Automatic backup made at its time; one downloaded and opened (*Bring it back* on a test
      server).
- [ ] The audit log shows each step above.
