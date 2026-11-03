# Running NOBS OSign: the admin's handbook

For whoever looks after NOBS OSign once it is set up ([SETUP.md](SETUP.md) is the first time). Each
part is short on purpose: what to do, where in the editor, and what to look at when it goes wrong.

## Every day: changing signatures

1. Change the signature in the editor. The draft is saved by itself; nothing reaches Outlook yet.
2. Look at it: the preview with *Sample values*, *All fields empty* (which lines go) and a real
   person (*User from Active Directory…*). **▦** next to *Signatures* shows all of them side by side.
3. **Publish…** lists what to look at first: images too large or without a text (alt), wider than
   600 pixels, `http://` links, unknown placeholders or building blocks, people who would get none.
4. Publish. Outlook uses it a minute or two later (the add-in reads `signatures.json` again with the
   next message).

Parts many signatures share (a disclaimer, a banner, the address) are **building blocks**
(*Building blocks* in the sidebar), put in as `{{name}}`: change them there, once.

## Releasing a change carefully

- **A pilot group first:** *Publish… → First only for a pilot group*, with an AD group of a few
  people. They get the new version, everyone else keeps the one before. When it is fine:
  *Publish… → Give it to everyone*; if not, *Stop the pilot*. Publishing normally (for everyone)
  also ends a pilot.
- **At a set time:** *Publish… → Later, at*: the draft as it is now goes out then (also for a pilot
  group). Changes made afterwards wait for the next publish. Cancel it in the same dialog.
- **Four eyes** (*Settings → Server → Four eyes*): one person asks for approval, another one
  publishes. A pilot or a set time is not offered then.
- **Back to a version before:** *History* → *Load into the draft* → publish.

## Who may do what

| | Members of the sign-in group | … also in the publishing group | Local account `admin` |
|---|---|---|---|
| Edit signatures, fields, building blocks | ✓ | ✓ | ✓ |
| Publish, pilot, set time, write users, change photos | asks for approval | ✓ | ✓ |
| Approve someone else's publish | – | ✓ | ✓ |
| Server connections, add-in files, backups, audit log, four eyes | – | – | ✓ |

The **publishing group** is optional (*Settings → Connections → Active Directory → Publishing
group*); empty, everyone in the sign-in group publishes. Membership is read when someone signs in:
taken out of the group, they keep their rights until they sign in again (at most 12 hours).

## Backups

*Settings → Server → Backup*:
- **Download a backup** now, locked with a passphrase (12 characters at least). Keep the file and
  the passphrase apart.
- **Automatically, every day:** a time, how many days to keep them, and a passphrase. The server
  keeps them in `backups/` of the data volume, or in `OSIGN_BACKUP_DIR` (another disk is better:
  set it in `.env`, outside the data volume). The newest always stays. *Download* fetches one.
- **Bring a backup back:** the file and its passphrase. It replaces the draft, the settings, the
  connections and the history; the audit log and the automatic backups stay.

A backup holds the server's key, so a new server opens the passwords in it. Without the passphrase
nobody opens a backup, not even with the server.

## Updates

```bash
sudo ./scripts/update.sh X.Y.Z
```

It checks the release's signature (attestation) before it installs. The editor says when a newer
version is out (status bar). When the add-in changed (the release notes say so): *Settings →
Outlook add-in → Create the add-in files* (and upload them); a new `manifest.xml` in the Microsoft
365 admin center only when the notes say so.

## When something is off

| What you see | Where to look |
|---|---|
| Someone gets no signature | *Users* → *Only without a signature* (with values from AD); *Who gets it* of the signature; the add-in's pane shows what the person is offered |
| A field is empty in Outlook | *Fields*: a Graph source (Entra ID) or an AD attribute; in AD mode, *Users* read and published |
| A photo is missing | *Photos*: named after the person's field (`mustera.jpg`); otherwise the placeholder photo |
| Publishing fails | *Settings → Server* (Status): SharePoint, the certificate, the connections; the audit log for the error |
| An old signature in Outlook | the add-in's pane (*OSign*) shows the version and when it was published; *Read my details again* |
| The pilot group still has the old one | the group's name as in AD, and whether the add-in can read groups (`GroupMember.Read.All`) |

The **audit log** (*Settings → Server*, local `admin`) has every sign-in, publish, approval and
change of the connections, with who and from where.

## Before a big change: try it in a test tenant

[TEST-TENANT.md](TEST-TENANT.md) is a list to go through once in a Microsoft 365 test tenant: the
parts the automatic tests stand in for (Graph, SharePoint, Outlook itself).
