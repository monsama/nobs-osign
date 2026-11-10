# Changelog

What changed, in plain words, for whoever runs NOBS OSign. Update with
`sudo ./scripts/update.sh <version>`. When a version says **add-in changed**, make the add-in files
again afterwards (status bar or Settings → Outlook add-in); a new `manifest.xml` is only needed
when it says so. The release notes on GitHub say the same for each version.

## 0.9.102 (add-in changed)
A full check again (bugs, security, use). Make the add-in files again and upload them once: from
then on a publish uploads the images only with the files uploaded last.
- **Images next to the add-in's files (classic Outlook):**
  - only those of signatures someone gets (their building blocks put in) and of campaigns on
    their days, a day more at the end; a signature for nobody or a block no one uses is not
    there, nor named in the public list;
  - files made but not uploaded (to look at, or as a ZIP) never go up with a publish; the publish
    says so instead;
  - one upload at a time; compared by content too, so an image replaced under the same name goes
    up again; just after midnight, not up to 15 minutes later; after a failure tried again later
    each time (up to 4 hours), and a Pages project removed in Cloudflare is not made again by it;
  - only with Cloudflare Pages (another web server would keep them stale);
  - the add-in takes them only for the same publish as the signatures it has, checks they are
    images (and at most 1 MB), loads them together with one time limit, keeps at most 2 MB of them,
    and does not ask again for a minute when they cannot be had.
- **Security:** an SVG in signature HTML is no longer put into the pane (an SVG animation could
  make a script link there); X-Forwarded-For is taken from no one until the proxy is found
  (with OSIGN_PROXY_HOST set, as docker-compose does); the images get `nosniff`; an upload error
  shows no details of the server.
- The campaign role's publish takes images however the banner names them.

## 0.9.101
- Fixed: with every signature deleted, the change could not be published ("There is no signature"
  stopped it), and "Unpublished changes" stayed. Now it can: the checks say the one thing to know
  (Outlook then gets none, from anyone), and *Publish* asks once more.
- Importing from CI-Sign: the list in parts. The photos per person first, then the templates
  ticked, the rest (copies, AutoBackups, CI-Sign's own files, text blocks) folded away below, with
  how many there are.

## 0.9.100
- **Publish checks, as best practice has it:** what a recipient would see (alt text, width,
  http://, dark mode, missing images, classic Outlook) is checked in the active signatures only,
  the ones someone gets now or on days still to come. Templates and tests nobody gets, and
  signatures whose days are over, are said in one grey line ("not checked"), with any images they
  lack.
- Who gets nothing for new mail: one note, not two, first in the list.
- Images missing in the draft: one of the same name in another form (the GIF made a PNG, a -2)
  is named, and *Use those in the draft* puts it in their place.
- "The users stay as published" is no longer a note (nothing to do): in the grey line.
- The notes come in the order they matter: who gets nothing, images missing, classic Outlook,
  then how it looks.

## 0.9.99
- A banner (or a signature with days of its own) goes next to the add-in's files on its first day
  and away after its last: a campaign published early cannot be seen there before it starts. The
  server looks every quarter of an hour, and tries again then when an upload failed.
- Setup: Cloudflare Pages marked as recommended, with why (sharp images in classic Outlook); the
  other web servers say what classic Outlook gets with them.

## 0.9.98 (add-in changed)
- **Classic Outlook: images sharp, also when the signature goes in by itself.** With the add-in's
  files on Cloudflare Pages, each publish also puts the signatures' and banners' images (not
  people's photos) next to them; classic Outlook loads them from there, as they are, and embeds them
  in the mail (the recipient sees them at once). The 30 KB in the mailbox are then for the text
  only; the smaller copies there are for when there is no network. The preview says so.
- Make the add-in files again once (they now know their own address); from then on every publish
  uploads the images by itself. A publish that cannot reach Cloudflare still publishes, and says so.
- Images no longer published go: from Cloudflare Pages with the next publish (only what is
  published is there), and from what each Outlook keeps of them.

## 0.9.97
- Housekeeping: the editor's image work in a file of its own (`10-images.js`: which images are
  used, made lighter, put in another's place), classic Outlook's room and the ways to fit it in
  another (`10-classic.js`); what was there twice or more (an image loaded and drawn, a file as
  base64, an image put in another's place) once.
- The actions in the preview's notes (*Prepare for classic Outlook*, *Add the photo*, *Open*, …)
  are buttons, not links: reached and named as buttons by a keyboard or a screen reader.
- Fixed: optimizing images while a building block was open left the old image names in it.

## 0.9.96
- The client's address (for the sign-in throttle and the audit log) is taken from X-Forwarded-For
  only when the request comes from the shared proxy (`nobs-proxy`, or `OSIGN_PROXY_HOST`); another
  container on its network can no longer give itself another address.
- Where the add-in is deployed: to whom only for the local admin; others see whether it is, which
  version, and when it was checked.

## 0.9.95 (add-in changed)
- **Images small enough to fit**, a further option for classic Outlook when preparing is not
  enough (in the note in the preview, and in the message after preparing): the images of the
  default signatures made as small in KB as it takes for all of it to fit (narrower, a JPEG at a
  lower quality, a PNG with fewer colours). They look less sharp, in every mail; it says when even
  that is not enough (then fewer images).
- Preparing makes smaller still only the images of the signatures that do not fit (not every
  image); an image without a width in the HTML keeps its size; a banner made of HTML is taken
  along. While the images are worked on, the editing waits and the message stays.
- From the publish checks, *Prepare for classic Outlook* closes the dialog, so that what it says
  (and offers next) can be seen. One name for it everywhere.
- Fixed: text typed in the HTML tab was written into a draft loaded afterwards (load published,
  history, a newer version), and ↶ did not undo it. Undo after preparing no longer brings back
  images that are gone. A size measured for classic Outlook was kept for a new file of the same
  name. Ctrl+S says "saved" only once it is.
- Texts: the server's messages in German, unused ones removed.
- **Add-in:** with SharePoint or the network gone, the logos come from the copy it kept (they
  were left out); a stalled answer no longer keeps a new mail waiting; a failed sign-in is not
  tried again for every mail; classic Outlook's copy measured as Outlook stores it, and kept
  without the largest image rather than not at all; the choice for new mail and replies goes into
  classic Outlook's copy at once, and says when Outlook did not save it; a banner in the user's
  language no longer takes room there with the others; the date is the local one; the add-in's
  name and description in de-CH, French and Italian too. Signature HTML shown in the pane for
  editing is made safe; a value from the directory goes into a link only as http(s), mailto or tel.
- **Server:** what publishing goes out with is checked against the signatures as shown, element by
  element (links swapped between two, or text hidden in a comment for Outlook only, are refused);
  a user disabled in AD or taken out of the sign-in group is signed out within 5 minutes; the
  campaign role cannot publish over an image the signatures use; at most 500 images and 200 MB in
  the draft; sign-in requests read 16 KB at most; unexpected errors no longer show their details.
- `update.sh`: the image's attestation checked as made by the release workflow; a release without
  its image digest refused when attestation is required.

## 0.9.94
- **One way to get a signature ready for classic Outlook:** *Prepare for classic Outlook*, in the
  note in the preview (and *Optimize for Outlook* in the publish checks). It makes the HTML
  lighter, optimizes the images (as small as shown, a still GIF as a PNG, a PNG with a palette, a
  photo as a JPEG), takes them smaller still when that is not enough, and says whether it fits now
  or which image still does not. The separate buttons (*Optimize* in the image bar, *Optimize for
  the web* and *Smaller still* in the image dialog, *Make lighter*) are gone.
- GIFs are optimized too (icons were left as they were).

## 0.9.93
- **Images: really smaller files.** A PNG (a logo, an icon) is saved with a palette, at most 256
  colours, as pngquant or TinyPNG do it, and keeps its transparency: *Optimize* only when nothing
  that shows changes, *Smaller still* (Shift+click) with the 256 most used colours. A logo of 7 KB
  is about 3 KB so.
- *Optimize* on one image says what it gained, or why there was nothing to gain.
- *Make lighter* is in the bar with *Design / HTML*, always in view.

## 0.9.92
- *Make lighter* in the HTML tab (and from the note on classic Outlook): what Word, Outlook and
  browsers leave in a signature and shows nothing (mso-… and other leftover styles, `lang`, empty
  spans, runs of spaces) taken out. Often the text, not the images, is what fills classic
  Outlook's room.
- The note on classic Outlook says what the room goes to: text, photo, images.
- An image optimized again keeps its name (no more -2-2-2).
- Fixed: optimizing images, or the HTML, with the HTML tab open put the old text back.

## 0.9.91 (add-in changed)
- *Optimize* in the image bar: the image clicked, for the web (Shift+click: smaller still). The
  note on classic Outlook in the preview and the publish check offer the same for all images.
- **Classic Outlook: what does not fit is made smaller for it**, step by step (narrower, a JPEG at
  a lower quality), only in the copy in the mailbox: every other Outlook, and *Insert* in the pane,
  gets the images as they are. The preview says when this happens, and warns only when even that
  is not enough.
- The note on classic Outlook says which images it goes without there (the add-in keeps the text
  and the smallest images first, as many as fit), not "without its images".

## 0.9.90
- *Optimize for the web* in the image dialog (and in the publish checks): every image shown as
  small as it is shown (twice that, for sharp screens), PNGs without transparency as JPEG; the
  originals no longer used go from the draft. *Smaller still*: as wide as shown and more
  compressed, for the lightest mails and the most room in classic Outlook.
- The preview shows a photo or placeholder changed under Photos at once (it was kept for five
  minutes by the browser).
- The bars above the design (image, link, table cell): every control on one middle line.

## 0.9.89
- Housekeeping: the stylesheet tidied (rules for one selector merged, the cascade kept, checked
  screen by screen in light and dark), unused rules, texts and code removed.
- Campaigns: the form split into its parts; typing redraws once it pauses, not at every key.
- The *Add-in* page in cards, like the campaigns.
- This changelog, and *Known limits* in the README.

## 0.9.88
- The campaign form laid out anew: a head with the name and its state, cards for when, banner,
  where and who, the preview beside. Fixed: a stray "null", and the page stretching below the
  status bar.
- Classic Outlook's room measured as the add-in makes the images; what takes it is named;
  *Save PNG photos as JPEG* for PNGs without transparency.

## 0.9.87 (add-in changed)
- **The campaign role**: an AD group (Settings → Server → Active Directory → Campaign group) whose
  members make and publish the campaigns only. The server keeps it so.
- *On a day → For a person*: one person's signatures and banners on a day.
- Publishing checks banners: alt text, dark mode, size.
- `docs/CAMPAIGNS.md`, a short guide for marketing.
- Image names that are paths once decoded are left out.

## 0.9.86 (add-in changed)
- Campaigns checked before publishing; banners in turns, on some weekdays; a new one from an
  earlier one; the ones over folded away.
- Classic Outlook's 30 KB said in the preview and before publishing; unused images removed.
- Fields and building blocks dragged in the design; four column layouts.
- The pane says when a banner goes in with a signature.

## 0.9.85 (add-in changed)
- **Banner campaigns**: a banner for some days put into the signatures there are (in
  `{{campaign}}` or below), made without HTML, in other languages, for some signatures or groups,
  in new mail only, with link tracking. A timeline, the preview on a day, a note when one ends with
  none after it (status bar and webhook), *Copy for next year*.

## 0.9.84 (add-in changed)
- House colours, spacing, a bar for a link and for a table cell, an image's corners and place.
- The preview with one field empty, and at phone width.

## 0.9.83
- Images sized by dragging their corner; *Original size*.

## 0.9.82
- A full security check (who may call what, paths, headers); the draft versioned (a save over a
  newer one asks).

## 0.9.67 - 0.9.81
- Full-width pages, tooltips, the setup assistant's steps to jump between, notifications to a
  webhook, replace in all signatures, the phone number format, the add-in in French and Italian,
  undo for the whole draft (↶ ↷), ⟳ without reloading, sessions list, CSV of users and the audit
  log, and many smaller fixes.
