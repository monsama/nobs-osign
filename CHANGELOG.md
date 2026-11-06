# Changelog

What changed, in plain words, for whoever runs NOBS OSign. Update with
`sudo ./scripts/update.sh <version>`. When a version says **add-in changed**, make the add-in files
again afterwards (status bar or Settings → Outlook add-in); a new `manifest.xml` is only needed
when it says so. The release notes on GitHub say the same for each version.

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
