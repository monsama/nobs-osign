/*
 * The editor's languages. The page is written in English; this replaces its texts with the
 * German ones when German is chosen (Settings → Language), or, until something is chosen, when the
 * browser's own language is German.
 *
 *   t('Text', a, b)   a text from the code, {0} {1} … filled in
 *   I18N.apply(root)  the page's own texts: every text node and placeholder/title/aria-label found
 *                     in DE.text, and the elements marked data-th="key" from DE.html
 */
'use strict';
(function () {
  var DE = { text: {}, html: {} };

  DE.text = {
    // top bar, sidebar
    'Load published': 'Veröffentlichte laden',
    'Settings': 'Einstellungen',
    'Publish…': 'Veröffentlichen…',
    'Signatures': 'Signaturen',
    'Setup': 'Einrichtung',
    'Fields': 'Felder',
    'placeholders': 'Platzhalter',
    'Add-in': 'Add-in',
    'behaviour & sign-in': 'Verhalten & Anmeldung',
    'Replace the draft with what is published': 'Den Entwurf durch das Veröffentlichte ersetzen',
    'Import templates (.htm, e.g. from CI-Sign or classic Outlook)': 'Vorlagen importieren (.htm, z. B. aus CI-Sign oder dem klassischen Outlook)',
    'Import templates': 'Vorlagen importieren',
    'New signature': 'Neue Signatur',
    'None yet': 'Noch keine',
    'New': 'Neu',
    'Reply': 'Antwort',
    'Limited': 'Eingeschränkt',
    // a signature
    'Signature name': 'Name der Signatur',
    'Duplicate': 'Duplizieren',
    'Delete': 'Löschen',
    'Default for new mail': 'Standard für neue Mails',
    'Default for replies & forwards': 'Standard für Antworten & Weiterleitungen',
    'Leave out empty lines': 'Leere Zeilen weglassen',
    'A line whose fields are all empty for a user is left out: no "Mobile:" without a number': 'Eine Zeile, deren Felder bei jemandem alle leer sind, fällt weg: kein «Mobile:» ohne Nummer',
    'Who gets it:': 'Wer sie bekommt:',
    'everyone': 'alle',
    'AD groups': 'AD-Gruppen',
    '(any of them; nested groups count)': '(eine davon genügt; verschachtelte Gruppen zählen)',
    'Departments': 'Abteilungen',
    '(comma separated)': '(mit Komma getrennt)',
    'Sender domains': 'Absender-Domänen',
    '(the From address)': '(die Absenderadresse)',
    'Empty means everyone. When several are filled in, a user has to match each of them.': 'Leer heisst: alle. Ist mehr als eines ausgefüllt, muss jemand alle erfüllen.',
    'Design': 'Gestaltung',
    'HTML': 'HTML',
    'Image': 'Bild',
    'Width': 'Breite',
    'Alt text': 'Alternativtext',
    'Link': 'Link',
    'Done': 'Fertig',
    'Preview': 'Vorschau',
    'Values for the preview': 'Werte für die Vorschau',
    'Sample values': 'Beispielwerte',
    'All fields empty': 'Alle Felder leer',
    'User from Active Directory…': 'Benutzer aus dem Active Directory…',
    'User from Active Directory (not reachable)': 'Benutzer aus dem Active Directory (nicht erreichbar)',
    'user name or mail': 'Benutzername oder Mail',
    'Look up': 'Suchen',
    'Plain text': 'Nur Text',
    'To:': 'An:',
    'Subject:': 'Betreff:',
    'Signature editor': 'Signatur-Editor',
    // toolbar
    'Font': 'Schrift',
    'Size': 'Grösse',
    'Text colour': 'Textfarbe',
    'Insert a field': 'Feld einfügen',
    '@@ Insert field': '@@ Feld einfügen',
    'Bold (Ctrl+B)': 'Fett (Ctrl+B)',
    'Italic (Ctrl+I)': 'Kursiv (Ctrl+I)',
    'Underline (Ctrl+U)': 'Unterstrichen (Ctrl+U)',
    'Align left': 'Linksbündig',
    'Centre': 'Zentriert',
    'Align right': 'Rechtsbündig',
    'Remove link': 'Link entfernen',
    'Two columns (logo beside the text)': 'Zwei Spalten (Logo neben dem Text)',
    'Line': 'Linie',
    'Clear formatting': 'Formatierung entfernen',
    'Undo (Ctrl+Z)': 'Rückgängig (Ctrl+Z)',
    'Redo (Ctrl+Y)': 'Wiederholen (Ctrl+Y)',
    // fields
    'Field': 'Feld',
    'Placeholder': 'Platzhalter',
    'Placeholder name': 'Name des Platzhalters',
    'AD attribute': 'AD-Attribut',
    'Graph source': 'Graph-Quelle',
    'Sample value': 'Beispielwert',
    '+ Add field': '+ Feld hinzufügen',
    'Remove field': 'Feld entfernen',
    'not available': 'nicht vorhanden',
    'New field': 'Neues Feld',
    // add-in options
    'Behaviour': 'Verhalten',
    "Switch off Outlook's own signature in the messages the add-in signs, so no mail gets two": 'Die eigene Signatur von Outlook in den Mails ausschalten, die das Add-in signiert, damit keine Mail zwei bekommt',
    "Users may choose which signature they use for new mail and for replies (in the add-in's pane)": 'Benutzer dürfen wählen, welche Signatur sie für neue Mails und für Antworten nehmen (im Bereich des Add-ins)',
    'Sign new meetings too': 'Auch neue Besprechungen signieren',
    "Insert a signature even when the user's details could not be read (the fields stay empty)": 'Auch dann eine Signatur einfügen, wenn die Angaben der Person nicht gelesen werden konnten (die Felder bleiben leer)',
    'Sign-in and permissions': 'Anmeldung und Berechtigungen',
    'How to register it': 'So wird sie registriert',
    'Install': 'Installation',
    'Register the add-in in Entra ID': 'Das Add-in in Entra ID registrieren',
    'No signatures yet': 'Noch keine Signaturen',
    'Start with a new one, or load what is published.': 'Mit einer neuen beginnen, oder das Veröffentlichte laden.',
    // settings
    'Language': 'Sprache',
    'Publish to': 'Veröffentlichen nach',
    '(the OSign folder in SharePoint, synced to this PC with OneDrive)': '(der OSign-Ordner in SharePoint, mit OneDrive auf diesen PC synchronisiert)',
    'Photos for the preview': 'Fotos für die Vorschau',
    '(one <username>.jpg per person; empty: the photos folder in the folder above)': '(ein <Benutzername>.jpg pro Person; leer: der Ordner photos im Ordner oben)',
    'Cancel': 'Abbrechen',
    'Save': 'Speichern',
    'Not set yet.': 'Noch nicht festgelegt.',
    'Found.': 'Gefunden.',
    'Not found on this PC: is the library synced?': 'Auf diesem PC nicht gefunden: ist die Bibliothek synchronisiert?',
    'No photos folder there (yet): the preview shows no photos.': 'Dort ist (noch) kein Ordner photos: die Vorschau zeigt keine Fotos.',
    'Settings saved.': 'Einstellungen gespeichert.',
    'Saved, but the folder is not there on this PC.': 'Gespeichert, aber den Ordner gibt es auf diesem PC nicht.',
    // publish
    'Publish': 'Veröffentlichen',
    'What changed': 'Was geändert hat',
    'Update signatures': 'Signaturen aktualisiert',
    'Close': 'Schliessen',
    'Publishing…': 'Wird veröffentlicht…',
    'Nothing had changed.': 'Es hatte sich nichts geändert.',
    'Published to {0}.': 'Veröffentlicht nach {0}.',
    'To {0}. OneDrive uploads it to SharePoint; new messages use it a minute or two later.': 'Nach {0}. OneDrive lädt es in SharePoint hoch; neue Mails verwenden es ein bis zwei Minuten später.',
    'No folder set yet (Settings).': 'Noch kein Ordner festgelegt (Einstellungen).',
    'There is no signature.': 'Es gibt keine Signatur.',
    'No signature is the default for new mail: users get none unless they pick one.': 'Keine Signatur ist Standard für neue Mails: niemand bekommt eine, ausser man wählt selbst eine.',
    'The default for new mail is limited to some users: everyone else gets none.': 'Der Standard für neue Mails gilt nur für einige: alle anderen bekommen keine.',
    'No signature is the default for replies.': 'Keine Signatur ist Standard für Antworten.',
    'No folder to publish to: set the synced OSign folder in Settings.': 'Kein Ordner zum Veröffentlichen: den synchronisierten OSign-Ordner in den Einstellungen festlegen.',
    'Signatures limited to AD groups: the app registration needs GroupMember.Read.All (delegated, with admin consent), and the groups have to be synced to Entra ID.': 'Signaturen nur für AD-Gruppen: die App-Registrierung braucht GroupMember.Read.All (delegiert, mit Administratorzustimmung), und die Gruppen müssen nach Entra ID synchronisiert sein.',
    'Photos per person ({0}) come from photos\\<username>.jpg in the OSign folder; who has none gets the signature without.': 'Fotos pro Person ({0}) kommen aus photos\\<Benutzername>.jpg im OSign-Ordner; wer keines hat, bekommt die Signatur ohne.',
    'Images that are no file of the signatures are left out in: {0}.': 'Bilder, die keine Datei der Signaturen sind, fallen weg in: {0}.',
    'Not fields, left as typed: {0}': 'Keine Felder, bleiben wie geschrieben: {0}',
    'Empty in Outlook (no Graph source): {0}': 'Leer in Outlook (keine Graph-Quelle): {0}',
    'Images not in the draft, left out until added: {0}': 'Bilder nicht im Entwurf, fallen weg, bis sie hinzugefügt sind: {0}',
    '{0} signature(s), ready.': '{0} Signatur(en), bereit.',
    // status
    'Saving…': 'Wird gespeichert…',
    'Not saved': 'Nicht gespeichert',
    'Unpublished changes': 'Unveröffentlichte Änderungen',
    'Published': 'Veröffentlicht',
    'Saved': 'Gespeichert',
    'Last published {0}': 'Zuletzt veröffentlicht {0}',
    // preview notes
    'Not a field: {0} - left as typed.': 'Kein Feld: {0} – bleibt wie geschrieben.',
    'No Graph source, so empty in Outlook: {0}. See Fields.': 'Keine Graph-Quelle, darum in Outlook leer: {0}. Siehe Felder.',
    'Values of {0} from Active Directory.': 'Werte von {0} aus dem Active Directory.',
    'This signature is open to them.': 'Diese Signatur steht ihnen zur Verfügung.',
    'This signature is NOT open to them.': 'Diese Signatur steht ihnen NICHT zur Verfügung.',
    'New mail: {0}, replies: {1}.': 'Neue Mails: {0}, Antworten: {1}.',
    'none': 'keine',
    'Hello,': 'Guten Tag,',
    'this is how the signature looks under a message.': 'so sieht die Signatur unter einer Mail aus.',
    'Active Directory is not reachable from here: this needs Windows on a machine in the domain.': 'Das Active Directory ist von hier aus nicht erreichbar: dafür braucht es Windows auf einem PC in der Domäne.',
    // images
    'PNG, JPG or GIF, at most 1 MB. It is published with the signatures.': 'PNG, JPG oder GIF, höchstens 1 MB. Es wird mit den Signaturen veröffentlicht.',
    'Upload an image…': 'Bild hochladen…',
    'No images yet.': 'Noch keine Bilder.',
    'remove': 'entfernen',
    'Remove from the draft': 'Aus dem Entwurf entfernen',
    'The image is larger than 1 MB.': 'Das Bild ist grösser als 1 MB.',
    '{0} is used in: {1}. Remove it anyway?': '{0} wird verwendet in: {1}. Trotzdem entfernen?',
    'Link address (https://…, mailto:…, tel:…). Fields work here too: mailto:@@mail': 'Linkadresse (https://…, mailto:…, tel:…). Felder gehen hier auch: mailto:@@mail',
    'A link starts with https://, mailto: or tel:.': 'Ein Link beginnt mit https://, mailto: oder tel:.',
    // import
    'Choose files…': 'Dateien wählen…',
    'Choose a folder…': 'Ordner wählen…',
    'Import': 'Importieren',
    'CI-Sign text blocks ({##…}), not a signature. ': 'CI-Sign-Textbausteine ({##…}), keine Signatur. ',
    '{0} fields': '{0} Felder',
    'no fields': 'keine Felder',
    ' · {0} image(s)': ' · {0} Bild(er)',
    ' · not picked: {0}': ' · nicht gewählt: {0}',
    ' · photo per person {0} → photos\\{1}': ' · Foto pro Person {0} → photos\\{1}',
    ' · replaces the signature of the same name': ' · ersetzt die gleichnamige Signatur',
    '{0} template file(s) found, {1} image(s) picked.': '{0} Vorlagendatei(en) gefunden, {1} Bild(er) gewählt.',
    'No .htm files among what was picked.': 'Unter dem Gewählten sind keine .htm-Dateien.',
    '{0} template(s) imported.': '{0} Vorlage(n) importiert.',
    '{0} template(s) imported, with {1} image(s).': '{0} Vorlage(n) importiert, mit {1} Bild(ern).',
    // dialogs and messages
    'Delete the signature "{0}"? It goes from Outlook with the next publish.': 'Die Signatur «{0}» löschen? Sie verschwindet beim nächsten Veröffentlichen aus Outlook.',
    'A placeholder name is letters, digits and _, starting with a letter.': 'Ein Platzhaltername besteht aus Buchstaben, Ziffern und _ und beginnt mit einem Buchstaben.',
    '@@{0} is already a field.': '@@{0} ist schon ein Feld.',
    'Change @@{0} to @@{1} in {2} signature(s) too?': '@@{0} auch in {2} Signatur(en) zu @@{1} ändern?',
    '@@{0} is used in: {1}. Remove the field anyway? The placeholder then stays as typed.': '@@{0} wird verwendet in: {1}. Das Feld trotzdem entfernen? Der Platzhalter bleibt dann wie geschrieben.',
    'Replace the draft with what is published? Changes not published are lost (the draft is kept once as draft.json.bak).': 'Den Entwurf durch das Veröffentlichte ersetzen? Nicht Veröffentlichtes geht verloren (der Entwurf bleibt einmal als draft.json.bak erhalten).',
    'Loaded the published signatures.': 'Die veröffentlichten Signaturen sind geladen.',
    'The saved draft could not be read; starting from an empty one.': 'Der gespeicherte Entwurf liess sich nicht lesen; es geht mit einem leeren weiter.',
    'The OSign server is not running. Start OSign.ps1 again.': 'Der OSign-Server läuft nicht. OSign.ps1 neu starten.',
    'The OSign server is not running.': 'Der OSign-Server läuft nicht.',
    'The server answered {0}.': 'Der Server antwortete {0}.',
    'This window is not signed in to the server: open the address OSign printed.': 'Dieses Fenster ist nicht beim Server angemeldet: die Adresse öffnen, die OSign angezeigt hat.',
    'Not connected': 'Nicht verbunden',
    'Failed.': 'Fehlgeschlagen.',
    'Signature': 'Signatur',
    '(no name)': '(ohne Name)',
    'Standard': 'Standard',
    '(copy)': '(Kopie)',
    'OK': 'OK',
    // messages from the server (OSign.ps1) that are always the same
    'Only PNG, JPG and GIF images can be used: Outlook does not show SVG or WebP in a signature.': 'Nur PNG-, JPG- und GIF-Bilder gehen: Outlook zeigt SVG oder WebP in einer Signatur nicht an.',
    'The image did not arrive whole.': 'Das Bild ist nicht vollständig angekommen.',
    'The image is empty.': 'Das Bild ist leer.',
    'The image is larger than 1 MB. A signature image is best well under 100 KB.': 'Das Bild ist grösser als 1 MB. Ein Signaturbild ist am besten deutlich unter 100 KB.',
    'No such image.': 'Dieses Bild gibt es nicht.',
    'Nothing has been published there yet.': 'Dort ist noch nichts veröffentlicht.',
    'Pick the folder to publish to in Settings: the OSign folder of SharePoint, synced to this PC.': 'Den Ordner zum Veröffentlichen in den Einstellungen wählen: den OSign-Ordner aus SharePoint, auf diesen PC synchronisiert.',
    'Looking up users needs Windows on a machine in the domain.': 'Benutzer suchen geht nur unter Windows auf einem PC in der Domäne.',
    'Looking up groups needs Windows on a machine in the domain.': 'Gruppen suchen geht nur unter Windows auf einem PC in der Domäne.',
    'More than one user matches that.': 'Mehr als ein Benutzer passt dazu.',
    'The request is not valid JSON.': 'Die Anfrage ist kein gültiges JSON.',
    'The configuration is not valid JSON.': 'Die Konfiguration ist kein gültiges JSON.'
  };

  DE.html = {
    fieldsLead: 'Ein Feld ist ein Platzhalter in einer Signatur, <code>@@givenName</code>, den das Add-in beim Einfügen mit dem Wert der Person füllt. Das Add-in liest die Werte aus Microsoft Graph, dem eigenen Profil der Person in Entra ID. <b>AD-Attribut</b> dient der Vorschau hier und wird aus dem Active Directory gelesen.',
    fieldsCallout: '<b>Ohne Graph-Quelle</b> bleibt ein Feld in Outlook leer. Graph kennt kein <code>description</code>, <code>homePhone</code> oder <code>info</code>. Um sie zu nutzen, sie mit Entra Connect aus dem AD synchronisieren, entweder nach <code>onPremisesExtensionAttributes.extensionAttribute1</code>…<code>15</code> (Exchange-Attribute) oder als Verzeichniserweiterungen (<code>extension_&lt;appId&gt;_description</code>), und diesen Pfad hier eintragen.',
    signinLead: 'Das Add-in meldet sich über Outlook an (ohne Passwortabfrage) und liest mit den Rechten der jeweiligen Person deren Angaben und Gruppen aus Entra ID und die Signaturen aus dem OSign-Ordner in SharePoint. Es verwendet eine App-Registrierung, die dort festgelegt wird, wo das Add-in gebaut wird – nicht hier und nicht in den Signaturen. <a href="#" id="lnkSetup">So wird sie registriert</a>',
    installLead: 'Das Manifest des Add-ins für das Microsoft 365 Admin Center (Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen) ist <code>manifest.xml</code> dort, wo das Add-in bereitgestellt wird, z. B. <code>https://&lt;owner&gt;.github.io/osign/manifest.xml</code>.',
    settingsLead: 'In SharePoint die Dokumentbibliothek der Startseite öffnen (https://&lt;tenant&gt;.sharepoint.com), einen Ordner <b>OSign</b> anlegen und <b>Synchronisieren</b> wählen (oder <b>Verknüpfung zu „Meine Dateien“ hinzufügen</b>). Er erscheint dann im Explorer; sein Pfad gehört ins erste Feld. Veröffentlichen schreibt <code>signatures.json</code> und <code>assets\\</code> dorthin; die Fotos kommen von Hand nach <code>photos\\</code>.',
    importLead: 'Einen Ordner wählen, oder die Vorlagendateien (<code>.htm</code>) zusammen mit ihren Bildern. Die Bilder, die eine Vorlage verwendet (<code>images/logo.png</code>), werden unter dem Gewählten am Dateinamen erkannt und Teil der Signaturen. Vorlagen aus CI-Sign und aus dem klassischen Outlook (<code>%APPDATA%\\Microsoft\\Signatures</code>) gehen.',
    help1: 'Entra Admin Center → <b>App-Registrierungen</b> → <b>Neue Registrierung</b>. Name: <i>OSign</i>. Konten: <i>in einem beliebigen Organisationsverzeichnis</i> (dann braucht der Build keine Mandanten-ID).',
    help2: '<b>Authentifizierung</b> → Plattform hinzufügen → <b>Single-Page-Anwendung</b>, Umleitungs-URI <code>brk-multihub://&lt;Host, von dem das Add-in kommt&gt;</code>, z. B. <code>brk-multihub://monsama.github.io</code>.',
    help3: '<b>API-Berechtigungen</b>, Microsoft Graph, delegiert: <code>User.Read</code>, <code>Sites.Read.All</code> (der OSign-Ordner) und, für Signaturen nach AD-Gruppe, <code>GroupMember.Read.All</code>. Dann <b>Administratorzustimmung erteilen</b>, damit niemand gefragt wird.',
    help4: 'Die <b>Anwendungs-ID (Client)</b> kommt in den Build: in GitHub unter Settings → Secrets and variables → Actions → Variables, <code>OSIGN_CLIENT_ID</code>; dann den Workflow <i>pages</i> ausführen.',
    help5: 'Der OSign-Ordner muss für alle lesbar sein, die signieren: in SharePoint mit allen Mitarbeitenden teilen (lesen).'
  };

  var LANGS = { en: 'English', de: 'Deutsch' };
  function chosen() { try { return localStorage.getItem('osign.lang'); } catch (e) { return null; } }
  var lang = chosen() || ((navigator.language || '').toLowerCase().indexOf('de') === 0 ? 'de' : 'en');
  if (!LANGS[lang]) lang = 'en';
  var dict = lang === 'de' ? DE : null;

  function norm(s) { return String(s).replace(/\s+/g, ' ').trim(); }

  function t(s) {
    var out = dict && Object.prototype.hasOwnProperty.call(dict.text, s) ? dict.text[s] : s;
    var args = Array.prototype.slice.call(arguments, 1);
    return out.replace(/\{(\d+)\}/g, function (all, i) { return args[i] !== undefined ? String(args[i]) : all; });
  }

  function apply(root) {
    root = root || document.body;
    if (!dict) return;
    root.querySelectorAll('[data-th]').forEach(function (el) { var h = dict.html[el.getAttribute('data-th')]; if (h) el.innerHTML = h; });
    var walker = document.createTreeWalker(root, 4 /* SHOW_TEXT */), nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(function (n) {
      if (n.parentNode && n.parentNode.closest && n.parentNode.closest('[data-th], code, script, style, textarea')) return;
      var key = norm(n.nodeValue);
      if (key && Object.prototype.hasOwnProperty.call(dict.text, key)) {
        var lead = /^\s*/.exec(n.nodeValue)[0], tail = /\s*$/.exec(n.nodeValue)[0];
        n.nodeValue = lead + dict.text[key] + tail;
      }
    });
    root.querySelectorAll('[placeholder], [title], [aria-label]').forEach(function (el) {
      ['placeholder', 'title', 'aria-label'].forEach(function (a) {
        var v = el.getAttribute(a);
        if (v && Object.prototype.hasOwnProperty.call(dict.text, v)) el.setAttribute(a, dict.text[v]);
      });
    });
  }

  function setLang(l) {
    try { localStorage.setItem('osign.lang', l); } catch (e) { /* this session only */ }
    location.reload();
  }

  document.documentElement.lang = lang;
  window.t = t;
  window.I18N = { lang: lang, langs: LANGS, apply: apply, setLang: setLang, t: t };
})();
