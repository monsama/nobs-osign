/*
 * signature-core.js - the one place a signature template turns into the HTML that goes into a mail.
 *
 * Used unchanged by three callers, so a preview in the editor is exactly what Outlook gets:
 *   - the editor's web UI (loaded as a plain <script>, exposes window.SignatureCore)
 *   - the Outlook add-in (bundled by esbuild through require())
 *   - the tests (node)
 *
 * Two steps:
 *   compile(html, doc)   EDITOR ONLY, needs a DOM. Marks every line that holds placeholders with
 *                        <!--sig-if:a,b--> ... <!--/sig-if-->, so the line can be dropped later
 *                        when all of its placeholders come out empty.
 *   render(compiled, values, opts)
 *                        Everywhere, string work only - the add-in's JavaScript-only runtime in
 *                        classic Outlook on Windows has no DOM. Fills @@placeholders (HTML-escaped)
 *                        and removes the marked lines whose fields are all empty.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SignatureCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SCHEMA_VERSION = 1;

  // A placeholder is @@ followed by a name: @@givenName, @@mobile. The name runs as far as letters,
  // digits and _ go, so @@mailNickname is never read as @@mail followed by "Nickname".
  var PLACEHOLDER = /@@([A-Za-z][A-Za-z0-9_]*)/g;
  // The photo for whoever has none of their own: photos/_default.<the extension the signature takes>.
  var PHOTO_DEFAULT = '_default';

  // The fields of the first version, the same placeholders the signatures used before (CI-Sign).
  // "graph" is where the add-in reads the value: a property of the user in Microsoft Graph
  // (/me). Graph has no description, homePhone or info, so those start empty; Entra Connect can
  // sync them as directory extensions (extension_<appId>_description) or into
  // onPremisesExtensionAttributes.extensionAttributeN, and that path goes in here.
  var DEFAULT_FIELDS = [
    { key: 'givenName',   label: 'Vorname',            ad: 'givenName',   graph: 'givenName',   sample: 'Anna' },
    { key: 'sn',          label: 'Nachname',           ad: 'sn',          graph: 'surname',     sample: 'Muster' },
    { key: 'description', label: 'Beschreibung 1',     ad: 'description', graph: '',            sample: 'Dipl. Ing. FH' },
    { key: 'department',  label: 'Abteilung',          ad: 'department',  graph: 'department',  sample: 'Informatik' },
    { key: 'title',       label: 'Funktion',           ad: 'title',       graph: 'jobTitle',    sample: 'Projektleiterin' },
    { key: 'mobile',      label: 'Handy',              ad: 'mobile',      graph: 'mobilePhone', sample: '+41 79 123 45 67' },
    { key: 'homePhone',   label: 'Direkte Nr. ext.',   ad: 'homePhone',   graph: '',            sample: '+41 44 123 45 67' },
    { key: 'info',        label: 'Abwesenheit',        ad: 'info',        graph: '',            sample: '' },
    { key: 'mail',        label: 'Mail',               ad: 'mail',        graph: 'mail',        sample: 'anna.muster@example.com' },
    { key: 'sAMAccountName', label: 'Benutzername',    ad: 'sAMAccountName', graph: 'onPremisesSamAccountName', sample: 'mustera' }
  ];

  // Fields a template may name beyond those: the AD attributes CI-Sign took a placeholder from
  // (@@telephoneNumber), with the Graph property for the same where there is one. A placeholder
  // matches whatever its case (@@wWWHomepage); a misspelling CI-Sign templates carry is known too.
  var KNOWN_FIELDS = [
    { key: 'displayName', label: 'Anzeigename', ad: 'displayName', graph: 'displayName', sample: 'Anna Muster' },
    { key: 'initials', label: 'Initialen', ad: 'initials', graph: '', sample: 'AM' },
    { key: 'company', label: 'Firma', ad: 'company', graph: 'companyName', sample: 'Contoso AG' },
    { key: 'telephoneNumber', label: 'Telefon', ad: 'telephoneNumber', graph: 'businessPhones[0]', sample: '+41 44 123 45 00' },
    { key: 'otherTelephone', label: 'Telefon 2', ad: 'otherTelephone', graph: '', sample: '+41 44 123 45 01' },
    { key: 'ipPhone', label: 'IP-Telefon', ad: 'ipPhone', graph: '', sample: '4500' },
    { key: 'facsimileTelephoneNumber', label: 'Fax', ad: 'facsimileTelephoneNumber', graph: 'faxNumber', sample: '+41 44 123 45 99' },
    { key: 'pager', label: 'Pager', ad: 'pager', graph: '', sample: '' },
    { key: 'physicalDeliveryOfficeName', label: 'Büro', ad: 'physicalDeliveryOfficeName', graph: 'officeLocation', sample: 'Zürich' },
    { key: 'streetAddress', label: 'Strasse', ad: 'streetAddress', graph: 'streetAddress', sample: 'Musterstrasse 1', also: ['streetAdress'] },
    { key: 'postOfficeBox', label: 'Postfach', ad: 'postOfficeBox', graph: '', sample: '' },
    { key: 'postalCode', label: 'PLZ', ad: 'postalCode', graph: 'postalCode', sample: '8000' },
    { key: 'l', label: 'Ort', ad: 'l', graph: 'city', sample: 'Zürich' },
    { key: 'st', label: 'Kanton', ad: 'st', graph: 'state', sample: 'ZH' },
    { key: 'co', label: 'Land', ad: 'co', graph: 'country', sample: 'Schweiz' },
    { key: 'c', label: 'Land (Kürzel)', ad: 'c', graph: '', sample: 'CH' },
    { key: 'wWWHomePage', label: 'Webseite', ad: 'wWWHomePage', graph: '', sample: 'www.contoso.ch' },
    { key: 'userPrincipalName', label: 'Anmeldename', ad: 'userPrincipalName', graph: 'userPrincipalName', sample: 'anna.muster@contoso.ch' },
    // The person's language (de-CH, en-US …): picks a signature's language variant.
    { key: 'preferredLanguage', label: 'Sprache', ad: 'preferredLanguage', graph: 'preferredLanguage', sample: 'de-CH' }
  ];
  // The field for a placeholder no field has yet: as it is known, or one of its own name (read
  // from the AD attribute of that name, which the server reads only if it is allowed).
  function fieldFor(key) {
    var k = String(key || '').toLowerCase();
    for (var i = 0; i < KNOWN_FIELDS.length; i++) {
      var f = KNOWN_FIELDS[i];
      if (f.key.toLowerCase() === k || (f.also || []).some(function (a) { return a.toLowerCase() === k; })) {
        return { key: key, label: f.label, ad: f.ad, graph: f.graph, sample: f.sample };
      }
    }
    return { key: key, label: key, ad: /^[A-Za-z][A-Za-z0-9-]{0,63}$/.test(key) ? key : '', graph: '', sample: '' };
  }

  // A new configuration: the starter fields, the default options, no signatures.
  function defaultConfig() {
    return {
      schemaVersion: SCHEMA_VERSION,
      updated: null,
      options: {
        // Outlook's own signature (classic: the one picked in File > Options > Mail > Signatures;
        // new Outlook: the roaming one) is switched off for the message, so a mail never has two.
        disableOutlookSignature: true,
        // New meetings get a signature too.
        meetings: false,
        // Users may pick, in the add-in's pane, which of the signatures open to them is used for
        // new mail and which for replies. Off: the central defaults always apply.
        allowUserChoice: true,
        // Users may make a signature of their own from a company one, all of it theirs to change
        // (in the add-in's pane, kept in their mailbox). Off: only the fields marked as editable.
        allowOwnSignatures: false,
        // Without the user's details (Graph not set up, or not reachable and nothing cached) a
        // signature would go out with no name in it. Off by default: nothing is inserted then.
        insertWithoutProfile: false,
        // Where the users' values come from:
        //   'graph'  Entra ID (Microsoft Graph), each field by its Graph source
        //   'ad'     Active Directory: the editor exports every user's fields and groups into
        //            users/<mail>.json next to signatures.json
        //   'mixed'  Graph for the fields that have a Graph source, AD for the others; groups from AD
        // New setups start with both (what the setup assistant recommends): the starter fields
        // without a Graph source (description, homePhone) then come from AD, not empty.
        valueSource: 'mixed'
      },
      fields: DEFAULT_FIELDS.map(function (f) { return Object.assign({}, f); }),
      signatures: []
    };
  }

  // A new signature for nobody yet, with the fields every signature has.
  function newSignature(id, name) {
    return {
      id: id,
      name: name || 'New signature',
      html: '',
      compiled: '',
      defaultNew: false,
      defaultReply: false,
      removeEmptyLines: true,
      // Who gets this signature: those in the groups, departments and sender domains given; with
      // none given, everyone only when "everyone" is set, else nobody (a new signature goes to
      // nobody until it is given an audience). A signature from before "everyone" was there has
      // no such flag: it goes to everyone, as it did.
      audience: { departments: [], domains: [], groups: [], everyone: false },
      // A campaign: only from / until these days (YYYY-MM-DD, both included); empty is always.
      validFrom: '',
      validUntil: ''
    };
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function placeholdersIn(text) {
    var out = [], m;
    PLACEHOLDER.lastIndex = 0;
    while ((m = PLACEHOLDER.exec(String(text || '')))) { if (out.indexOf(m[1]) < 0) out.push(m[1]); }
    return out;
  }

  // ---------------------------------------------------------------------------------------------
  // compile (needs a DOM: the browser's document, or jsdom's in the tests)
  // ---------------------------------------------------------------------------------------------
  var BLOCK = { P: 1, DIV: 1, LI: 1, TR: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, BLOCKQUOTE: 1, TABLE: 1, UL: 1, OL: 1, TBODY: 1, THEAD: 1, TD: 1, TH: 1 };
  // Lines that can be taken out whole. A cell (td) is not one of them: removing a cell shifts the
  // ones after it into the wrong column, so the row (tr) it sits in is what goes.
  var REMOVABLE = { P: 1, DIV: 1, LI: 1, TR: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, BLOCKQUOTE: 1 };

  // The placeholders in a node's own content - its text and attributes and those of its inline
  // children - leaving out what sits in nested block elements, which are lines of their own.
  function ownPlaceholders(node, cells) {
    var keys = [];
    function add(list) { list.forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); }
    function walk(n, top) {
      if (n.nodeType === 3) { add(placeholdersIn(n.nodeValue)); return; }
      if (n.nodeType !== 1) return;
      // A row counts what is in its cells (and what those cells hold), a nested table aside.
      if (!top && BLOCK[n.nodeName] && !(cells && n.nodeName !== 'TABLE')) return;
      // A person's photo (photos/@@sAMAccountName.jpg) is not a field of the line: without one,
      // the placeholder photo goes in, or nothing, and the logo next to it stays.
      for (var i = 0; i < n.attributes.length; i++) {
        if (n.nodeName === 'IMG' && /^src$/i.test(n.attributes[i].name) && /(^|\/)photos\//i.test(n.attributes[i].value)) continue;
        add(placeholdersIn(n.attributes[i].value));
      }
      for (var c = n.firstChild; c; c = c.nextSibling) walk(c, false);
    }
    walk(node, true);
    return keys;
  }

  function wrap(doc, first, last, keys) {
    var parent = first.parentNode;
    parent.insertBefore(doc.createComment('sig-if:' + keys.join(',')), first);
    parent.insertBefore(doc.createComment('/sig-if'), last.nextSibling);
  }

  // Inside one block, lines split by <br> ("Tel @@homePhone<br>Mobile @@mobile") are each a line.
  // So are the lines inside an inline element that holds <br>s itself, as CI-Sign writes them
  // (<font>@@info<br><br><img banner></font>): the banner is a line of its own, not part of
  // the line of an empty field.
  function markBrLines(doc, block) {
    var segs = [[]];
    for (var c = block.firstChild; c; c = c.nextSibling) {
      // A nested block is a line of its own (it gets its own marks), never part of a <br> line.
      if (c.nodeType === 1 && BLOCK[c.nodeName]) { segs.push([]); continue; }
      // An inline element that starts a line and holds <br>s itself (CI-Sign's <font>@@info<br><br>
      // <img>): its own lines. Not when something stands before it on the line (a label, "Tel"),
      // nor when it carries a field itself (<a href="mailto:@@mail">): then it is part of the line.
      if (c.nodeType === 1 && c.querySelector && c.querySelector('br') && !segs[segs.length - 1].some(visible) && !attrPlaceholders(c).length) {
        segs.push([]); markBrLines(doc, c); continue;
      }
      segs[segs.length - 1].push(c);
      if (c.nodeType === 1 && c.nodeName === 'BR') segs.push([]);
    }
    segs.forEach(function (seg) {
      if (!seg.length) return;
      var keys = [];
      seg.forEach(function (n) {
        var ks = n.nodeType === 3 ? placeholdersIn(n.nodeValue) : ownPlaceholders(n);
        ks.forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); });
      });
      if (keys.length) wrap(doc, seg[0], seg[seg.length - 1], keys);
    });
  }

  // Something on a line that shows: text that is not only space, or an element.
  function visible(n) { return n.nodeType === 1 || (n.nodeType === 3 && /\S/.test(n.nodeValue.replace(/\u00a0/g, 'x'))); }
  // The placeholders in an element's own attributes (an <a href="mailto:@@mail">).
  function attrPlaceholders(el) {
    var keys = [];
    for (var i = 0; i < el.attributes.length; i++) placeholdersIn(el.attributes[i].value).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); });
    return keys;
  }
  // A block holding a <br> or a block of its own is a container of lines, not one line; so is one
  // whose first inline element holds the <br>s (CI-Sign: <td><font>@@info<br><br><img></font>).
  function hasLines(el) {
    for (var c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && (c.nodeName === 'BR' || BLOCK[c.nodeName])) return true;
      if (c.nodeType === 1 && c.querySelector('br') && !attrPlaceholders(c).length) return true;
    }
    return false;
  }

  // A signature's HTML made into what the add-in fills in: fields as markers, lines that may go when empty marked.
  function compile(html, doc) {
    if (!doc) throw new Error('compile needs a document');
    var host = doc.createElement('div');
    host.innerHTML = String(html || '');
    // Strip what never belongs in a signature: scripts, and the editor's own markers if a compiled
    // template was pasted back in.
    Array.prototype.slice.call(host.querySelectorAll('script')).forEach(function (s) { s.parentNode.removeChild(s); });
    var walker = doc.createTreeWalker(host, 128 /* SHOW_COMMENT */);
    var stale = [];
    while (walker.nextNode()) if (/^\/?sig-if/.test(walker.currentNode.nodeValue)) stale.push(walker.currentNode);
    stale.forEach(function (n) { n.parentNode.removeChild(n); });

    var blocks = Array.prototype.slice.call(host.querySelectorAll('*')).filter(function (el) { return BLOCK[el.nodeName]; });
    // The top-level content counts as a block too: a signature that is just "@@givenName<br>@@mail".
    blocks.unshift(host);
    blocks.forEach(function (el) {
      if (el.nodeName === 'TR') {
        // A row goes as a whole when its fields are all empty ("Phone | @@telephoneNumber<br>@@mobile"),
        // but not one whose cell holds lines and a picture of its own (a field, then a banner): its
        // lines are judged each, so the picture stays.
        var multi = false;
        for (var td = el.firstChild; td; td = td.nextSibling) {
          if (td.nodeType === 1 && td.querySelector('br, p, div, table') && Array.prototype.some.call(td.querySelectorAll('img'), function (im) { return (im.getAttribute('src') || '').indexOf('@@') < 0; })) multi = true;
        }
        var rowKeys = multi ? [] : ownPlaceholders(el, true);
        if (rowKeys.length) wrap(doc, el, el, rowKeys);
        return;
      }
      if (el === host || hasLines(el)) { markBrLines(doc, el); return; }
      if (!REMOVABLE[el.nodeName]) return;
      var keys = ownPlaceholders(el);
      if (keys.length) wrap(doc, el, el, keys);
    });
    return host.innerHTML;
  }

  // ---------------------------------------------------------------------------------------------
  // render (string only - runs in the add-in's JavaScript-only runtime)
  // ---------------------------------------------------------------------------------------------
  var IF_BLOCK = /<!--sig-if:([A-Za-z0-9_,]*)-->((?:(?!<!--sig-if:)[\s\S])*?)<!--\/sig-if-->/;

  // Building blocks (a disclaimer, a banner, the address): kept once in config.blocks as
  // { key, name, html } and put into a signature as {{key}}, on a line of its own or in the text.
  // Expanded before a signature is compiled; a block's own {{…}} are not (no loops).
  var BLOCK_RE = /\{\{([A-Za-z][A-Za-z0-9_-]{0,39})\}\}/g;
  function blocksIn(html) {
    var out = [], m;
    BLOCK_RE.lastIndex = 0;
    while ((m = BLOCK_RE.exec(String(html || '')))) if (out.indexOf(m[1]) < 0) out.push(m[1]);
    return out;
  }
  // Building blocks put in where a signature names them ({{key}}); unknown ones left as typed.
  function expandBlocks(html, blocks) {
    var map = {};
    (blocks || []).forEach(function (b) { if (b && b.key) map[String(b.key).toLowerCase()] = String(b.html || ''); });
    var get = function (k) { return Object.prototype.hasOwnProperty.call(map, k.toLowerCase()) ? map[k.toLowerCase()] : null; };
    // In one pass: on a line of its own (a paragraph or div holding only the block) the block in its
    // place, in the text where it stands; what a block holds is not looked at again.
    return String(html || '').replace(/<(p|div)\b[^>]*>\s*\{\{([A-Za-z][A-Za-z0-9_-]{0,39})\}\}\s*<\/\1>|\{\{([A-Za-z][A-Za-z0-9_-]{0,39})\}\}/gi, function (all, tag, k1, k2) {
      var b = get(k1 || k2); return b === null ? all : b;
    });
  }

  // Whether a signature's compiled HTML (what the add-in renders) says what its html says: the
  // same elements in the same order, each with the same attributes, and the same text between
  // them. compile() only adds its line marks (<!--sig-if:…--> … <!--/sig-if-->) and the <tbody>
  // the browser puts into a table; any other comment in compiled is refused, as Outlook reads
  // some of them (<!--[if mso]>…<![endif]-->) and shows what is in them. The server checks it, as
  // it cannot compile without a DOM: so what an approver reads is what goes out.
  var MARK = /^<!--(?:sig-if:[A-Za-z0-9_,]*|\/sig-if)-->$/;
  var ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
  function decodeEntities(x) {
    return String(x).replace(/&(#[0-9]{1,7}|#x[0-9a-f]{1,6}|[a-z]+);/gi, function (all, e) {
      if (e.charAt(0) === '#') {
        var n = /^#x/i.test(e) ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        if (!(n > 0 && n <= 0x10ffff)) return all;
        if (n < 0x10000) return String.fromCharCode(n);
        n -= 0x10000;
        return String.fromCharCode(0xd800 + (n >> 10), 0xdc00 + (n & 0x3ff));
      }
      var k = e.toLowerCase();
      return Object.prototype.hasOwnProperty.call(ENTITIES, k) ? ENTITIES[k] : all;
    }).replace(/ /g, ' ');
  }
  // HTML as a list of what it is made of, in order: a start tag with its attributes (sorted by
  // name), an end tag, a stretch of text, what a style or the like holds (kept as written), a
  // comment; null when the HTML cannot be read (a tag or comment that does not end).
  var RAW_TEXT = { script: 1, style: 1, xmp: 1, iframe: 1, noembed: 1, noframes: 1, textarea: 1, title: 1 };
  function htmlTokens(x) {
    var s = String(x || ''), out = [], i = 0, text = '';
    var flush = function () { if (text) { out.push({ k: 't', v: text }); text = ''; } };
    while (i < s.length) {
      var lt = s.indexOf('<', i);
      if (lt < 0) { text += s.slice(i); break; }
      text += s.slice(i, lt);
      var rest = s.slice(lt);
      if (rest.slice(0, 4) === '<!--') {
        var end = s.indexOf('-->', lt + 4);
        if (end < 0) return null;
        flush(); out.push({ k: 'c', v: s.slice(lt, end + 3) }); i = end + 3; continue;
      }
      if (/^<[!?]/.test(rest)) {
        // <!doctype>, <![CDATA[ … ]]>, <![if !mso]>, <?xml …>: read by the browser as a comment.
        var gt = s.indexOf('>', lt);
        if (gt < 0) return null;
        flush(); out.push({ k: 'c', v: s.slice(lt, gt + 1) }); i = gt + 1; continue;
      }
      var m = /^<(\/?)([A-Za-z][^\s\/>]*)/.exec(rest);
      if (!m) { text += '<'; i = lt + 1; continue; }
      flush();
      var name = m[2].toLowerCase(), j = lt + m[0].length, attrs = [], seen = {};
      // The attributes: name, name=value, name="value", name='value'; the first of a name counts.
      for (;;) {
        while (j < s.length && /[\s\/]/.test(s.charAt(j))) j++;
        if (j >= s.length) return null;
        if (s.charAt(j) === '>') { j++; break; }
        var an = /^[^\s\/>=]+|^=[^\s\/>=]*/.exec(s.slice(j))[0];
        j += an.length;
        while (j < s.length && /\s/.test(s.charAt(j))) j++;
        var val = '';
        if (s.charAt(j) === '=') {
          j++;
          while (j < s.length && /\s/.test(s.charAt(j))) j++;
          var q = s.charAt(j);
          if (q === '"' || q === "'") {
            var close = s.indexOf(q, j + 1);
            if (close < 0) return null;
            val = s.slice(j + 1, close); j = close + 1;
          } else {
            var uv = /^[^\s>]*/.exec(s.slice(j))[0];
            val = uv; j += uv.length;
          }
        }
        an = an.toLowerCase();
        if (!seen[an]) { seen[an] = 1; attrs.push(an + '=' + decodeEntities(val).replace(/\s+/g, ' ').trim()); }
      }
      if (m[1]) { out.push({ k: 'e', v: name }); i = j; continue; }
      out.push({ k: 's', v: name + (attrs.length ? ' ' + attrs.sort().join(' ') : '') });
      i = j;
      // What a style or the like holds is not read as HTML: up to its end tag, compared as written.
      if (RAW_TEXT[name]) {
        var re = new RegExp('</' + name + '[\\s/>]', 'i'), at = re.exec(s.slice(i));
        var stop = at ? i + at.index : s.length;
        out.push({ k: 'r', v: s.slice(i, stop) }); i = stop;
      }
    }
    flush();
    return out;
  }
  // What a reader sees of HTML, in order, for comparing: the tags (without the <tbody> a browser
  // adds and a document's own html, head and body), the text between them with its white space
  // made one. Comments are left out here (compiledMatches looks at them first).
  function readable(tokens) {
    var out = [], text = '';
    var flush = function () { var t = decodeEntities(text).replace(/\s+/g, ' ').trim(); if (t) out.push('t:' + t); text = ''; };
    for (var i = 0; i < tokens.length; i++) {
      var t = tokens[i];
      if (t.k === 't') { text += t.v; continue; }
      if (t.k === 'c') continue;
      if (t.k === 'r') { flush(); out.push('r:' + t.v.replace(/\s+/g, ' ').trim()); continue; }
      var name = t.v.split(' ')[0];
      if (name === 'tbody' || name === 'html' || name === 'head' || name === 'body') continue;
      flush(); out.push(t.k + ':' + t.v);
    }
    flush();
    return out;
  }
  function compiledMatches(html, compiled) {
    if (typeof compiled !== 'string' || compiled === '') return true;
    var c = htmlTokens(compiled);
    if (!c) return false;
    // A comment in compiled that compile() does not write: refused (it could hold what the html does not show).
    if (c.some(function (t) { return t.k === 'c' && !MARK.test(t.v); })) return false;
    // compile() leaves out the html's scripts and any line marks pasted into it.
    var h = htmlTokens(String(html || '').replace(/<script\b[\s\S]*?<\/script\s*>/gi, ''));
    if (!h) return false;
    var a = readable(h), b = readable(c);
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  function isEmpty(v) { return v === undefined || v === null || String(v).trim() === ''; }

  // opts.assetBase: where the images live. Templates keep them as "assets/logo.png"; the add-in
  // puts the published address in front, the editor its own.
  // opts.removeEmptyLines: false keeps every line, empty or not.
  function render(compiled, values, opts) {
    opts = opts || {};
    values = values || {};
    var html = String(compiled || '');
    // Innermost marked line first: once it is settled, the line around it can be judged on its own.
    // A placeholder the values do not know (a typo, a field taken out) never makes a line go.
    var m, guard = 0, dropped = false;
    while ((m = IF_BLOCK.exec(html)) && guard++ < 10000) {
      var keys = m[1] ? m[1].split(',') : [];
      var drop = opts.removeEmptyLines !== false && keys.length > 0 &&
        keys.every(function (k) { return Object.prototype.hasOwnProperty.call(values, k) && isEmpty(values[k]); });
      if (drop) dropped = true;
      html = html.slice(0, m.index) + (drop ? '' : m[2]) + html.slice(m.index + m[0].length);
    }
    // The last line of a paragraph went: the <br> that ended the line before it would now end the
    // paragraph, and Outlook (Word) shows that as an empty line.
    if (dropped) html = html.replace(/<br\s*\/?>(\s*)(<\/(?:p|div|td|th|li|h[1-6]|blockquote)>)/gi, '$1$2');
    // A value that goes into an address (href, src) is taken only as a web, mail or phone address,
    // or as a relative one: anything else (javascript:, data:) could run or load in the mail.
    var inUrl = urlRanges(html);
    html = html.replace(PLACEHOLDER, function (all, key, at) {
      if (!Object.prototype.hasOwnProperty.call(values, key)) return opts.keepUnknown === false ? '' : all;
      if (isEmpty(values[key])) return '';
      if (inUrl(at) && !linkValueOk(values[key])) return '';
      return escapeHtml(values[key]);
    });
    // Images live in two folders next to signatures.json: assets/ (logos, icons) and photos/ (one
    // per person, by name: photos/@@sAMAccountName.jpg). opts.assetBase is put in front of both,
    // where the editor serves them; the add-in leaves them and embeds them in the mail instead.
    if (opts.assetBase !== undefined) {
      var base = String(opts.assetBase);
      if (base && !/\/$/.test(base)) base += '/';
      html = html.replace(/(\s(?:src|background)\s*=\s*["'])(assets|photos)\//gi, function (all, pre, dir) { return pre + base + dir.toLowerCase() + '/'; });
    }
    // A photo for a user with no name (photos/.jpg: no login, a cloud-only account): the placeholder
    // photo, photos/_default.jpg, which the add-in leaves out too when there is none.
    html = html.replace(/(\s(?:src)\s*=\s*["'](?:[^"']*\/)?photos\/)\.(png|jpe?g|gif)(["'])/gi, function (all, pre, ext, q) { return pre + PHOTO_DEFAULT + '.' + ext + q; });
    // An image no mail program can load - a path on some computer ("images/logo.gif", as imported),
    // an address still holding a placeholder, a photo for a user with no name - would show as a
    // broken image: it goes.
    if (opts.dropBrokenImages !== false) {
      html = html.replace(/<img\b[^>]*>/gi, function (tag) {
        var src = imgSrc(tag);
        var ok = /^(https?:|cid:|data:image\/|\/)/i.test(src) || /^(assets|photos)\/[^/]/i.test(src);
        if (/(^|\/)photos\/\./i.test(src)) ok = false;
        return ok && src.indexOf('@@') < 0 ? tag : '';
      });
    }
    return html;
  }

  // Where in the HTML the values of href, src and background attributes are: a function that
  // answers whether a position is inside one of them.
  function urlRanges(html) {
    var ranges = [], tre = /<[A-Za-z][^>]*>/g, are = /\s(?:href|src|background)\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, t, a;
    while ((t = tre.exec(html))) {
      are.lastIndex = 0;
      while ((a = are.exec(t[0]))) { var end = t.index + a.index + a[0].length; ranges.push([end - a[1].length, end]); }
    }
    return function (at) { return ranges.some(function (r) { return at >= r[0] && at < r[1]; }); };
  }
  // A value fit for an address: http(s):, mailto:, tel:, or none at all (relative, or the part
  // after a mailto: the template has). Spaces, control and invisible characters do not count, as
  // browsers skip them too (java\tscript:).
  function linkValueOk(v) {
    var s = String(v).replace(/[\u0000-\u0020\u007f-\u009f\u00ad\u200b-\u200f\u2028\u2029\u2060-\u2065\ufeff]/g, '');
    var m = /^([A-Za-z][A-Za-z0-9+.-]*):/.exec(s);
    return !m || /^(https?|mailto|tel)$/i.test(m[1]);
  }

  function imgSrc(tag) {
    var m = /\ssrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i.exec(tag);
    return m ? (m[1] !== undefined ? m[1] : m[2] !== undefined ? m[2] : m[3]) : '';
  }

  // A % that starts no escape ("photos/100%.jpg", a user name with a %) is the name as it is.
  function decodeName(s) { try { return decodeURIComponent(s); } catch (e) { return s; } }

  // The images a rendered signature loads from its own folders, as { tag, src, dir, name }:
  // what the add-in has to fetch and embed.
  function localImages(html) {
    var out = [], re = /<img\b[^>]*>/gi, m;
    while ((m = re.exec(String(html || '')))) {
      var src = imgSrc(m[0]), mm = /^(assets|photos)\/([^/?#]+)$/i.exec(src);
      // A name that is a path once decoded (%2F, ..) is none: never a file outside the folder.
      var nm = mm ? decodeName(mm[2]) : '';
      if (mm && !/[\\/]|^\.\.?$|\.\./.test(nm)) out.push({ tag: m[0], src: src, dir: mm[1].toLowerCase(), name: nm });
    }
    return out;
  }

  // For a mail written in plain text. Crude on purpose: no DOM here either.
  function htmlToText(html) {
    var t = String(html || '')
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|tr|li|h[1-6]|table|blockquote)>/gi, '\n')
      .replace(/<\/t[dh]>/gi, '\t')
      .replace(/<a\s[^>]*href\s*=\s*["']mailto:[^"']*["'][^>]*>/gi, '')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'").replace(/&#(\d+);/g, function (a, n) { return String.fromCharCode(+n); })
      .replace(/&amp;/g, '&');
    return t.split('\n').map(function (l) { return l.replace(/[ \t]+$/g, '').replace(/^[ \t]+/, ''); })
      .join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  // ---------------------------------------------------------------------------------------------
  // Who gets which signature
  // ---------------------------------------------------------------------------------------------
  function lowerList(a) { return (a || []).map(function (s) { return String(s).trim().toLowerCase(); }).filter(Boolean); }

  function domainOf(address) {
    var s = String(address || '').toLowerCase(), at = s.lastIndexOf('@');
    return at < 0 ? '' : s.slice(at + 1);
  }

  // The day as the user's clock has it, YYYY-MM-DD.
  function today(now) {
    var d = now instanceof Date ? now : new Date();
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  var DAY = /^\d{4}-\d{2}-\d{2}$/;
  // A campaign's days: is the signature on that day? One with no days set always is.
  function inPeriod(sig, now) {
    var d = today(now), from = String(sig.validFrom || ''), until = String(sig.validUntil || '');
    if (DAY.test(from) && d < from) return false;
    if (DAY.test(until) && d > until) return false;
    return true;
  }
  function isCampaign(sig) { return DAY.test(String(sig.validFrom || '')) || DAY.test(String(sig.validUntil || '')); }

  // values: the user's fields (department is matched); from: the sender address of the mail;
  // groups: the names (or ids) of the user's groups, nested ones included - null when they are not
  // known, and then no signature limited to groups is open to the user. now: the day (a Date),
  // for campaigns; today when left out.
  function isAvailable(sig, values, from, groups, now) {
    if (!inPeriod(sig, now)) return false;
    var a = sig.audience || {};
    var deps = lowerList(a.departments), doms = lowerList(a.domains), grps = lowerList(a.groups);
    if (deps.length && deps.indexOf(String((values || {}).department || '').trim().toLowerCase()) < 0) return false;
    if (doms.length && doms.indexOf(domainOf(from)) < 0) return false;
    if (grps.length) {
      var mine = lowerList(groups);
      if (!grps.some(function (g) { return mine.indexOf(g) >= 0; })) return false;
    }
    if (!deps.length && !doms.length && !grps.length) return a.everyone !== false;
    return true;
  }
  // Nobody gets it: no group, department or domain, and not for everyone.
  function forNobody(sig) {
    var a = (sig && sig.audience) || {};
    return !lowerList(a.departments).length && !lowerList(a.domains).length && !lowerList(a.groups).length && a.everyone === false;
  }

  function available(config, values, from, groups, now) {
    return ((config && config.signatures) || []).filter(function (s) { return isAvailable(s, values, from, groups, now); });
  }

  // The group names the configuration limits signatures to: the add-in only needs to know these.
  function groupsUsed(config) {
    var out = [];
    ((config && config.signatures) || []).forEach(function (s) {
      lowerList((s.audience || {}).groups).forEach(function (g) { if (out.indexOf(g) < 0) out.push(g); });
    });
    return out;
  }


  // The language of a person (or a signature) as two letters: de-CH, de_CH and DE are 'de'.
  function langOf(values) { return String((values || {}).preferredLanguage || '').trim().split(/[-_]/)[0].toLowerCase(); }

  // kind: 'new' or 'reply'. prefs: what the user picked in the add-in, { new: id|'none', reply: id|'none' }.
  // Returns the signature, or null for none.
  function pick(config, values, from, kind, prefs, groups, now) {
    var list = available(config, values, from, groups, now);
    var allowChoice = !config.options || config.options.allowUserChoice !== false;
    var want = allowChoice && prefs ? prefs[kind] : null;
    if (want === 'none') return null;
    if (want) { for (var i = 0; i < list.length; i++) if (list[i].id === want) return list[i]; }
    // Of the defaults open to the user, a campaign running now wins; then the one meant for fewer
    // people: a default for one sender domain, group or department beats the default for everyone.
    // A signature in a language (its language variant): the one in the person's own language wins;
    // one in another language only when there is no other default for them.
    var flag = kind === 'reply' ? 'defaultReply' : 'defaultNew', best = null, bestScore = -Infinity, lang = langOf(values);
    for (var j = 0; j < list.length; j++) {
      if (!list[j][flag]) continue;
      var a = list[j].audience || {}, sl = langOf({ preferredLanguage: list[j].language });
      var score = (sl ? (sl === lang ? 16 : -100) : 0) +
        (isCampaign(list[j]) ? 8 : 0) + (lowerList(a.domains).length ? 4 : 0) + (lowerList(a.groups).length ? 2 : 0) + (lowerList(a.departments).length ? 1 : 0);
      if (score > bestScore) { best = list[j]; bestScore = score; }
    }
    return best;
  }

  // ---------------------------------------------------------------------------------------------
  // Banner campaigns (config.campaigns): a banner for some days, put into signatures that are
  // there anyway, so a campaign is no copy of a signature. A signature says where with
  // {{campaign}} (on a line of its own); without it, the banner goes below (or above) it all.
  // { id, name, validFrom, validUntil, html, compiled, signatures: [ids] (none: all),
  //   position: 'below'|'above', newOnly, audience (as a signature's; none: everyone),
  //   langs: { en: { html, compiled }, … } (the banner in a language: for those who speak it) }
  // ---------------------------------------------------------------------------------------------
  var SLOT_RE = /<(p|div)\b[^>]*>\s*\{\{campaign\}\}\s*<\/\1>|\{\{campaign\}\}/gi;
  function hasSlot(html) { SLOT_RE.lastIndex = 0; var r = SLOT_RE.test(String(html || '')); SLOT_RE.lastIndex = 0; return r; }
  // The campaign running for this signature and person on that day, or null. ctx: { values, from,
  // groups, kind ('new'|'reply'), now, anyone (true: who it is for not looked at: the editor's preview) }.
  function campaignFor(config, sig, ctx) {
    ctx = ctx || {};
    var list = (config && config.campaigns) || [], best = null, bestScore = -Infinity, lang = langOf(ctx.values);
    if (!sig) return null;
    for (var i = 0; i < list.length; i++) {
      var c = list[i];
      if (!c || !inPeriod(c, ctx.now)) continue;
      // Only on some days of the week (0 Sunday … 6 Saturday), as the user's clock has it.
      if (c.days && c.days.length && c.days.indexOf((ctx.now instanceof Date ? ctx.now : new Date()).getDay()) < 0) continue;
      if (c.newOnly && ctx.kind === 'reply') continue;
      var targets = c.signatures || [];
      if (targets.length && targets.indexOf(sig.id) < 0 && targets.indexOf(sig.variantOf) < 0) continue;
      if (!ctx.anyone && !isAvailable({ audience: c.audience || {} }, ctx.values, ctx.from, ctx.groups, ctx.now)) continue;
      var a = c.audience || {};
      var score = (c.langs && lang && c.langs[lang] ? 16 : 0) + (isCampaign(c) ? 0 : -8) + (lowerList(a.domains).length ? 4 : 0) + (lowerList(a.groups).length ? 2 : 0) + (lowerList(a.departments).length ? 1 : 0);
      if (score > bestScore) { best = c; bestScore = score; }
    }
    return best;
  }
  // The banner's HTML (compiled) in the person's language, or the one for all; with banners that
  // take turns (c.more), one of them: `turn` (a number) says which, else by chance, per mail.
  function campaignHtml(c, values, turn) {
    if (!c) return '';
    var lang = langOf(values), v = c.langs && lang && c.langs[lang];
    if (v) return String(v.compiled || v.html || '');
    var all = [c].concat((c.more || []).filter(function (x) { return x && (x.compiled || x.html); }));
    var i = typeof turn === 'number' ? Math.abs(Math.floor(turn)) % all.length : Math.floor(Math.random() * all.length);
    return String(all[i].compiled || all[i].html || '');
  }
  // A signature's compiled HTML with the banner in its place (or below / above), or its place gone.
  function fillCampaign(compiled, banner, position) {
    var html = String(compiled || ''), b = String(banner || '');
    if (hasSlot(html)) return html.replace(SLOT_RE, function () { return b ? '<div>' + b + '</div>' : ''; });
    if (!b) return html;
    return position === 'above' ? '<div>' + b + '</div>' + html : html + '<div>' + b + '</div>';
  }
  // The signature as it goes out on that day: a copy with the banner put in (or its place empty).
  function withCampaign(config, sig, ctx) {
    if (!sig) return sig;
    var base = sig.compiled || sig.html || '';
    var c = campaignFor(config, sig, ctx);
    if (!c && !hasSlot(base)) return sig;
    var out = {};
    for (var k in sig) if (Object.prototype.hasOwnProperty.call(sig, k)) out[k] = sig[k];
    out.compiled = fillCampaign(base, campaignHtml(c, ctx && ctx.values, ctx && ctx.turn), c && c.position);
    out.campaign = c ? c.id : '';
    return out;
  }
  // A link in a banner with what tells the web statistics where the visit came from (utm_…):
  // made again each time, so it never goes in twice.
  function campaignSlug(name) {
    return String(name || 'campaign').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'campaign';
  }
  function withUtm(html, name) {
    var slug = campaignSlug(name);
    return String(html || '').replace(/(\shref\s*=\s*")(https?:\/\/[^"]*)(")/gi, function (all, pre, url, post) {
      var hash = '', at = url.indexOf('#');
      if (at >= 0) { hash = url.slice(at); url = url.slice(0, at); }
      var q = url.indexOf('?'), path = q >= 0 ? url.slice(0, q) : url;
      var keep = (q >= 0 ? url.slice(q + 1).replace(/&amp;/g, '&').split('&') : []).filter(function (x) { return x && !/^utm_(source|medium|campaign)=/i.test(x); });
      keep.push('utm_source=email', 'utm_medium=signature', 'utm_campaign=' + slug);
      return pre + path + '?' + keep.join('&amp;') + hash + post;
    });
  }
  // Campaigns that end within `days` with none after them for the same signatures: [{ c, left }].
  function campaignsEnding(config, days, now) {
    var d = today(now), list = (config && config.campaigns) || [], out = [];
    var plus = function (day, n) { var x = new Date(day + 'T12:00:00'); x.setDate(x.getDate() + n); return today(x); };
    list.forEach(function (c) {
      if (!c || !DAY.test(String(c.validUntil || '')) || c.validUntil < d || !inPeriod(c, now)) return;
      var left = Math.round((Date.parse(c.validUntil + 'T12:00:00') - Date.parse(d + 'T12:00:00')) / 86400000);
      if (left > days) return;
      var next = plus(c.validUntil, 1);
      var same = function (o) { return JSON.stringify((o.signatures || []).slice().sort()) === JSON.stringify((c.signatures || []).slice().sort()); };
      var follows = list.some(function (o) { return o !== c && same(o) && (!o.validFrom || o.validFrom <= next) && (!o.validUntil || o.validUntil >= next); });
      if (!follows) out.push({ c: c, left: left });
    });
    return out;
  }

  // ---------------------------------------------------------------------------------------------
  // A user's own variants: a copy of a company signature under a name of their own, with their own
  // text in the fields the editor marked "editable in the add-in" (a greeting, a line of their
  // own). Layout, images and every other field stay the company's. Kept in the user's mailbox.
  // ---------------------------------------------------------------------------------------------
  var VARIANT_MAX = 20, VARIANT_VALUE_MAX = 300, VARIANT_NAME_MAX = 60, VARIANT_OWN_MAX = 12000;
  // Signatures of their own (all of it): only when the editor allows it.
  function ownAllowed(config) {
    return !!config && !(config.options && config.options.allowUserChoice === false) && !!(config.options && config.options.allowOwnSignatures);
  }

  function editableKeys(config) {
    return ((config && config.fields) || []).filter(function (f) { return f && f.userEditable; }).map(function (f) { return f.key; });
  }
  // Variants are open when users may choose and at least one field may be changed.
  function variantsAllowed(config) {
    return !!config && !(config.options && config.options.allowUserChoice === false) && (editableKeys(config).length > 0 || ownAllowed(config));
  }
  // The editable fields a signature uses: only for those can a variant say something.
  function editableIn(config, sig) {
    var used = placeholdersIn(sig && sig.html);
    return editableKeys(config).filter(function (k) { return used.indexOf(k) >= 0; });
  }
  // The configuration with the user's variants among its signatures. A variant of a signature that
  // is gone, or not open to the user, goes with it; its values only for fields still editable.
  function withVariants(config, variants) {
    if (!variantsAllowed(config) || !Array.isArray(variants) || !variants.length) return config;
    var keys = editableKeys(config), extra = [];
    variants.slice(0, VARIANT_MAX).forEach(function (v) {
      if (!v || !/^my-[A-Za-z0-9_-]{1,40}$/.test(v.id || '')) return;
      var base = null;
      (config.signatures || []).forEach(function (s) { if (s.id === v.base) base = s; });
      if (!base) return;
      var vals = {};
      keys.forEach(function (k) { if (v.values && typeof v.values[k] === 'string') vals[k] = v.values[k].slice(0, VARIANT_VALUE_MAX); });
      // All of it their own (the pane compiled it): only while the editor allows it; else the
      // variant is not there (nor what it held), and comes back when it is allowed again.
      if (typeof v.own === 'string') {
        if (!ownAllowed(config) || v.own.length > VARIANT_OWN_MAX) return;
        extra.push(Object.assign({}, base, {
          id: v.id, name: String(v.name || base.name).slice(0, VARIANT_NAME_MAX), variantOf: base.id,
          defaultNew: false, defaultReply: false, own: true, html: v.own, compiled: v.own, userValues: undefined
        }));
        return;
      }
      if (!keys.length) return;
      extra.push(Object.assign({}, base, {
        id: v.id, name: String(v.name || base.name).slice(0, VARIANT_NAME_MAX), variantOf: base.id,
        defaultNew: false, defaultReply: false, userValues: vals
      }));
    });
    return Object.assign({}, config, { signatures: (config.signatures || []).concat(extra) });
  }
  // The values a signature is filled with: a variant's own text over the user's details.
  function valuesFor(sig, values) {
    if (!sig || !sig.userValues) return values;
    return Object.assign({}, values, sig.userValues);
  }

  // ---------------------------------------------------------------------------------------------
  // Values from a Microsoft Graph user (/me)
  // ---------------------------------------------------------------------------------------------
  // A path is a property name, with dots into objects and [n] into lists:
  //   jobTitle   onPremisesExtensionAttributes.extensionAttribute1   businessPhones[0]
  function readPath(obj, path) {
    var parts = String(path || '').match(/[^.[\]]+|\[\d+\]/g);
    if (!parts) return '';
    var cur = obj;
    for (var i = 0; i < parts.length; i++) {
      if (cur === null || cur === undefined) return '';
      var p = parts[i];
      cur = p.charAt(0) === '[' ? cur[+p.slice(1, -1)] : cur[p];
    }
    if (Array.isArray(cur)) cur = cur[0];
    return cur === null || cur === undefined ? '' : String(cur);
  }

  // ---------------------------------------------------------------------------------------------
  // Where each field's value comes from (options.valueSource)
  // ---------------------------------------------------------------------------------------------
  function valueSource(config) {
    var v = config && config.options && config.options.valueSource;
    return v === 'ad' || v === 'mixed' ? v : 'graph';
  }
  // 'graph', 'ad' or 'none' (then the field stays empty in Outlook).
  function fieldSource(config, f) {
    var mode = valueSource(config);
    if (mode === 'ad') return f.ad ? 'ad' : 'none';
    if (f.graph) return 'graph';
    return mode === 'mixed' && f.ad ? 'ad' : 'none';
  }
  function graphFields(config) { return (config.fields || []).filter(function (f) { return fieldSource(config, f) === 'graph'; }); }
  function adFields(config) { return (config.fields || []).filter(function (f) { return fieldSource(config, f) === 'ad'; }); }
  // Whether the editor has to export users from AD at all.
  function usesAdExport(config) { return valueSource(config) !== 'graph'; }

  // The file a user's exported values are in, by their mail address: users/<this>.
  function userFileName(mail) {
    var m = String(mail || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+$/.test(m)) return '';
    return m.replace(/[^a-z0-9@._-]/g, '_') + '.json';
  }

  // One set of values from both sources: each field from where it is to come from, written the way
  // the field says (a phone number alike for everyone).
  function mergeValues(config, graphValues, adValues) {
    var v = {};
    (config.fields || []).forEach(function (f) {
      var src = fieldSource(config, f);
      var from = src === 'graph' ? graphValues : src === 'ad' ? adValues : null;
      v[f.key] = formatValue(f, from && from[f.key] !== undefined && from[f.key] !== null ? String(from[f.key]) : '');
    });
    return v;
  }

  // A phone number written one way: +41 79 123 45 67. 079 123 45 67 and 0041… become +41…; other
  // countries get + for 00 and lose (0), their spacing kept. Anything else (an extension, letters,
  // too short) stays as it was.
  var CC2 = '20 27 30 31 32 33 34 36 39 40 41 43 44 45 46 47 48 49 51 52 53 54 55 56 57 58 60 61 62 63 64 65 66 81 82 84 86 90 91 92 93 94 95 98'.split(' ');
  // A phone number written one way (+41 79 123 45 67); see below.
  function formatPhone(value, country) {
    var raw = String(value || '').trim();
    if (!raw || /[A-Za-z]/.test(raw) || /[^\d\s+()./-]/.test(raw)) return raw;
    var d = raw.replace(/\(0\)/g, '').replace(/[^\d+]/g, '');
    if (/^00\d/.test(d)) d = '+' + d.slice(2);
    else if (/^0\d/.test(d)) d = '+' + (country || '41') + d.slice(1);
    if (!/^\+\d{7,15}$/.test(d)) return raw;
    var n = d.slice(1);
    var cc = /^[17]/.test(n) ? n.slice(0, 1) : CC2.indexOf(n.slice(0, 2)) >= 0 ? n.slice(0, 2) : n.slice(0, 3);
    var rest = n.slice(cc.length);
    // Switzerland and Liechtenstein as they are written there: 79 123 45 67, 232 12 34.
    if (cc === '41' && rest.length === 9) return '+41 ' + rest.slice(0, 2) + ' ' + rest.slice(2, 5) + ' ' + rest.slice(5, 7) + ' ' + rest.slice(7);
    if ((cc === '423' && rest.length === 7)) return '+423 ' + rest.slice(0, 3) + ' ' + rest.slice(3, 5) + ' ' + rest.slice(5);
    // Elsewhere area codes differ in length: the spacing as it was typed, with + and without (0).
    return raw.replace(/\s*\(0\)\s*/g, ' ').replace(/^00/, '+').replace(/^0(?=\d)/, '+' + (country || '41') + ' ').replace(/\s+/g, ' ').trim();
  }
  function formatValue(field, value) {
    if (!field || !value) return value;
    return field.format === 'phone' ? formatPhone(value) : value;
  }

  // The top-level properties Graph has to be asked for ($select) to fill these fields.
  function graphSelect(fields) {
    var out = [];
    (fields || []).forEach(function (f) {
      var top = String(f.graph || '').split(/[.[]/)[0].trim();
      if (top && /^[A-Za-z0-9_]+$/.test(top) && out.indexOf(top) < 0) out.push(top);
    });
    return out;
  }

  function valuesFromGraph(fields, user) {
    var v = {};
    (fields || []).forEach(function (f) { v[f.key] = f.graph ? readPath(user, f.graph) : ''; });
    return v;
  }

  // ---------------------------------------------------------------------------------------------
  // Config check: the editor runs it before saving and publishing, the add-in when it loads one.
  // ---------------------------------------------------------------------------------------------
  function validate(config) {
    var errors = [];
    if (!config || typeof config !== 'object') return ['The configuration is not an object.'];
    if (config.schemaVersion !== SCHEMA_VERSION) errors.push('Unknown schemaVersion ' + config.schemaVersion + '.');
    var keys = {};
    (config.fields || []).forEach(function (f, i) {
      if (!f || !/^[A-Za-z][A-Za-z0-9_]*$/.test(f.key || '')) errors.push('Field ' + (i + 1) + ': the placeholder name may hold only letters, digits and _, and starts with a letter.');
      else if (keys[f.key]) errors.push('Field @@' + f.key + ' is there twice.');
      else keys[f.key] = 1;
      if (f && f.format && f.format !== 'phone') errors.push('Field @@' + f.key + ': unknown format ' + f.format + '.');
    });
    var bkeys = {};
    (config.blocks || []).forEach(function (b, i) {
      if (b && String(b.key || '').toLowerCase() === 'campaign') errors.push('{{campaign}} is the place of the campaign banners: a building block cannot be called so.');
      else if (!b || !/^[A-Za-z][A-Za-z0-9_-]{0,39}$/.test(b.key || '')) errors.push('Building block ' + (i + 1) + ': its name in a signature ({{name}}) holds letters, digits, - and _, starting with a letter.');
      else if (bkeys[String(b.key).toLowerCase()]) errors.push('The building block {{' + b.key + '}} is there twice.');
      else bkeys[String(b.key).toLowerCase()] = 1;
    });
    // The house colours (the editor's palette): #rrggbb, at most 16.
    if (config.colors !== undefined && (!Array.isArray(config.colors) || config.colors.length > 16 || config.colors.some(function (c) { return !/^#[0-9a-f]{6}$/i.test(String(c)); }))) errors.push('The house colours are up to 16 colours like #1d4ed8.');
    var cids = {};
    (config.campaigns || []).forEach(function (c, i) {
      var nm = c && c.name ? '«' + c.name + '»' : String(i + 1);
      if (!c || !/^[A-Za-z0-9_-]{1,64}$/.test(c.id || '')) errors.push('Campaign ' + (i + 1) + ' has no valid id.');
      else if (cids[c.id]) errors.push('Two campaigns share the id ' + c.id + '.');
      else cids[c.id] = 1;
      if (c && ((c.validFrom && !DAY.test(c.validFrom)) || (c.validUntil && !DAY.test(c.validUntil)))) errors.push('Campaign ' + nm + ': its days are dates like 2026-12-01.');
      else if (c && c.validFrom && c.validUntil && c.validFrom > c.validUntil) errors.push('Campaign ' + nm + ' ends before it starts.');
      if (c && c.position && c.position !== 'below' && c.position !== 'above') errors.push('Campaign ' + nm + ': unknown position ' + c.position + '.');
      if (c && c.days !== undefined && (!Array.isArray(c.days) || c.days.some(function (d) { return d !== Math.floor(d) || d < 0 || d > 6; }))) errors.push('Campaign ' + nm + ': its weekdays are 0 (Sunday) to 6 (Saturday).');
      if (c && c.more !== undefined && (!Array.isArray(c.more) || c.more.length > 9)) errors.push('Campaign ' + nm + ': up to 10 banners take turns.');
    });
    var ids = {};
    (config.signatures || []).forEach(function (s, i) {
      if (!s || !/^[A-Za-z0-9_-]{1,64}$/.test(s.id || '')) errors.push('Signature ' + (i + 1) + ' has no valid id.');
      else if (ids[s.id]) errors.push('Two signatures share the id ' + s.id + '.');
      else ids[s.id] = 1;
      if (s && !String(s.name || '').trim()) errors.push('Signature ' + (i + 1) + ' has no name.');
      else if (s && String(s.name).length > 200) errors.push('Signature ' + (i + 1) + ': the name is longer than 200 characters.');
      if (s && ((s.validFrom && !DAY.test(s.validFrom)) || (s.validUntil && !DAY.test(s.validUntil)))) errors.push('Signature ' + (i + 1) + ': the campaign\'s days are dates like 2026-12-01.');
      else if (s && s.validFrom && s.validUntil && s.validFrom > s.validUntil) errors.push('Signature ' + (i + 1) + ': the campaign ends before it starts.');
    });
    return errors;
  }

  // ---------------------------------------------------------------------------------------------
  // Reading signature files written elsewhere (classic Outlook, CI-Sign): their encoding is told
  // by a byte order mark (CI-Sign writes UTF-16), by a <meta charset>, or not at all.
  // ---------------------------------------------------------------------------------------------
  function decodeHtmlFile(bytes) {
    var b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b.subarray(2));
    if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) return new TextDecoder('utf-16be').decode(b.subarray(2));
    if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder('utf-8').decode(b.subarray(3));
    // UTF-16 without a mark: every other byte of plain text is 0.
    if (b.length >= 4 && b[1] === 0 && b[3] === 0 && b[0] !== 0) return new TextDecoder('utf-16le').decode(b);
    var head = new TextDecoder('latin1').decode(b.subarray(0, 4096));
    var cs = (/charset\s*=\s*["']?([\w-]+)/i.exec(head) || [])[1];
    if (cs) { try { return new TextDecoder(cs).decode(b); } catch (e) { /* unknown name: below */ } }
    try { return new TextDecoder('utf-8', { fatal: true }).decode(b); }
    catch (e) { return new TextDecoder('windows-1252').decode(b); }
  }

  // An image's file name as the editor stores it (the same rule as the backend's).
  function assetName(name) {
    var base = String(name || '').split(/[\\/]/).pop();
    var dot = base.lastIndexOf('.');
    var ext = dot > 0 ? base.slice(dot).toLowerCase() : '';
    var stem = (dot > 0 ? base.slice(0, dot) : base).replace(/[^A-Za-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'image';
    return stem + ext;
  }

  return {
    SCHEMA_VERSION: SCHEMA_VERSION,
    decodeHtmlFile: decodeHtmlFile,
    assetName: assetName,
    DEFAULT_FIELDS: DEFAULT_FIELDS,
    fieldFor: fieldFor,
    defaultConfig: defaultConfig,
    newSignature: newSignature,
    placeholdersIn: placeholdersIn,
    blocksIn: blocksIn,
    compiledMatches: compiledMatches,
    langOf: langOf,
    expandBlocks: expandBlocks,
    escapeHtml: escapeHtml,
    compile: compile,
    render: render,
    htmlToText: htmlToText,
    localImages: localImages, 
    isAvailable: isAvailable,
    inPeriod: inPeriod,
    isCampaign: isCampaign,
    hasSlot: hasSlot,
    campaignFor: campaignFor,
    campaignHtml: campaignHtml,
    fillCampaign: fillCampaign,
    withCampaign: withCampaign,
    withUtm: withUtm,
    campaignSlug: campaignSlug,
    campaignsEnding: campaignsEnding,
    forNobody: forNobody,
    PHOTO_DEFAULT: PHOTO_DEFAULT,
    today: today,
    available: available,
    pick: pick, 
    editableIn: editableIn,
    variantsAllowed: variantsAllowed,
    ownAllowed: ownAllowed,
    VARIANT_OWN_MAX: VARIANT_OWN_MAX,
    withVariants: withVariants,
    valuesFor: valuesFor,
    VARIANT_MAX: VARIANT_MAX,
    VARIANT_VALUE_MAX: VARIANT_VALUE_MAX,
    VARIANT_NAME_MAX: VARIANT_NAME_MAX,
    groupsUsed: groupsUsed,
    valueSource: valueSource,
    fieldSource: fieldSource,
    graphFields: graphFields,
    adFields: adFields,
    usesAdExport: usesAdExport,
    userFileName: userFileName,
    mergeValues: mergeValues, formatPhone: formatPhone, formatValue: formatValue, 
    graphSelect: graphSelect,
    valuesFromGraph: valuesFromGraph,
    validate: validate
  };
});
