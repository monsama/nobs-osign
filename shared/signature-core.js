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

  // The fields of the first version, the same placeholders the signatures used before (CI-Sign).
  // "graph" is where the add-in reads the value: a property of the user in Microsoft Graph
  // (/me). Graph has no description, homePhone or info, so those start empty; Entra Connect can
  // sync them as directory extensions (extension_<appId>_description) or into
  // onPremisesExtensionAttributes.extensionAttributeN, and that path goes in here.
  var DEFAULT_FIELDS = [
    { key: 'givenName',   label: 'Vorname',            ad: 'givenName',   graph: 'givenName',   sample: 'Anna' },
    { key: 'sn',          label: 'Name',               ad: 'sn',          graph: 'surname',     sample: 'Muster' },
    { key: 'description', label: 'Beschreibung 1',     ad: 'description', graph: '',            sample: 'Dipl. Ing. FH' },
    { key: 'department',  label: 'Beschreibung 2',     ad: 'department',  graph: 'department',  sample: 'Informatik' },
    { key: 'title',       label: 'Beschreibung 3',     ad: 'title',       graph: 'jobTitle',    sample: 'Projektleiterin' },
    { key: 'mobile',      label: 'Handy',              ad: 'mobile',      graph: 'mobilePhone', sample: '+41 79 123 45 67' },
    { key: 'homePhone',   label: 'Direkte Nr. ext.',   ad: 'homePhone',   graph: '',            sample: '+41 44 123 45 67' },
    { key: 'info',        label: 'Abwesenheit',        ad: 'info',        graph: '',            sample: '' },
    { key: 'mail',        label: 'Mail',               ad: 'mail',        graph: 'mail',        sample: 'anna.muster@example.com' },
    { key: 'sAMAccountName', label: 'Benutzername',    ad: 'sAMAccountName', graph: 'onPremisesSamAccountName', sample: 'mustera' }
  ];

  function defaultConfig() {
    return {
      schemaVersion: SCHEMA_VERSION,
      updated: null,
      assetBaseUrl: '',
      auth: { clientId: '', tenantId: '' },
      options: {
        // Outlook's own signature (classic: the one picked in File > Options > Mail > Signatures;
        // new Outlook: the roaming one) is switched off for the message, so a mail never has two.
        disableOutlookSignature: true,
        // New meetings get a signature too.
        meetings: false,
        // Users may pick, in the add-in's pane, which of the signatures open to them is used for
        // new mail and which for replies. Off: the central defaults always apply.
        allowUserChoice: true,
        // Without the user's details (Graph not set up, or not reachable and nothing cached) a
        // signature would go out with no name in it. Off by default: nothing is inserted then.
        insertWithoutProfile: false
      },
      fields: DEFAULT_FIELDS.map(function (f) { return Object.assign({}, f); }),
      signatures: []
    };
  }

  function newSignature(id, name) {
    return {
      id: id,
      name: name || 'New signature',
      html: '',
      compiled: '',
      defaultNew: false,
      defaultReply: false,
      removeEmptyLines: true,
      // Who gets this signature. Empty lists mean everyone.
      audience: { departments: [], domains: [], groups: [] }
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
      for (var i = 0; i < n.attributes.length; i++) add(placeholdersIn(n.attributes[i].value));
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
  function markBrLines(doc, block) {
    var segs = [[]];
    for (var c = block.firstChild; c; c = c.nextSibling) {
      // A nested block is a line of its own (it gets its own marks), never part of a <br> line.
      if (c.nodeType === 1 && BLOCK[c.nodeName]) { segs.push([]); continue; }
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

  // A block holding a <br> or a block of its own is a container of lines, not one line.
  function hasLines(el) {
    for (var c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && (c.nodeName === 'BR' || BLOCK[c.nodeName])) return true;
    return false;
  }

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
        var rowKeys = ownPlaceholders(el, true);
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
    html = html.replace(PLACEHOLDER, function (all, key) {
      if (!Object.prototype.hasOwnProperty.call(values, key)) return opts.keepUnknown === false ? '' : all;
      return isEmpty(values[key]) ? '' : escapeHtml(values[key]);
    });
    if (opts.assetBase !== undefined) {
      var base = String(opts.assetBase);
      if (base && !/\/$/.test(base)) base += '/';
      html = html.replace(/(\s(?:src|background)\s*=\s*["'])assets\//gi, function (all, pre) { return pre + base + 'assets/'; });
    }
    // An image no mail program can load - a path on some computer ("images/logo.gif", as imported),
    // or one whose address still holds a placeholder - would show as a broken image: it goes.
    if (opts.dropBrokenImages !== false) {
      html = html.replace(/<img\b[^>]*>/gi, function (tag) {
        var m = /\ssrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i.exec(tag);
        var src = m ? (m[1] !== undefined ? m[1] : m[2] !== undefined ? m[2] : m[3]) : '';
        var ok = /^(https?:|cid:|data:image\/|\/)/i.test(src) || /^assets\//i.test(src);
        return ok && src.indexOf('@@') < 0 ? tag : '';
      });
    }
    return html;
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

  // values: the user's fields (department is matched); from: the sender address of the mail;
  // groups: the names (or ids) of the user's groups, nested ones included - null when they are not
  // known, and then no signature limited to groups is open to the user.
  function isAvailable(sig, values, from, groups) {
    var a = sig.audience || {};
    var deps = lowerList(a.departments), doms = lowerList(a.domains), grps = lowerList(a.groups);
    if (deps.length && deps.indexOf(String((values || {}).department || '').trim().toLowerCase()) < 0) return false;
    if (doms.length && doms.indexOf(domainOf(from)) < 0) return false;
    if (grps.length) {
      var mine = lowerList(groups);
      if (!grps.some(function (g) { return mine.indexOf(g) >= 0; })) return false;
    }
    return true;
  }

  function available(config, values, from, groups) {
    return ((config && config.signatures) || []).filter(function (s) { return isAvailable(s, values, from, groups); });
  }

  // The group names the configuration limits signatures to: the add-in only needs to know these.
  function groupsUsed(config) {
    var out = [];
    ((config && config.signatures) || []).forEach(function (s) {
      lowerList((s.audience || {}).groups).forEach(function (g) { if (out.indexOf(g) < 0) out.push(g); });
    });
    return out;
  }

  // kind: 'new' or 'reply'. prefs: what the user picked in the add-in, { new: id|'none', reply: id|'none' }.
  // Returns the signature, or null for none.
  function pick(config, values, from, kind, prefs, groups) {
    var list = available(config, values, from, groups);
    var allowChoice = !config.options || config.options.allowUserChoice !== false;
    var want = allowChoice && prefs ? prefs[kind] : null;
    if (want === 'none') return null;
    if (want) { for (var i = 0; i < list.length; i++) if (list[i].id === want) return list[i]; }
    // Of the defaults open to the user, the one meant for fewer people wins: a default for one
    // sender domain, group or department beats the default for everyone.
    var flag = kind === 'reply' ? 'defaultReply' : 'defaultNew', best = null, bestScore = -1;
    for (var j = 0; j < list.length; j++) {
      if (!list[j][flag]) continue;
      var a = list[j].audience || {};
      var score = (lowerList(a.domains).length ? 4 : 0) + (lowerList(a.groups).length ? 2 : 0) + (lowerList(a.departments).length ? 1 : 0);
      if (score > bestScore) { best = list[j]; bestScore = score; }
    }
    return best;
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
    });
    var ids = {};
    (config.signatures || []).forEach(function (s, i) {
      if (!s || !/^[A-Za-z0-9_-]{1,64}$/.test(s.id || '')) errors.push('Signature ' + (i + 1) + ' has no valid id.');
      else if (ids[s.id]) errors.push('Two signatures share the id ' + s.id + '.');
      else ids[s.id] = 1;
      if (s && !String(s.name || '').trim()) errors.push('Signature ' + (i + 1) + ' has no name.');
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
    defaultConfig: defaultConfig,
    newSignature: newSignature,
    placeholdersIn: placeholdersIn,
    escapeHtml: escapeHtml,
    compile: compile,
    render: render,
    htmlToText: htmlToText,
    domainOf: domainOf,
    isAvailable: isAvailable,
    available: available,
    pick: pick,
    groupsUsed: groupsUsed,
    readPath: readPath,
    graphSelect: graphSelect,
    valuesFromGraph: valuesFromGraph,
    validate: validate
  };
});
