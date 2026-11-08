/*
 * Signature HTML put into the pane's own page (to edit it, or to change one's fields right in it):
 * the company's markup and the user's details, but nothing in it may run there. So it is read
 * apart first (DOMParser, where nothing runs or loads), cleaned, and its nodes moved over - not
 * written as text again, which the browser could read differently a second time.
 */

// Elements that run, load a page of their own, or change how the page around them is read.
const DROP = 'script,iframe,frame,frameset,object,embed,applet,base,meta,link,noscript,noembed,noframes,template,xmp,plaintext';
// Attributes holding an address.
const URL_ATTR = /^(href|src|xlink:href|action|background|poster|lowsrc|dynsrc|cite|longdesc|data)$/i;

// An address fit for the pane: http(s):, mailto:, tel:, cid:, an image as data:, or relative.
// Spaces, control and invisible characters do not count (java\tscript: is javascript:).
export function safeUrl(value) {
  const s = String(value || '').replace(/[\u0000-\u0020\u007f-\u009f\u00ad\u200b-\u200f\u2028\u2029\u2060-\u2065\ufeff]/g, '');
  const m = /^([A-Za-z][A-Za-z0-9+.-]*):/.exec(s);
  if (!m) return true;
  return /^(https?|mailto|tel|cid)$/i.test(m[1]) || /^data:image\//i.test(s);
}

// Cleans what is under root, in place: what runs or loads goes, styles stay (but none that runs
// script, as old Internet Explorer's expression() did).
export function cleanTree(root) {
  root.querySelectorAll(DROP).forEach((n) => n.remove());
  root.querySelectorAll('style').forEach((n) => { if (/expression\s*\(|javascript:|behavior\s*:|-moz-binding/i.test(n.textContent)) n.remove(); });
  root.querySelectorAll('*').forEach((n) => {
    [...n.attributes].forEach((a) => {
      if (/^on/i.test(a.name) || /^(srcset|srcdoc|formaction)$/i.test(a.name)) n.removeAttribute(a.name);
      else if (URL_ATTR.test(a.name) && !safeUrl(a.value)) n.removeAttribute(a.name);
      else if (/^style$/i.test(a.name) && /expression\s*\(|javascript:|behavior\s*:|-moz-binding/i.test(a.value)) n.removeAttribute(a.name);
    });
  });
  return root;
}

// The cleaned nodes of html, made for doc (the pane's page).
export function safeNodes(html, doc) {
  doc = doc || document;
  const Parser = (doc.defaultView && doc.defaultView.DOMParser) || DOMParser;
  const parsed = new Parser().parseFromString('<!doctype html><html><head></head><body>' + String(html || '') + '</body></html>', 'text/html');
  cleanTree(parsed.body);
  return [...parsed.body.childNodes].map((n) => doc.importNode(n, true));
}

// In place of el.innerHTML = html.
export function setSafeHtml(el, html) {
  el.replaceChildren(...safeNodes(html, el.ownerDocument));
}
