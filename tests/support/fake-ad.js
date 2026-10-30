// Stands in for Active Directory over LDAP, for the server version (server/ldap.js,
// server/directory.js). It evaluates the filters it is sent, with AD's own two rules: IN_CHAIN
// (1.2.840.113556.1.4.1941, membership through nested groups) and the bit test
// (1.2.840.113556.1.4.803, a disabled account in userAccountControl).
const ldap = require('ldapjs');

const IN_CHAIN = '1.2.840.113556.1.4.1941';
const BIT_AND = '1.2.840.113556.1.4.803';
const lc = (s) => String(s).toLowerCase();

// people: { login: { pw, name, mail, attrs, disabled, groups: [group name] } }
// groups: { name: [group names it is in] }
function fakeAd({ base, people, groups, service }) {
  const entries = [];
  const groupDn = (g) => 'CN=' + g + ',OU=Groups,' + base;
  for (const [g, parents] of Object.entries(groups)) {
    entries.push({ dn: groupDn(g), attrs: { objectclass: ['top', 'group'], name: [g], samaccountname: [g], distinguishedname: [groupDn(g)], memberof: parents.map(groupDn) } });
  }
  for (const [login, p] of Object.entries(people)) {
    const dn = 'CN=' + p.name + ',OU=Users,' + base;
    const a = { objectclass: ['top', 'person', 'user'], objectcategory: ['person'], samaccountname: [login], displayname: [p.name], distinguishedname: [dn],
      userprincipalname: [login + '@example.ch'], useraccountcontrol: [String(p.disabled ? 514 : 512)], memberof: (p.groups || []).map(groupDn) };
    if (p.mail) a.mail = [p.mail];
    for (const [k, v] of Object.entries(p.attrs || {})) if (v !== '') a[lc(k)] = [v];
    entries.push({ dn, attrs: a });
  }
  const byDn = (dn) => entries.find((e) => lc(e.dn) === lc(dn));
  // The groups an entry is in, nested ones included.
  function inGroups(e, seen = new Set()) {
    for (const g of e.attrs.memberof || []) {
      if (seen.has(lc(g))) continue;
      seen.add(lc(g));
      const ge = byDn(g); if (ge) inGroups(ge, seen);
    }
    return seen;
  }
  function matches(f, e) {
    const vals = (attr) => e.attrs[lc(attr)] || [];
    switch (f.type) {
      case 'AndFilter': return f.clauses.every((c) => matches(c, e));
      case 'OrFilter': return f.clauses.some((c) => matches(c, e));
      case 'NotFilter': return !matches(f.filter, e);
      case 'PresenceFilter': return vals(f.attribute).length > 0;
      case 'EqualityFilter': return vals(f.attribute).some((v) => lc(v) === lc(f.value));
      case 'SubstringFilter': return vals(f.attribute).some((v) => {
        const s = lc(v);
        return s.startsWith(lc(f.subInitial || '')) && s.endsWith(lc(f.subFinal || '')) && (f.subAny || []).every((x) => s.includes(lc(x)));
      });
      case 'ExtensibleFilter':
        if (f.matchingRule === IN_CHAIN && lc(f.attribute) === 'memberof') return inGroups(e).has(lc(f.matchValue));
        if (f.matchingRule === IN_CHAIN && lc(f.attribute) === 'member') { const m = byDn(f.matchValue); return !!m && inGroups(m).has(lc(e.dn)); }
        if (f.matchingRule === BIT_AND) return vals(f.attribute).some((v) => (Number(v) & Number(f.matchValue)) === Number(f.matchValue));
        throw new Error('rule not known here: ' + f.matchingRule);
      default: throw new Error('filter not known here: ' + f.type);
    }
  }

  const searches = [];
  const server = ldap.createServer();
  server.bind(base, (req, res, next) => {
    const dn = lc(req.dn.toString()), pw = req.credentials;
    if (dn === lc(service.dn) && pw === service.pw) { res.end(); return next(); }
    const who = Object.keys(people).find((k) => dn === lc('cn=' + k + '@example.ch,' + base));
    if (who && pw === people[who].pw) { res.end(); return next(); }
    return next(new ldap.InvalidCredentialsError());
  });
  server.search('', (req, res, next) => {
    // The RootDSE, which AD lets anyone read: the domain it serves.
    if (req.dn.toString() === '') {
      res.send(new ldap.SearchEntry({ objectName: '', attributes: [
        new ldap.Attribute({ type: 'defaultNamingContext', values: [base] }), new ldap.Attribute({ type: 'dnsHostName', values: ['dc01.example.ch'] })] }));
      res.end(); return next();
    }
    // What was asked: the filter, and the attributes asked for.
    searches.push(req.filter.toString() + ' attributes=' + (req.attributes || []).join(','));
    const scope = lc(req.dn.toString());
    const hits = entries.filter((e) => lc(e.dn).endsWith(scope) && matches(req.filter, e));
    const limit = req.sizeLimit || 0;
    for (const e of limit ? hits.slice(0, limit) : hits) res.send({ dn: e.dn, attributes: e.attrs });
    if (limit && hits.length > limit) return next(new ldap.SizeLimitExceededError());
    res.end(); return next();
  });
  return {
    searches, entries,
    // A change in the directory, between two reads.
    set(login, attr, value) { const e = entries.find((x) => (x.attrs.samaccountname || [])[0] === login); if (value === null) delete e.attrs[lc(attr)]; else e.attrs[lc(attr)] = [value]; },
    remove(login) { const i = entries.findIndex((x) => (x.attrs.samaccountname || [])[0] === login); entries.splice(i, 1); },
    hideAll() { const keep = entries.filter((e) => !(e.attrs.objectcategory || []).includes('person')); entries.length = 0; entries.push(...keep); },
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve('ldap://127.0.0.1:' + server.address().port))),
    close: () => server.close()
  };
}

module.exports = { fakeAd };
