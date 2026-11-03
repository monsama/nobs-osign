/*
 * The directory: Active Directory (with the service account) and Entra ID (Microsoft Graph), for
 * the preview, the suggestions while typing, and the users' values read and written. Routes for
 * server/app.js; live() gives what changes with the connections (adOn, graph).
 */
'use strict';
const core = require('../shared/signature-core.js');
const { directoryProblem, lookup, groupsStartingWith, usersMatching, departmentsMatching } = require('./directory');
const { graphProblem } = require('./graph');

// What a field may read from Entra ID for the preview: contact details, as for AD
// (server/directory.js), and the attributes synchronised from AD for signatures
// (onPremisesExtensionAttributes, directory extensions). Not anything else of anyone.
const GRAPH_CONTACT = ['displayName', 'givenName', 'surname', 'jobTitle', 'department', 'companyName', 'officeLocation', 'businessPhones', 'mobilePhone',
  'faxNumber', 'mail', 'userPrincipalName', 'streetAddress', 'postalCode', 'city', 'state', 'country', 'preferredLanguage', 'onPremisesSamAccountName',
  'employeeId', 'employeeType', 'aboutMe', 'mySite', 'otherMails', 'proxyAddresses'].map((x) => x.toLowerCase());
function graphOk(p) {
  const head = String(p).split(/[.[]/)[0];
  return GRAPH_CONTACT.includes(head.toLowerCase()) || /^onPremisesExtensionAttributes$/i.test(head) || /^extension_[A-Za-z0-9]{32}_[A-Za-z0-9_]{1,100}$/.test(head);
}

function createDirectoryRoutes({ live, dir, graphCfg, users, settings, store, history, publisher, needsApproval }) {
  const routes = {
    // Active Directory, with the service account.
    'ad/lookup': async (d, s, adOn, graph) => {
      if (!adOn) return { error: directoryProblem(dir) };
      // Nothing typed: the one signed in, as OSign.ps1 takes the Windows user.
      const q = String(d.query || '').trim() || (s.user.local ? '' : s.user.login);
      try { return await lookup(dir, q, Array.isArray(d.fields) ? d.fields : []); } catch (e) { return { error: e.message }; }
    },
    // A person from Entra ID, for the preview: the fields with a Graph source, and the groups.
    'entra/lookup': async (d, s, adOn, graph) => {
      if (!graph) return { error: graphProblem(graphCfg) || 'Publishing to SharePoint is not set up.' };
      // Only what the fields read (the draft's or the published ones): not any property of anyone.
      const known = new Set();
      for (const text of [store.getDraft().configJson || '', history.latest() || '']) {
        try { for (const f of (JSON.parse(text) || {}).fields || []) if (f && f.graph) known.add(String(f.graph)); } catch (e) { /* none */ }
      }
      const fields = (Array.isArray(d.fields) ? d.fields : []).filter((f) => f && f.key && f.graph && known.has(String(f.graph)) && graphOk(String(f.graph))).map((f) => ({ key: String(f.key), graph: String(f.graph) }));
      try {
        const r = await graph.user(d.query, core.graphSelect(fields));
        return { values: core.valuesFromGraph(fields, r.user), displayName: r.user.displayName || d.query, groups: r.groups };
      } catch (e) { return { error: e.message }; }
    },
    'ad/groups': async (d, s, adOn, graph) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return { groups: await groupsStartingWith(dir, d.query) }; } catch (e) { return { error: e.message }; }
    },
    // Suggestions while typing: people (for the preview) and departments (for the audience).
    'ad/users': async (d, s, adOn, graph) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return { users: await usersMatching(dir, d.query) }; } catch (e) { return { error: e.message }; }
    },
    'ad/departments': async (d, s, adOn, graph) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return { departments: await departmentsMatching(dir, d.query) }; } catch (e) { return { error: e.message }; }
    },
    'entra/users': async (d, s, adOn, graph) => {
      if (!graph) return { error: graphProblem(graphCfg) || 'Publishing to SharePoint is not set up.' };
      try { return { users: await graph.users(d.query) }; } catch (e) { return { error: e.message }; }
    },
    'ad/read': async (d, s, adOn, graph) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return await users.apiRead(settings.target(), d); } catch (e) { return { error: 'Active Directory: ' + e.message }; }
    },
    'ad/snapshot': (d, s, adOn, graph) => users.apiSnapshot(d),
    'ad/export': async (d, s, adOn, graph) => {
      if (!publisher(s)) return { error: 'Only the publishing group writes the users.' };
      // With four eyes, the users are written when the publish is approved, as it was approved.
      if (needsApproval()) return { error: 'With approval on, the users are written when the publish is approved.' };
      try { return await users.apiExport(settings.target(), d); } catch (e) { return { error: 'Writing the users: ' + e.message }; }
    }
  };
  // What changes with the connections, read when a route runs.
  for (const [k, fn] of Object.entries(routes)) routes[k] = (d, s) => { const { adOn, graph } = live(); return fn(d, s, adOn, graph); };
  return { routes };
}

module.exports = { createDirectoryRoutes, graphOk };
