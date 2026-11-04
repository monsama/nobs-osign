/*
 * The server's connections (Settings → Connections), for the local account admin: what is saved,
 * the tests, the certificate, the domain detected; the setup script and the deployment check; and
 * the SharePoint folder for the setup assistant. Routes for server/app.js; live() gives what
 * changes with the connections (graph).
 */
'use strict';
const addin = require('./addin');
const { testSignIn, detect } = require('./ldap');
const { testService } = require('./directory');
const { graphProblem } = require('./graph');
const { check: checkSettings } = require('./settings');

function createConnectionRoutes({ connections, connectionsAnswer, settings, certs, setup, deployment, addinId, addinUrlFromPages, configure, targetFor,
  graphCfg, ldapCfg, dir, pages, fetchFn, adminOnly, live, log }) {
  const routes = {
    // The setup script: made for this setup (the site and the add-in's address as saved), and
    // whether it has sent its result yet.
    'setup/start': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      if (!connections.answer().keyOk) return { error: 'The server has no key to keep the client secret with (secrets/app_key): run sudo ./deploy-nobs-osign.sh again.' };
      await addinUrlFromPages();
      const st = settings.get();
      // A new certificate for this run: only its public half goes into the script.
      const cert = certs.pending();
      return setup.start({ cert: cert.certDer, certThumbprint: cert.thumbprint, origin: String(d.origin || ''), site: st.spFolder ? st.spSite || 'root' : '', addinUrl: st.addinUrl, groups: !!d.groups, users: !!d.users,
        // The apps of an earlier run, the only ones the script may use again.
        knownAddin: st.clientId, knownServer: graphCfg.clientId || '' });
    },
    'deploy/start': (d, s) => (adminOnly(s) ? { error: adminOnly(s) } : deployment.start({ origin: String(d.origin || ''), addinId: addinId() })),
    'deploy/status': () => ({ last: deployment.last(), waiting: deployment.waiting() }),
    'setup/status': (d, s) => (adminOnly(s) ? { error: adminOnly(s) } : setup.status()),
    // "Test" in Server connections: each part, with what is saved, for real. → { message } or { error }.
    'connections/test': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      try {
        if (d.what === 'signIn') {
          const look = dir.serviceDn && dir.servicePassword ? { dn: dir.serviceDn, password: dir.servicePassword } : null;
          const r = await testSignIn(ldapCfg, look);
          if (r.group === false) return { error: 'The domain controller answers, but there is no sign-in group by that name: check the group\'s DN.' };
          return { message: r.group ? 'The domain controller answers, and the sign-in group is there.' : 'The domain controller answers (TLS as set). With the service account set, the group is checked too.' };
        }
        // Active Directory as a whole: the domain controller, the sign-in group, the service account.
        // What is missing is said in the same words as at the top of the tab.
        if (d.what === 'ad') {
          const problems = connectionsAnswer().problems;
          if (problems.signIn) return { error: problems.signIn };
          if (problems.directory) {
            await testSignIn(ldapCfg, null);
            return { error: problems.directory };
          }
          const look = { dn: dir.serviceDn, password: dir.servicePassword };
          let r;
          try { r = await testSignIn(ldapCfg, look); }
          catch (e) {
            if (look && e instanceof Error && /InvalidCredentials|invalid credentials|52e/i.test(e.name + ' ' + e.message)) return { error: 'The domain controller answers, but the service account cannot sign in: check its login and password.' };
            throw e;
          }
          if (r.group === false) return { error: 'The domain controller answers and the service account signs in, but there is no sign-in group by that name: check it.' };
          const u = await testService(dir);
          return u.users ? { message: 'All set: the domain controller answers, the service account signs in and reads users, and the sign-in group is there.' }
            : { error: 'The service account signs in, but sees no users with a mail address: check where users and groups are read (Advanced).' };
        }
        if (d.what === 'directory') {
          const r = await testService(dir);
          return r.users ? { message: 'The service account signs in and reads users.' } : { error: 'The service account signs in, but sees no users with a mail address there: check where users and groups are read.' };
        }
        if (d.what === 'graph') {
          if (!live().graph) return { error: connectionsAnswer().problems.graph };
          await live().graph.token();
          const st = settings.get();
          if (!st.spFolder) return { message: 'The server signs in to Microsoft 365. Choose the SharePoint folder in the setup assistant to check the site too.' };
          try { await settings.target().exists(''); } catch (e) {
            if (e.status === 401 || e.status === 403) return { error: 'The server signs in to Microsoft 365, but may not reach the site yet (Sites.Selected): run the setup script, or grant it as the setup assistant shows.' };
            throw e;
          }
          return { message: 'The server signs in to Microsoft 365 and reaches the site of the signatures.' };
        }
        if (d.what === 'pages') {
          if (!addin.pagesEnabled(pages)) return { error: connectionsAnswer().problems.pages };
          const r = await addin.pagesTest(pages, fetchFn);
          return { message: r.project ? 'The token works; the Pages project is there: ' + r.url : 'The token works; the Pages project is made on the first upload.' };
        }
        return { error: 'unknown' };
      } catch (e) { return { error: e.message }; }
    },
    // The server's certificate for entering by hand: the one in use, or a new one (in use from
    // now), to upload to its app in Entra. Only the public half.
    'connections/cert': (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      if (!connections.keyOk()) return { error: 'The server has no key to keep the certificate\'s key with (secrets/app_key): run sudo ./deploy-nobs-osign.sh again.' };
      let c = certs.active();
      if (!c) { const p = certs.pending(); certs.activate(p.thumbprint); configure(); c = certs.active(); }
      return { der: c.certDer, thumbprint: c.thumbprint, expires: c.expires };
    },
    // The domain, read from the domain controller (its RootDSE): "Detect" in Server connections.
    'connections/detect': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      try { return await detect(ldapCfg); } catch (e) { return { error: e.message }; }
    },
    // The server's connections (Settings → Server connections), for the local account admin.
    'connections/get': (d, s) => (adminOnly(s) ? { error: adminOnly(s) } : connectionsAnswer()),
    'connections/save': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      const r = connections.save({ values: d.values, secrets: d.secrets && typeof d.secrets === 'object' ? d.secrets : {}, ca: d.ca });
      if (r.error) return r;
      // Back to a client secret: the certificate goes (it stays on the app until removed there).
      if (d.removeCert === true) certs.remove();
      configure();
      await addinUrlFromPages();
      log.info('connections: changed by ' + s.user.login);
      return connectionsAnswer();
    },
    // A sharing link: the folder it leads to (its own address), for the setup assistant.
    // The assistant's and the connections': the local account admin's, as setting the folder is.
    'sp/share': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      const link = String(d.url || '').trim();
      if (!/^https:\/\/[a-z0-9-]+(-my)?\.sharepoint\.com\/:[a-z]:\//i.test(link)) return { error: 'That is not a sharing link of SharePoint or OneDrive.' };
      if (!live().graph) return { error: 'The server cannot open sharing links before publishing to SharePoint is set up.' };
      let r;
      try { r = await live().graph.share(link); }
      catch (e) {
        if (e.status === 401 || e.status === 403) return { error: 'The server may not open sharing links (its app registration may only reach the sites granted to it).' };
        if (e.status === 404) return { error: 'This sharing link leads nowhere: it was removed or has expired.' };
        return { error: e.message };
      }
      if (!r.folder) return { error: 'This link leads to a file (' + r.name + '), not to a folder.' };
      return { webUrl: r.webUrl, name: r.name };
    },
    // Whether the server reaches a SharePoint folder, before it is saved (the setup assistant).
    'sp/check': async (d, who) => {
      if (adminOnly(who)) return { error: adminOnly(who) };
      const s = Object.assign(settings.get(), { spSite: String(d.spSite || 'root').trim(), spFolder: String(d.spFolder || '').trim() });
      if (!s.spFolder) return { error: 'Name the folder in the site, like OSign.' };
      const err = checkSettings(s); if (err) return { error: err };
      if (!live().graph) return { error: graphProblem(graphCfg) || 'Not checked yet: the setup script sets up SharePoint (step «Register the app in Entra ID»). Paste the address now; it is checked then.' };
      const t = targetFor(s);
      try { return { found: await t.exists(''), photos: await t.exists('photos'), where: t.describe() }; }
      catch (e) {
        // The usual two, in words the assistant can show next to its instructions.
        if (e.status === 401 || e.status === 403) return { error: 'The server may not reach this site yet: its app registration needs write access to it (Sites.Selected).' };
        if (e.status === 404) return { error: 'There is no such SharePoint site. Check the address.' };
        return { error: e.message };
      }
    }
  };
  return { routes };
}

module.exports = { createConnectionRoutes };
