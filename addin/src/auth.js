/*
 * The user's token for Microsoft Graph, through Outlook's own sign-in (nested app authentication):
 * no password is asked, and the app registration in the published configuration is the one used.
 */
import { createNestablePublicClientApplication } from '@azure/msal-browser';

/* global Office */
let pca = null, pcaFor = '';

export async function getGraphToken(cfg, interactive, scopes) {
  const auth = (cfg && cfg.auth) || {};
  if (!auth.clientId) throw new Error('The add-in has no app registration set up yet.');
  if (!(Office.context.requirements && Office.context.requirements.isSetSupported('NestedAppAuth', '1.1'))) {
    throw new Error('This Outlook cannot sign in add-ins (nested app authentication). Update Outlook.');
  }
  const key = auth.clientId + '|' + (auth.tenantId || '');
  if (!pca || pcaFor !== key) {
    pca = await createNestablePublicClientApplication({
      auth: { clientId: auth.clientId, authority: 'https://login.microsoftonline.com/' + (auth.tenantId || 'organizations') }
    });
    pcaFor = key;
  }
  const request = { scopes: scopes && scopes.length ? scopes : ['User.Read'] };
  try {
    return (await pca.acquireTokenSilent(request)).accessToken;
  } catch (e) {
    if (!interactive) throw e;
    return (await pca.acquireTokenPopup(request)).accessToken;
  }
}
