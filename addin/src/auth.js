/*
 * The user's token for Microsoft Graph, through Outlook's own sign-in (nested app authentication):
 * no password is asked. The app registration is filled in when the add-in's files are made (OSIGN_SETUP.clientId);
 * its ID is not a secret, and nothing else of a company is in here.
 */
import { createNestablePublicClientApplication } from '@azure/msal-browser';

/* global Office, OSIGN_SETUP */
let pca = null;

export async function getGraphToken(interactive, scopes) {
  if (!OSIGN_SETUP.clientId) throw new Error('These add-in files have no app registration (client ID).');
  if (!(Office.context.requirements && Office.context.requirements.isSetSupported('NestedAppAuth', '1.1'))) {
    throw new Error('This Outlook cannot sign in add-ins (nested app authentication). Update Outlook.');
  }
  if (!pca) {
    pca = createNestablePublicClientApplication({
      // Without a tenant: the user's own (the app registration is then for any organization).
      auth: { clientId: OSIGN_SETUP.clientId, authority: 'https://login.microsoftonline.com/' + (OSIGN_SETUP.tenantId || 'organizations') }
    });
  }
  const app = await pca;
  const request = { scopes: scopes && scopes.length ? scopes : ['User.Read'] };
  try {
    return (await app.acquireTokenSilent(request)).accessToken;
  } catch (e) {
    if (!interactive) throw e;
    return (await app.acquireTokenPopup(request)).accessToken;
  }
}
