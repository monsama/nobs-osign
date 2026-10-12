/*
 * The user's token for Microsoft Graph, through Outlook's own sign-in (nested app authentication):
 * no password is asked. The app registration is set when the add-in is built (OSIGN_CLIENT_ID);
 * its ID is not a secret, and nothing else of a company is in here.
 */
import { createNestablePublicClientApplication } from '@azure/msal-browser';

/* global Office, __CLIENT_ID__, __TENANT_ID__ */
let pca = null;

export async function getGraphToken(interactive, scopes) {
  if (!__CLIENT_ID__) throw new Error('This build of the add-in has no app registration (OSIGN_CLIENT_ID).');
  if (!(Office.context.requirements && Office.context.requirements.isSetSupported('NestedAppAuth', '1.1'))) {
    throw new Error('This Outlook cannot sign in add-ins (nested app authentication). Update Outlook.');
  }
  if (!pca) {
    pca = createNestablePublicClientApplication({
      // Without a tenant: the user's own (the app registration is then for any organization).
      auth: { clientId: __CLIENT_ID__, authority: 'https://login.microsoftonline.com/' + (__TENANT_ID__ || 'organizations') }
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
