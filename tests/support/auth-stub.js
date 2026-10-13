// Stands in for addin/src/auth.js in the tests: Outlook's sign-in (nested app authentication) is
// not there outside Outlook. The token names its scopes, so the stand-in Graph can check them.
export async function getGraphToken(interactive, scopes) {
  if (globalThis.__osignSignInFails) throw new Error('sign-in failed');
  return 'token:' + (scopes || []).join(' ');
}
