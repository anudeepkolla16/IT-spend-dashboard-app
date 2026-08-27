const crypto = require('crypto');

module.exports = async (req, res) => {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, PUBLIC_APP_URL } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !PUBLIC_APP_URL) {
    res.status(500).send('Login is not configured (missing AZURE_TENANT_ID / AZURE_CLIENT_ID / PUBLIC_APP_URL).');
    return;
  }

  // The sign-in cookie is set on whatever host this runs on, but Microsoft
  // always sends the user back to PUBLIC_APP_URL. Started from a preview
  // deployment, those are two different hosts: the browser never carries the
  // cookie across, and the callback rejects the sign-in as "invalid or expired"
  // — a full round trip through Microsoft to reach a dead end that says nothing
  // about what actually went wrong. Say it here instead, before that trip.
  const canonicalHost = (() => { try { return new URL(PUBLIC_APP_URL).host; } catch (_) { return ''; } })();
  const host = String(req.headers.host || '');
  if (canonicalHost && host && host !== canonicalHost) {
    res.status(400).send(`<!DOCTYPE html><html><body style="font-family:sans-serif;background:#0f1220;color:#e7e9f3;padding:40px">
      <h2>Sign-in is not set up for this deployment</h2>
      <p>You are on <b>${host}</b>, but sign-in returns to <b>${canonicalHost}</b> — the address this app is registered under —
      so the browser cannot carry the sign-in across and it would fail on the way back.</p>
      <p><a href="https://${canonicalHost}/api/auth/login" style="color:#6c8cff">Sign in at ${canonicalHost}</a></p>
      <p style="color:#8b90a8;font-size:13px">To sign in to a preview deployment instead, set <code>PUBLIC_APP_URL</code> for Vercel's
      Preview environment to this host and add <code>https://${host}/api/auth/callback</code> to the app registration's redirect URIs.</p>
    </body></html>`);
    return;
  }

  const state = crypto.randomBytes(16).toString('hex');
  const redirectUri = `${PUBLIC_APP_URL.replace(/\/$/, '')}/api/auth/callback`;

  const authorizeUrl = new URL(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/authorize`);
  authorizeUrl.searchParams.set('client_id', AZURE_CLIENT_ID);
  authorizeUrl.searchParams.set('response_type', 'code');
  authorizeUrl.searchParams.set('redirect_uri', redirectUri);
  authorizeUrl.searchParams.set('response_mode', 'query');
  authorizeUrl.searchParams.set('scope', 'openid profile email');
  authorizeUrl.searchParams.set('state', state);

  res.setHeader('Set-Cookie', `oauth_state=${state}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`);
  res.redirect(302, authorizeUrl.toString());
};
