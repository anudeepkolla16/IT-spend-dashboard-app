const crypto = require('crypto');

module.exports = async (req, res) => {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, PUBLIC_APP_URL } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !PUBLIC_APP_URL) {
    res.status(500).send('Login is not configured (missing AZURE_TENANT_ID / AZURE_CLIENT_ID / PUBLIC_APP_URL).');
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
