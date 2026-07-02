const { sign, parseCookies, decodeIdTokenPayload } = require('../../lib/session');

const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

function htmlError(res, status, message) {
  res.status(status).send(`<!DOCTYPE html><html><body style="font-family:sans-serif;background:#0f1220;color:#e7e9f3;padding:40px">
    <h2>Sign-in failed</h2><p>${message}</p><p><a href="/api/auth/login" style="color:#6c8cff">Try again</a></p>
  </body></html>`);
}

module.exports = async (req, res) => {
  try {
    const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, PUBLIC_APP_URL, ALLOWED_EMAILS } = process.env;
    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET || !PUBLIC_APP_URL) {
      return htmlError(res, 500, 'Login is not fully configured on the server.');
    }

    const url = new URL(req.url, `https://${req.headers.host}`);
    const code = url.searchParams.get('code');
    const state = url.searchParams.get('state');
    const errorParam = url.searchParams.get('error');
    if (errorParam) return htmlError(res, 400, `Microsoft returned an error: ${errorParam} — ${url.searchParams.get('error_description') || ''}`);

    const cookies = parseCookies(req.headers.cookie);
    if (!code || !state || !cookies.oauth_state || state !== cookies.oauth_state) {
      return htmlError(res, 400, 'Invalid or expired sign-in request. Please try again.');
    }

    const redirectUri = `${PUBLIC_APP_URL.replace(/\/$/, '')}/api/auth/callback`;
    const tokenRes = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        code,
        redirect_uri: redirectUri,
        scope: 'openid profile email',
      }),
    });
    if (!tokenRes.ok) {
      const text = await tokenRes.text();
      return htmlError(res, 502, `Token exchange failed: ${text.slice(0, 300)}`);
    }
    const tokenJson = await tokenRes.json();
    const claims = decodeIdTokenPayload(tokenJson.id_token);
    const email = String(claims.preferred_username || claims.email || '').toLowerCase().trim();
    const name = claims.name || email;

    const allowed = (ALLOWED_EMAILS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    if (!allowed.length || !allowed.includes(email)) {
      return htmlError(res, 403, `Signed in as <b>${email}</b>, but this account isn't on the dashboard's access list. Ask the dashboard owner to add you.`);
    }

    const session = sign({ email, name, exp: Date.now() + SESSION_TTL_MS });
    res.setHeader('Set-Cookie', [
      `session=${session}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${Math.floor(SESSION_TTL_MS / 1000)}`,
      `oauth_state=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0`,
    ]);
    res.writeHead(302, { Location: '/' });
    res.end();
  } catch (err) {
    htmlError(res, 500, (err && err.message) || String(err));
  }
};
