// Run with: node --test
//
// Who gets in, and what a refusal tells the person refused. The two refusals
// used to wear one message — "this account isn't on the access list" — so an
// empty ALLOWED_EMAILS was indistinguishable from a misspelled address, and
// nothing was logged to settle it from the other end.

const test = require('node:test');
const assert = require('node:assert');

process.env.AZURE_TENANT_ID = 'tenant-1';
process.env.AZURE_CLIENT_ID = 'client-1';
process.env.AZURE_CLIENT_SECRET = 'secret-1';
process.env.PUBLIC_APP_URL = 'https://dash.example.com';
process.env.SESSION_SECRET = 'test-secret';

const { b64url, verify, parseCookies } = require('../lib/session');

// A signed-in identity, as Microsoft would hand it back.
function idToken(email) {
  const claims = { preferred_username: email, name: email };
  return `x.${b64url(Buffer.from(JSON.stringify(claims)))}.y`;
}
let issuedFor = 'nobody@example.com';
global.fetch = async () => ({
  ok: true, status: 200,
  json: async () => ({ id_token: idToken(issuedFor), access_token: 'at' }),
  text: async () => '',
});

const callback = require('../api/auth/callback');

function invoke(email, allowedEmails) {
  issuedFor = email;
  if (allowedEmails === undefined) delete process.env.ALLOWED_EMAILS;
  else process.env.ALLOWED_EMAILS = allowedEmails;
  return new Promise((resolve) => {
    const state = 'abc123';
    const req = {
      method: 'GET',
      url: `/api/auth/callback?code=the-code&state=${state}`,
      headers: { host: 'dash.example.com', cookie: `oauth_state=${state}` },
    };
    const out = { statusCode: 200, headers: {} };
    const res = {
      status(c) { out.statusCode = c; return this; },
      setHeader(k, v) { out.headers[k] = v; },
      send(body) { out.body = body; resolve(out); },
      redirect(c, to) { out.statusCode = c; out.location = to; resolve(out); },
      json(p) { out.body = p; resolve(out); },
    };
    callback(req, res);
  });
}

// console.warn/log, captured so the tests can read what the owner would see in
// the runtime logs.
let logged = [];
const realWarn = console.warn, realLog = console.log;
test.beforeEach(() => {
  logged = [];
  console.warn = (m) => logged.push(String(m));
  console.log = (m) => logged.push(String(m));
});
test.afterEach(() => { console.warn = realWarn; console.log = realLog; });

test('an address on the list is signed in, and the session says who', async () => {
  const out = await invoke('rajamma@sarasanalytics.com', 'anudeep.kolla@sarasanalytics.com, Rajamma@Sarasanalytics.com');
  assert.strictEqual(out.statusCode, 302, 'straight to the dashboard');
  assert.strictEqual(out.location, '/');
  // Case and the spaces around a comma are the owner's typing, not a rejection.
  const cookie = (out.headers['Set-Cookie'] || []).find(c => c.startsWith('session='));
  const session = verify(parseCookies(cookie.split(';')[0]).session);
  assert.strictEqual(session.email, 'rajamma@sarasanalytics.com');
  assert.match(logged.join('\n'), /signed in rajamma@sarasanalytics\.com/);
});

test('an empty list says so, and says a redeploy is what applies one', async () => {
  const out = await invoke('subha.kumar@sarasanalytics.com', '');
  assert.strictEqual(out.statusCode, 403);
  assert.match(out.body, /no access list at all/);
  assert.match(out.body, /ALLOWED_EMAILS/);
  assert.match(out.body, /redeploy/i, 'setting the variable alone changes nothing');
  assert.match(logged.join('\n'), /ALLOWED_EMAILS is empty/);
});

test('an unlisted address is told it is the address Microsoft used that must match', async () => {
  const out = await invoke('santoshi.ch@sarasanalytics.com', 'anudeep.kolla@sarasanalytics.com,rajamma@sarasanalytics.com');
  assert.strictEqual(out.statusCode, 403);
  assert.match(out.body, /santoshi\.ch@sarasanalytics\.com/, 'the person is told which address was tried');
  assert.match(out.body, /\(2 on it\)/, 'and that a list does exist, so this is a spelling question');
  assert.match(out.body, /alias/, 'the likeliest cause, named');
  // The log is enough to settle it from the owner's end, without the list in it.
  const line = logged.find(l => l.includes('refused'));
  assert.match(line, /not among the 2 addresses/);
  assert.ok(!line.includes('rajamma'), 'a refusal never prints the allowlist');
});

test('no session cookie is set for anyone refused', async () => {
  for (const list of ['', 'someone.else@example.com']) {
    const out = await invoke('outsider@example.com', list);
    assert.strictEqual(out.statusCode, 403);
    assert.strictEqual(out.headers['Set-Cookie'], undefined);
  }
});
