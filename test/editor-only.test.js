// Run with: node --test
//
// Hiding a button is not access control.
//
// `canEdit` was checked in exactly one handler — the cell editor — and the
// dashboard hid the rest of the controls from people who are not editors. That
// hiding was doing the whole job, and it does not do any of it: the buttons
// post to ordinary endpoints, and anyone with an account could reach them from
// the console, from curl, or from a page left open since before they were
// taken off the list.
//
// What that left open to every reader: the mailbox sync and the archive tidy
// (which move files in SharePoint), applying a statement's amounts, repairing
// totals, and importing invoices from a folder — all writes to the workbook
// that is the system of record.
//
// So every writing route goes through one gate, and these tests drive it with
// a reader's session rather than reading the source and hoping.

const test = require('node:test');
const assert = require('node:assert');

process.env.SESSION_SECRET = 'test-secret-for-editor-gate';
process.env.ALLOWED_EMAILS = 'boss@sarasanalytics.com,reader@sarasanalytics.com';
process.env.EDITOR_EMAILS = 'boss@sarasanalytics.com';

const { sign } = require('../lib/session');
const { requireEditor } = require('../lib/require-editor');

const cookieFor = email => `session=${sign({ email, name: email, exp: Date.now() + 3600e3 })}`;

function fakeRes() {
  const r = { code: null, body: null, headers: {} };
  r.status = c => { r.code = c; return r; };
  r.json = b => { r.body = b; return r; };
  r.setHeader = (k, v) => { r.headers[k] = v; };
  r.end = () => r;
  return r;
}
const reqWith = cookie => ({ method: 'POST', headers: { cookie }, body: {}, query: {} });

test('an editor is let through', () => {
  const res = fakeRes();
  const session = requireEditor(reqWith(cookieFor('boss@sarasanalytics.com')), res);
  assert.ok(session, 'the editor must get a session back');
  assert.strictEqual(session.email, 'boss@sarasanalytics.com');
  assert.strictEqual(res.code, null, 'and nothing answered on their behalf');
});

test('a reader with a valid session is refused, not merely un-shown the button', () => {
  const res = fakeRes();
  const out = requireEditor(reqWith(cookieFor('reader@sarasanalytics.com')), res);
  assert.strictEqual(out, null, 'the handler must stop');
  assert.strictEqual(res.code, 403, 'signed in is not the same as allowed to write');
  assert.match(res.body.error, /not/i);
});

test('no session at all is 401, not 403', () => {
  for (const cookie of [undefined, '', 'session=rubbish', 'session=a.b']) {
    const res = fakeRes();
    assert.strictEqual(requireEditor(reqWith(cookie), res), null);
    assert.strictEqual(res.code, 401, `"${cookie}" is nobody, not somebody refused`);
  }
});

test('a forged session does not become an editor', () => {
  // The payload says the boss; the signature was made with another key. Built
  // by hand rather than with sign(), which would use the real one.
  const { createHmac } = require('node:crypto');
  const b64url = buf => Buffer.from(buf).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const payload = b64url(JSON.stringify({ email: 'boss@sarasanalytics.com', exp: Date.now() + 3600e3 }));
  const sig = b64url(createHmac('sha256', 'not-the-real-secret').update(payload).digest());

  const res = fakeRes();
  assert.strictEqual(requireEditor(reqWith(`session=${payload}.${sig}`), res), null,
    'a signature made with the wrong key must not pass');
  assert.strictEqual(res.code, 401);
});

test('an expired editor session is not an editor', () => {
  const res = fakeRes();
  const stale = `session=${sign({ email: 'boss@sarasanalytics.com', exp: Date.now() - 1000 })}`;
  assert.strictEqual(requireEditor(reqWith(stale), res), null);
  assert.strictEqual(res.code, 401);
});

// --- the routes that were open ----------------------------------------------

const fs = require('node:fs');
const path = require('node:path');
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

test('every writing route calls the gate', () => {
  for (const f of ['api/amounts.js', 'api/invoices/import.js', 'api/invoices/sync-cron.js']) {
    const src = read(f);
    assert.match(src, /require\(['"][./]*(?:\.\.\/)*lib\/require-editor['"]\)/, `${f} must load the gate`);
    assert.match(src, /requireEditor\(req, res\)/, `${f} must call it`);
  }
});

test('the amounts route is gated before it dispatches, so a new action is gated too', () => {
  const src = read('api/amounts.js');
  const gate = src.indexOf('requireEditor(req, res)');
  const dispatch = src.indexOf('const handler = ROUTES[action]');
  assert.ok(gate > -1 && dispatch > gate, 'the gate must come before the lookup');
});

test('the scheduled run still authenticates with CRON_SECRET, not a session', () => {
  // The nightly job has no browser session. Gating it on the editor list
  // without keeping that path would silently stop the cron.
  const src = read('api/invoices/sync-cron.js');
  assert.match(src, /const isCron = !!secret && auth === `Bearer \$\{secret\}`;/);
  assert.match(src, /if \(!isCron\) \{\s*\n\s*session = requireEditor\(req, res\);/);
});

test('the page hides the controls, and says that is not what stops anyone', () => {
  const html = read('index.html');
  assert.match(html, /body:not\(\.can-edit\) \.admin-only\{display:none!important\}/);
  for (const marker of [
    /class="btn ghost maint-toggle admin-only"/,      // the maintenance expander
    /class="toolbar admin-only"/,                      // Check Totals … Refresh now
    /class="ktiles-actions admin-only"/,               // Fetch invoices / Import invoices
    /class="nav-item admin-only" data-page="questions"/,
    /class="nav-item admin-only" data-page="rules"/,
  ]) assert.match(html, marker, `${marker} should be editor-only`);
});

test('a reader does not ask for the question queue', () => {
  // It is a 403 now, and a reader has nothing to answer. The call moved into
  // the auth block because CAN_EDIT is undefined when the script first runs.
  const html = read('index.html');
  assert.match(html, /if \(me\.canEdit && typeof loadPending === 'function'\) loadPending\(true\);/);
  assert.doesNotMatch(html, /\n};\nloadPending\(true\);/, 'the unconditional boot call must be gone');
});
