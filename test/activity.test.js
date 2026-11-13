// Run with: node --test
//
// "I also need a feature in admin login for see who is using the dashboard and
// what activities they did."
//
// Two logs, one feed. The sheet writer already kept `_amount-log.json`: every
// cell this app has written, with who did it and what the cell held before.
// What it could not tell you is who merely signed in, or who ran the mailbox
// sync — those leave no cell behind. Those go in `_activity-log.json`, and the
// reader merges the two.
//
// Writing the log is instrumentation, never a step in the work: a sync that
// cannot write its own entry has still run, and a person whose sign-in could
// not be recorded is still signed in. So the writer swallows everything.

const test = require('node:test');
const assert = require('node:assert');

const graph = require('../lib/graph');

let stored = null;
let failWrite = false;
graph.resolveArchiveRoot = async () => ({ path: 'Invoices', itemId: 'root-1' });
graph.readJsonFile = async (_t, _d, path) => {
  if (path.endsWith('_activity-log.json')) return stored;
  throw new Error('404');
};
graph.writeJsonFile = async (_t, _d, path, obj) => {
  if (failWrite) throw new Error('Graph said no');
  if (path.endsWith('_activity-log.json')) stored = obj;
  return {};
};

// lib/amounts/activity destructures these when it loads, so they are stubbed
// once, here, and read mutable state — reassigning them later would not take.
const spendSheet = require('../lib/spend-sheet');
let logEntries = [];
let logThrows = false;
spendSheet.readLog = async () => {
  if (logThrows) throw new Error('404');
  return { entries: logEntries };
};
graph.getGraphToken = async () => 'tok';
graph.resolveDriveId = async () => 'drive-1';
process.env.TARGET_USER_UPN = 'owner@example.com';

const activity = require('../lib/activity');
const activityFeed = require('../lib/amounts/activity');

const askFeed = () => new Promise((resolve) => {
  const res = { setHeader(){}, status(c){ this.code = c; return this; }, json(b){ resolve({ code: this.code, body: b }); } };
  activityFeed({ method: 'GET', query: {}, headers: {} }, res);
});

test.beforeEach(() => { stored = null; failWrite = false; logEntries = []; logThrows = false; });

test('an event is recorded with who, what and when', async () => {
  const ok = await activity.recordActivity('tok', 'drive-1', 'sign-in',
    'Reader@Sarasanalytics.com', 'Signed in', 'Reader Name');
  assert.strictEqual(ok, true);
  assert.strictEqual(stored.entries.length, 1);
  const e = stored.entries[0];
  assert.strictEqual(e.kind, 'sign-in');
  assert.strictEqual(e.by, 'reader@sarasanalytics.com', 'one spelling per person, or the feed splits them');
  assert.strictEqual(e.what, 'Signed in');
  assert.ok(Date.parse(e.at), 'and a timestamp that parses');
});

test('a log that cannot be written does not break what the person was doing', async () => {
  // The caller is mid sign-in or mid sync. Neither should end differently
  // because a JSON file in SharePoint would not save.
  failWrite = true;
  const ok = await activity.recordActivity('tok', 'drive-1', 'sync', 'a@b.com', 'Fetched invoices', '');
  assert.strictEqual(ok, false, 'it reports the miss');
  // and did not throw, which is the whole point
});

test('a missing log file reads as no history, not as an error', async () => {
  stored = null;
  assert.deepStrictEqual(await activity.readActivity('tok', 'drive-1'), []);
});

test('the file cannot grow without bound', async () => {
  stored = { entries: Array.from({ length: activity.MAX_ENTRIES + 50 }, (_, i) => ({ at: `2026-01-01T00:00:${String(i % 60).padStart(2,'0')}.000Z`, kind: 'sign-in', by: 'a@b.com', what: `#${i}` })) };
  await activity.recordActivity('tok', 'drive-1', 'sign-in', 'a@b.com', 'newest', '');
  assert.strictEqual(stored.entries.length, activity.MAX_ENTRIES, 'it is read on every page open');
  assert.strictEqual(stored.entries[stored.entries.length - 1].what, 'newest', 'and the newest is the one kept');
});

test('a long field is cut rather than stored whole', async () => {
  await activity.recordActivity('tok', 'drive-1', 'sync', 'a@b.com', 'x'.repeat(5000), 'y'.repeat(5000));
  assert.ok(stored.entries[0].what.length <= 300);
  assert.ok(stored.entries[0].detail.length <= 1000);
});

// --- the merged feed --------------------------------------------------------

test('the feed merges sign-ins with the cell writes, newest first', async () => {
  stored = { entries: [
    { at: '2026-11-10T08:00:00.000Z', kind: 'sign-in', by: 'reader@x.com', what: 'Signed in', detail: '' },
    { at: '2026-11-10T09:00:00.000Z', kind: 'sync', by: 'boss@x.com', what: 'Fetched invoices from the mailbox', detail: 'mode=mail' },
  ] };
  logEntries = [
    { at: '2026-11-10T08:30:00.000Z', by: 'boss@x.com', sheet: 'Spendings', attribution: 'dashboard-edit',
      cells: [{ app: 'AWS', month: '2026-10', address: 'O4', before: 11, after: 2401.67 }] },
    { at: '2026-11-10T07:00:00.000Z', by: 'invoice-sync', sheet: 'Spendings', attribution: 'invoice',
      cells: [{ app: 'DBT Cloud', month: '2026-11', address: 'P9', before: null, after: 425 }] },
  ];
  const out = await askFeed();

  assert.strictEqual(out.code, 200);
  const feed = out.body.entries;
  assert.deepStrictEqual(feed.map(e => e.at), [
    '2026-11-10T09:00:00.000Z',
    '2026-11-10T08:30:00.000Z',
    '2026-11-10T08:00:00.000Z',
    '2026-11-10T07:00:00.000Z',
  ], 'newest first, both sources interleaved');

  const edit = feed.find(e => e.at === '2026-11-10T08:30:00.000Z');
  assert.strictEqual(edit.kind, 'edit', 'a person wrote it');
  assert.match(edit.what, /AWS 2026-10/);
  assert.match(edit.what, /\$11\.00 → \$2,401\.67/, 'what it was and what it became');

  const bySync = feed.find(e => e.at === '2026-11-10T07:00:00.000Z');
  assert.strictEqual(bySync.kind, 'sync-write', 'the job wrote it, not a person');
  assert.strictEqual(bySync.by, 'the invoice sync');
  assert.match(bySync.what, /nothing → \$425\.00/, 'an empty cell reads as nothing, not as $0.00');
});

test('the people list counts only people, and remembers when each was last here', async () => {
  stored = { entries: [
    { at: '2026-11-10T08:00:00.000Z', kind: 'sign-in', by: 'reader@x.com', what: 'Signed in' },
    { at: '2026-11-09T08:00:00.000Z', kind: 'sign-in', by: 'reader@x.com', what: 'Signed in' },
  ] };
  logEntries = [
    { at: '2026-11-10T07:00:00.000Z', by: 'invoice-sync', cells: [{ app: 'A', month: '2026-01', before: null, after: 1 }] },
  ];
  const { body: out } = await askFeed();

  assert.deepStrictEqual(out.people.map(p => p.by), ['reader@x.com'], 'a job is not a person');
  assert.strictEqual(out.people[0].signIns, 2);
  assert.strictEqual(out.people[0].lastAt, '2026-11-10T08:00:00.000Z');
});

test('a missing audit trail still yields the sign-ins', async () => {
  stored = { entries: [{ at: '2026-11-10T08:00:00.000Z', kind: 'sign-in', by: 'a@x.com', what: 'Signed in' }] };
  logThrows = true;
  const { body: out } = await askFeed();
  assert.strictEqual(out.entries.length, 1);
  assert.strictEqual(out.counts.writes, 0);
});

// --- the wiring -------------------------------------------------------------

const fs = require('node:fs');
const path = require('node:path');
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

test('a sign-in is recorded before the redirect, not after it', () => {
  // Work started after a response has been sent may not survive on a
  // serverless function, so the entry would be lost exactly when someone
  // wanted to know who had been in.
  const src = read('api/auth/callback.js');
  const rec = src.indexOf("recordActivityStandalone('sign-in'");
  const redirect = src.indexOf("res.redirect(302, '/')");
  assert.ok(rec > -1, 'the sign-in must be recorded');
  assert.ok(redirect > rec, 'and recorded before the redirect');
  assert.match(src, /await recordActivityStandalone/, 'awaited, or it races the response');
});

test('the invoice jobs record what was run and who ran it', () => {
  const src = read('api/invoices/sync-cron.js');
  assert.match(src, /if \(session && RUN_LABEL\[mode\]\)/, 'only a person, not the nightly cron');
  assert.match(src, /recordActivity\(token, targetDriveId, 'sync', session\.email/);
  for (const mode of ['mail', 'periods-apply', 'pending-resolve', 'rules-save'])
    assert.match(src, new RegExp(`'?${mode}'?:`), `${mode} needs a label`);
  // A read is not an activity: logging every page open would bury the feed.
  const labels = src.slice(src.indexOf('const RUN_LABEL'), src.indexOf('};', src.indexOf('const RUN_LABEL')));
  assert.doesNotMatch(labels, /\bpending:/, 'loading the question list is a read');
  assert.doesNotMatch(labels, /\binspect:/);
});

test('the activity route is reachable only as an editor', () => {
  // It says who has been in the dashboard, which is not a reader's business.
  const src = read('api/amounts.js');
  assert.match(src, /const ROUTES = \{[^}]*activity[^}]*\}/, 'it hangs off the gated route');
  assert.ok(src.indexOf('requireEditor(req, res)') < src.indexOf('const handler = ROUTES[action]'),
    'and the gate runs before the dispatch');
});

test('the page is editor-only too, and in the router', () => {
  const html = read('index.html');
  assert.match(html, /<button class="nav-item admin-only" data-page="activity">/);
  assert.match(html, /<div class="card admin-only" id="activityCard" data-page="activity"/);
  assert.match(html, /activity:\s*\{ title: 'Activity'/);
  assert.match(html, /if \(name === 'activity' && !actData && !actLoading\) loadActivity\(false\);/);
});
