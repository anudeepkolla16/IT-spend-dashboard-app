// Run with: node --test
//
// Graph throttles, and a run of this app makes a few hundred calls to it. Every
// 429 used to become a line in the run summary beside the invoice it lost, and
// every listing stopped at 200 children without saying so. These are the tests
// for both.

const test = require('node:test');
const assert = require('node:assert');
const { graphFetch, graphListAll, listFilesRecursive } = require('../lib/graph');
const { folderChildren } = require('../lib/mail-sync');

// Graph's own shape, enough of it: a status, headers.get, and a json/text body.
// Retry-After "0" keeps the tests instant while still exercising the header path.
const reply = (status, body, retryAfter) => ({
  ok: status >= 200 && status < 300,
  status,
  headers: { get: (h) => (String(h).toLowerCase() === 'retry-after' && retryAfter !== undefined ? String(retryAfter) : null) },
  json: async () => body,
  text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
});

// Replaces global.fetch with a scripted sequence of replies (or thrown errors).
function stub(sequence) {
  const original = global.fetch;
  const calls = [];
  let i = 0;
  global.fetch = async (url) => {
    calls.push(String(url));
    const next = sequence[Math.min(i, sequence.length - 1)];
    i++;
    if (next instanceof Error) throw next;
    return typeof next === 'function' ? next(String(url)) : next;
  };
  return { calls, restore: () => { global.fetch = original; } };
}

test('a throttled call is retried, not reported as a failure', async () => {
  const s = stub([reply(429, 'slow down', 0), reply(200, { ok: true })]);
  try {
    const res = await graphFetch('https://graph/x', {});
    assert.strictEqual(res.status, 200);
    assert.strictEqual(s.calls.length, 2);
  } finally { s.restore(); }
});

test('the 5xx family is retried too', async () => {
  const s = stub([reply(503, 'unavailable', 0), reply(504, 'gateway', 0), reply(200, { ok: true })]);
  try {
    assert.strictEqual((await graphFetch('https://graph/x', {})).status, 200);
    assert.strictEqual(s.calls.length, 3);
  } finally { s.restore(); }
});

test('an answer is an answer — 404 and 403 are never retried', async () => {
  for (const status of [404, 403, 400, 500]) {
    const s = stub([reply(status, 'no')]);
    try {
      assert.strictEqual((await graphFetch('https://graph/x', {})).status, status);
      assert.strictEqual(s.calls.length, 1, `${status} must not be retried`);
    } finally { s.restore(); }
  }
});

test('retrying is bounded, and the last response is handed back to be reported', async () => {
  const s = stub([reply(429, 'still throttled', 0)]);
  try {
    const res = await graphFetch('https://graph/x', {});
    assert.strictEqual(res.status, 429);       // the caller reports a real failure
    assert.strictEqual(s.calls.length, 4);     // and it gave up rather than looping
  } finally { s.restore(); }
});

test('a network fault is retried, and rethrown if it never clears', async () => {
  const s1 = stub([new Error('socket hang up'), reply(200, { ok: true })]);
  try {
    assert.strictEqual((await graphFetch('https://graph/x', {})).status, 200);
  } finally { s1.restore(); }

  const s2 = stub([new Error('ECONNRESET')]);
  try {
    await assert.rejects(() => graphFetch('https://graph/x', {}), /ECONNRESET/);
  } finally { s2.restore(); }
});

test('the run deadline is never spent sleeping', async () => {
  // A one-second backoff with no time left to serve it: hand the throttle back
  // now, so the run finishes and the next one picks the work up.
  const s = stub([reply(429, 'slow down', 1)]);
  try {
    const res = await graphFetch('https://graph/x', {}, { deadline: Date.now() - 1 });
    assert.strictEqual(res.status, 429);
    assert.strictEqual(s.calls.length, 1);
  } finally { s.restore(); }
});

// --- Paging ---------------------------------------------------------------

test('a listing longer than one page is read whole', async () => {
  const s = stub([
    reply(200, { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph/page2' }),
    reply(200, { value: [{ id: 'b' }] }),
  ]);
  try {
    const items = await graphListAll('tok', 'https://graph/page1', 'test listing');
    assert.deepEqual(items.map(i => i.id), ['a', 'b']);
    assert.deepEqual(s.calls, ['https://graph/page1', 'https://graph/page2']);
  } finally { s.restore(); }
});

test('a failed listing throws rather than returning half an answer', async () => {
  const s = stub([reply(403, 'accessDenied')]);
  try {
    await assert.rejects(() => graphListAll('tok', 'https://graph/x', 'test listing'), /403/);
  } finally { s.restore(); }
});

test('listFilesRecursive follows the next page of a big folder', async () => {
  // Two hundred invoices in one folder used to be the ceiling, silently: the
  // 201st simply did not exist as far as the checklist and the totals knew.
  const s = stub([
    reply(200, { value: [{ id: 'f1', name: 'a.pdf', file: {} }], '@odata.nextLink': 'https://graph/next' }),
    reply(200, { value: [{ id: 'f2', name: 'b.pdf', file: {} }] }),
  ]);
  try {
    const files = await listFilesRecursive('tok', 'drive1', 'folder1');
    assert.deepEqual(files.map(f => f.name), ['a.pdf', 'b.pdf']);
  } finally { s.restore(); }
});

// --- What a failed listing must not be mistaken for ------------------------

test('a folder that does not exist yet lists as empty', async () => {
  const s = stub([reply(404, 'itemNotFound')]);
  try {
    const out = await folderChildren('tok', 'drive1', 'Procurment bills/Cumul/Sep-26');
    assert.strictEqual(out.files.size, 0);
    assert.deepEqual(out.folders, []);
    assert.strictEqual(out.missing, true);
  } finally { s.restore(); }
});

test('a listing that fails is an error, not an empty folder', async () => {
  // Swallowed, this is how a duplicate month folder gets made: "no month folders
  // here" means the run creates Aug-26 beside the existing Aug, and the month's
  // invoices end up split across two folders, each totalling short.
  const s = stub([reply(500, 'internalServerError')]);
  try {
    await assert.rejects(() => folderChildren('tok', 'drive1', 'Procurment bills/Cumul'), /could not list/);
  } finally { s.restore(); }
});

test('folderChildren pages, and splits files from subfolders', async () => {
  const s = stub([
    reply(200, { value: [{ name: 'Aug-26', folder: {} }, { name: 'a.pdf', file: {} }], '@odata.nextLink': 'https://graph/next' }),
    reply(200, { value: [{ name: 'Sep-26', folder: {} }, { name: 'b.pdf', file: {} }] }),
  ]);
  try {
    const out = await folderChildren('tok', 'drive1', 'Procurment bills/Cumul');
    assert.deepEqual(out.folders, ['Aug-26', 'Sep-26']);
    assert.deepEqual([...out.files].sort(), ['a.pdf', 'b.pdf']);
  } finally { s.restore(); }
});
