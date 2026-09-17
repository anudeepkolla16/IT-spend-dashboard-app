// Run with: node --test
//
// Editing the spend sheet from the dashboard. The sheet is the system of record
// and several people can open the dashboard, so the gate is in the handler and
// these tests hold it there: hiding a button in a page stops nobody.

const test = require('node:test');
const assert = require('node:assert');

process.env.TARGET_USER_UPN = 'owner@example.com';
process.env.TARGET_FILE_PATH = 'Test/Spend.xlsx';
process.env.SESSION_SECRET = 'test-secret';

const { sign } = require('../lib/session');
const graph = require('../lib/graph');
graph.getGraphToken = async () => 'test-token';
graph.resolveDriveId = async () => 'drive-1';

const MONTH_SERIALS = [46023, 46054, 46082];               // Jan-26 … Mar-26
const MONTH_LABELS = ['Jan-26', 'Feb-26', 'Mar-26'];
const SHEET_VALUES = [
  ['Spendings', '', '', '', '', '', '', '', '', ''],
  ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Renewal data', 'Recurring/Onetime', 'FREQUENCY', 'Payment Method', ...MONTH_SERIALS],
  ['Adobe', 'Marketing', 'Bhavana', '1st of every month', 'Recurring', 'Monthly', 'US Debit Card', 37.16, 37.16, 37.16],
  ['Cumul(Luzmo)', 'Consulting', 'Ganesh', '28th of every month', 'Recurring', 'Monthly', 'US Bank', 2178.12, 2735.93, 557.28],
  ['Total', '', '', '', '', '', '', 2215.28, 2773.09, 594.44],
];
const SHEET_TEXT = SHEET_VALUES.map((row, i) => row.map((cell, j) =>
  (i === 1 && j >= 7) ? MONTH_LABELS[j - 7] : (cell == null ? '' : String(cell))));

const excel = require('../lib/excel');
excel.resolveItemId = async () => 'item-1';
excel.listWorksheets = async () => [{ id: '1', name: 'Spendings' }];
excel.readUsedRange = async () => ({ values: SHEET_VALUES, formulas: [], text: SHEET_TEXT, start: { col: 0, row: 0 }, address: 'Spendings!A1:J5' });
let sessionsOpened = 0, sessionsClosed = 0, writes = [];
excel.createSession = async () => { sessionsOpened++; return 'session-1'; };
excel.closeSession = async () => { sessionsClosed++; };
excel.writeCell = async (_t, _d, _i, sheetName, address, value) => { writes.push({ sheetName, address, value }); return {}; };

// The archive, for the lock a saved amount earns.
const FRESH_RULES = () => ({ version: 1, vendors: [{ name: 'Adobe', domains: ['adobe.com'], subject: [], app: 'Adobe' }], locks: [], locksSeeded: 99 });
let rulesFile = FRESH_RULES();
let logEntries = [];
graph.resolveArchiveRoot = async () => ({ path: 'Test/Invoices', itemId: 'arch-1', candidates: [], resolved: true, expiresAt: Infinity });
graph.readJsonFile = async (_t, _d, path) => {
  if (path.endsWith('_vendor-rules.json')) return rulesFile ? JSON.parse(JSON.stringify(rulesFile)) : null;
  if (path.endsWith('_amount-log.json')) return { entries: logEntries };
  return null;
};
graph.writeJsonFile = async (_t, _d, path, obj) => {
  if (path.endsWith('_vendor-rules.json')) rulesFile = obj;
  if (path.endsWith('_amount-log.json')) logEntries = obj.entries || [];
  return {};
};

const editHandler = require('../lib/amounts/edit');

const OWNER = 'owner@example.com';
const cookieFor = (email) => `session=${sign({ email, name: email, exp: Date.now() + 60000 })}`;

function invoke(body, cookie, method) {
  return new Promise((resolve) => {
    const req = { method: method || 'POST', body, query: {}, headers: { cookie: cookie || '' } };
    const res = {
      statusCode: 200, setHeader() {},
      status(code) { this.statusCode = code; return this; },
      json(payload) { resolve({ status: this.statusCode, body: payload }); },
    };
    editHandler(req, res);
  });
}

test.beforeEach(() => { writes = []; logEntries = []; sessionsOpened = 0; sessionsClosed = 0; rulesFile = FRESH_RULES(); delete process.env.EDITOR_EMAILS; });

// --- the gate ---------------------------------------------------------------

test('only an account on the editor list may write, whoever is signed in', async () => {
  // Signed in, on the list (it defaults to the owner of the workbook).
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, cookieFor(OWNER))).status, 200);

  // Signed in, not on the list: a reader of the dashboard is not an editor.
  const other = await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, cookieFor('someone.else@example.com'));
  assert.strictEqual(other.status, 403);
  assert.match(other.body.error, /not allowed to edit/);

  // No session at all.
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, '')).status, 403);
  // A forged cookie.
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, 'session=not.a.real.token')).status, 403);
  // Reading is not writing.
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, cookieFor(OWNER), 'GET')).status, 405);

  // EDITOR_EMAILS widens it, case and spacing ignored.
  process.env.EDITOR_EMAILS = ' Someone.Else@example.com , third@example.com ';
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, cookieFor('someone.else@example.com'))).status, 200);
  assert.strictEqual((await invoke({ app: 'Adobe', month: '2026-02', amount: 40 }, cookieFor(OWNER))).status, 403,
    'an explicit list replaces the default rather than adding to it');
});

// --- what it will accept ----------------------------------------------------

test('an amount is bounded, and a typo with an extra digit is refused', async () => {
  const bad = async (amount) => (await invoke({ app: 'Adobe', month: '2026-02', amount }, cookieFor(OWNER))).body.error;
  assert.match(await bad('abc'), /not a number/);
  assert.match(await bad(-5), /cannot be negative/);
  assert.match(await bad(999999999), /extra digit/);
  assert.strictEqual(writes.length, 0, 'nothing reached the sheet');

  // Money as people type it.
  const ok = await invoke({ app: 'Adobe', month: '2026-02', amount: '$1,234.567' }, cookieFor(OWNER));
  assert.strictEqual(ok.status, 200);
  assert.strictEqual(ok.body.after, 1234.57, 'rounded to cents');
  assert.strictEqual(ok.body.before, 37.16);
  assert.strictEqual(writes[0].address, 'I3', 'the Feb-26 column of the Adobe row');
});

test('a cell can be cleared, which also drops its lock', async () => {
  await invoke({ app: 'Cumul(Luzmo)', month: '2026-03', amount: 14638 }, cookieFor(OWNER));
  assert.deepStrictEqual(rulesFile.locks.map(l => [l.app, l.month, l.value]), [['Cumul(Luzmo)', '2026-03', 14638]]);
  const cleared = await invoke({ app: 'Cumul(Luzmo)', month: '2026-03', amount: '' }, cookieFor(OWNER));
  assert.strictEqual(cleared.body.after, '');
  assert.strictEqual(cleared.body.shown, 'cleared');
  assert.deepStrictEqual(rulesFile.locks, [], 'an empty cell is not a figure to defend');
});

test('a saved amount is locked, so the next sync does not revise it back', async () => {
  const r = await invoke({ app: 'Cumul(Luzmo)', month: '2026-03', amount: 14638 }, cookieFor(OWNER));
  assert.deepStrictEqual(r.body.lock, { locked: true });
  const lock = rulesFile.locks[0];
  assert.strictEqual(lock.value, 14638);
  assert.match(lock.note, /Set in the dashboard by owner@example\.com/);
  assert.ok(rulesFile.vendors.length, 'the vendor rules survive the write');

  // Editing the same cell again replaces the lock rather than stacking one.
  await invoke({ app: 'Cumul(Luzmo)', month: '2026-03', amount: 99 }, cookieFor(OWNER));
  assert.strictEqual(rulesFile.locks.length, 1);
  assert.strictEqual(rulesFile.locks[0].value, 99);
});

test('an unreadable rules file is never replaced by an empty one', async () => {
  const keep = rulesFile;
  rulesFile = null;                               // a failed read, not an empty archive
  const r = await invoke({ app: 'Adobe', month: '2026-02', amount: 50 }, cookieFor(OWNER));
  assert.strictEqual(r.status, 200, 'the edit itself still succeeds');
  assert.strictEqual(r.body.lock.locked, false);
  assert.match(r.body.lock.why, /could not be read/);
  assert.strictEqual(rulesFile, null, 'and nothing was written over it');
  rulesFile = keep;
});

test('the editable columns are named, and the application name is not one of them', async () => {
  const r = await invoke({ app: 'Adobe', field: 'renewalDate', value: '19-09-2027' }, cookieFor(OWNER));
  assert.strictEqual(r.status, 200);
  assert.strictEqual(r.body.before, '1st of every month');
  assert.strictEqual(writes[0].address, 'D3', 'the Renewal data column');

  // Renaming a row from here would orphan its invoices, rules and locks.
  const renamed = await invoke({ app: 'Adobe', field: 'name', value: 'Adobe CC' }, cookieFor(OWNER));
  assert.strictEqual(renamed.status, 400);
  assert.match(renamed.body.error, /not an editable column/);
  assert.match(renamed.body.error, /dept, poc, renewalDate/);

  const long = await invoke({ app: 'Adobe', field: 'poc', value: 'x'.repeat(61) }, cookieFor(OWNER));
  assert.strictEqual(long.status, 400);
  assert.match(long.body.error, /limited to 60 characters/);
});

test('one thing at a time, and only rows and months the sheet actually has', async () => {
  const both = await invoke({ app: 'Adobe', month: '2026-02', field: 'poc', value: 'x' }, cookieFor(OWNER));
  assert.match(both.body.error, /not both/);
  assert.match((await invoke({ app: 'Adobe' }, cookieFor(OWNER))).body.error, /either a month.*or a field/);
  assert.match((await invoke({ app: '', month: '2026-02', amount: 1 }, cookieFor(OWNER))).body.error, /Which application/);
  assert.match((await invoke({ app: 'Ghost App', month: '2026-02', amount: 1 }, cookieFor(OWNER))).body.error, /not a row in the sheet/);
  assert.match((await invoke({ app: 'Adobe', month: '2027-11', amount: 1 }, cookieFor(OWNER))).body.error, /no column for 2027-11/);
  assert.match((await invoke({ app: 'Adobe', month: 'Feb', amount: 1 }, cookieFor(OWNER))).body.error, /reads as 2026-08/);
  assert.strictEqual(writes.length, 0);
  // The Total row carries =SUM() and is not an app, so it cannot be addressed.
  assert.match((await invoke({ app: 'Total', month: '2026-02', amount: 1 }, cookieFor(OWNER))).body.error, /not a row in the sheet/);
});

test('every write is recorded with who made it and what was there before', async () => {
  await invoke({ app: 'Adobe', month: '2026-02', amount: 41 }, cookieFor(OWNER));
  const entry = logEntries[logEntries.length - 1];
  assert.strictEqual(entry.by, OWNER);
  assert.strictEqual(entry.attribution, 'dashboard-edit');
  assert.deepStrictEqual(entry.cells, [{ app: 'Adobe', month: '2026-02', field: null, address: 'I3', before: 37.16, after: 41 }]);
  assert.strictEqual(sessionsOpened, sessionsClosed, 'the workbook session is always closed');
});

// --- "Cycle" is two columns, not one ----------------------------------------
//
// The sheet has no Cycle column: api/spend-data derives what the dashboard
// shows from Recurring/Onetime and FREQUENCY, where "one" in the first wins.
// Writing only FREQUENCY would leave a one-time row reading One-time whatever
// was typed, and typing "One-time" into FREQUENCY would come back as Monthly.

const RO = 4, FREQ = 5;   // Recurring/Onetime, FREQUENCY — columns E and F
function withRecurringOnetime(text, fn) {
  const wasV = SHEET_VALUES[2][RO], wasT = SHEET_TEXT[2][RO];
  SHEET_VALUES[2][RO] = text; SHEET_TEXT[2][RO] = text;
  return Promise.resolve(fn()).finally(() => { SHEET_VALUES[2][RO] = wasV; SHEET_TEXT[2][RO] = wasT; });
}

test('a recurring frequency is written to FREQUENCY', async () => {
  const r = await invoke({ app: 'Adobe', field: 'cycle', value: 'Half Yearly' }, cookieFor(OWNER));
  assert.strictEqual(r.status, 200);
  assert.deepStrictEqual(writes, [{ sheetName: 'Spendings', address: 'F3', value: 'Half Yearly' }]);
  assert.strictEqual(r.body.before, 'Monthly');
});

test('marking a row one-time writes Recurring/Onetime, which is what decides it', async () => {
  const r = await invoke({ app: 'Adobe', field: 'cycle', value: 'One-time' }, cookieFor(OWNER));
  assert.strictEqual(r.status, 200);
  // Not F3: FREQUENCY has no reading for "One-time", so the row would have come
  // back Monthly and the edit would have looked like it did nothing.
  assert.deepStrictEqual(writes, [{ sheetName: 'Spendings', address: 'E3', value: 'Onetime' }]);
});

test('a one-time row given a frequency stops being one-time', async () => {
  await withRecurringOnetime('Onetime', async () => {
    const r = await invoke({ app: 'Adobe', field: 'cycle', value: 'Annual' }, cookieFor(OWNER));
    assert.strictEqual(r.status, 200);
    assert.deepStrictEqual(writes, [
      { sheetName: 'Spendings', address: 'F3', value: 'Annual' },
      { sheetName: 'Spendings', address: 'E3', value: 'Recurring' },
    ]);
    // Both cells are in the audit log, each with what it held before.
    const cells = logEntries[logEntries.length - 1].cells;
    assert.deepStrictEqual(cells.map(c => [c.address, c.before, c.after]),
      [['F3', 'Monthly', 'Annual'], ['E3', 'Onetime', 'Recurring']]);
    assert.strictEqual(sessionsOpened, 1);
    assert.strictEqual(sessionsClosed, 1, 'both writes share one workbook session');
  });
});

test('a row already one-time and marked one-time again is left alone', async () => {
  await withRecurringOnetime('Onetime', async () => {
    const r = await invoke({ app: 'Adobe', field: 'cycle', value: 'Onetime' }, cookieFor(OWNER));
    assert.strictEqual(r.status, 200);
    assert.deepStrictEqual(writes, [{ sheetName: 'Spendings', address: 'E3', value: 'Onetime' }]);
  });
});

test('the raw columns stay addressable on their own', async () => {
  await invoke({ app: 'Adobe', field: 'recurring', value: 'Recurring' }, cookieFor(OWNER));
  assert.deepStrictEqual(writes, [{ sheetName: 'Spendings', address: 'E3', value: 'Recurring' }]);
});
