// Run with: node --test
//
// Exercises the preview/apply handlers against a stubbed Graph layer. The point
// of these tests is cell targeting: a proposal for "AWS" in "Jun-26" must write
// to exactly the cell the sheet layout puts it in, and nowhere else.

const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('xlsx');

// The handlers read these before touching Graph.
process.env.TARGET_USER_UPN = 'test@example.com';
process.env.TARGET_FILE_PATH = 'Test/Spend.xlsx';
process.env.SESSION_SECRET = 'test-secret';

// Stub the Graph layer before anything that captures its exports is loaded.
const graph = require('../lib/graph');
graph.getGraphToken = async () => 'test-token';
graph.resolveDriveId = async () => 'drive-1';

let aliasStore = {};
let logStore = null;
graph.readJsonFile = async (_t, _d, path) => {
  if (path.endsWith('_amount-map.json')) return { aliases: aliasStore };
  if (path.endsWith('_amount-log.json')) return logStore;
  return null;
};
graph.writeJsonFile = async (_t, _d, path, obj) => {
  if (path.endsWith('_amount-map.json')) aliasStore = obj.aliases;
  if (path.endsWith('_amount-log.json')) logStore = obj;
  return {};
};

// The live sheet's shape: a title row, a header row, app rows, then a Total row
// carrying =SUM() formulas that must never be touched.
const SHEET_VALUES = [
  ['Spendings', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Renewal data', 'Recurring/Onetime', 'FREQUENCY', 'Payment Method', 'Jan-26', 'Feb-26', 'Mar-26', 'Apr-26', 'May-26', 'Jun-26'],
  ['Adobe', 'Marketing', 'Bhavana', '', 'Recurring', 'Monthly', 'US Debit Card', 37.16, 37.16, 37.16, 37.16, 37.16, ''],
  ['AWS', 'Engineering', 'Ajay', '', 'Recurring', 'Monthly', 'US Debit Card', 4552.64, 2673.27, 2116.53, 2397.77, 2525.33, ''],
  ['Bubble Starter', 'Product', 'Ganesh', '', 'Recurring', 'Monthly', 'US Debit Card', 320, 1073.71, 765.11, 732.67, 570.91, ''],
  ['Chargebee', 'Product', 'Abhishek', '', 'Recurring', 'Monthly', 'US Debit Card', 126.44, 126.44, 126.44, 126.44, 158.91, 169.43],
  ['Sentry.io', 'Product', 'Venky', '', 'Recurring', 'Monthly', 'US Debit Card', 30.81, 47.35, 121.87, 80.99, 79.96, ''],
  ['Google cloud', 'DE', 'Ajay', '', 'Recurring', 'Monthly', 'US Debit Card', 20000, 29095.51, 27338.52, 13216.44, 27404.19, ''],
  ['Total', '', '', '', '', '', '', 25067.05, 33053.44, 30505.63, 16591.47, 30776.46, 169.43],
];

const excel = require('../lib/excel');
excel.resolveItemId = async () => 'item-1';
excel.listWorksheets = async () => [{ id: '1', name: 'Spendings' }, { id: '2', name: 'Invoices tracker' }];
excel.readUsedRange = async () => ({
  values: SHEET_VALUES,
  formulas: [],
  text: [],
  start: { col: 0, row: 0 }, // usedRange starts at A1
  address: 'Spendings!A1:M9',
});

let written = [];
let sessionsOpened = 0, sessionsClosed = 0;
excel.createSession = async () => { sessionsOpened++; return 'session-1'; };
excel.closeSession = async () => { sessionsClosed++; };
const realWriteCells = excel.writeCells;
excel.writeCells = async (_t, _d, _i, sheetName, cells) => {
  written.push(...cells.map(c => ({ ...c, sheetName })));
  return cells.map(c => ({ ...c, ok: true }));
};

const previewHandler = require('../api/amounts/preview');
const applyHandler = require('../api/amounts/apply');

function statementBuffer(rows) {
  const aoa = [['BOA ', 'Date', 'Month', 'Description', 'Amount', 'Comments']];
  for (const r of rows) aoa.push(['BOA-0218', r[0], '26-Jun', r[1], r[2], r[3]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function invoke(handler, body, cookie) {
  return new Promise((resolve) => {
    const req = { method: 'POST', body, query: {}, headers: { cookie: cookie || '' } };
    const res = {
      statusCode: 200,
      _json: null,
      setHeader() {},
      status(code) { this.statusCode = code; return this; },
      json(payload) { this._json = payload; resolve({ status: this.statusCode, body: payload }); },
    };
    handler(req, res);
  });
}

const STATEMENT = [
  ['6/2/2026', 'ADOBE *ADOBE 06/01 PURCHASE 408-536-6000 CA DEBIT CARD *4154', -37.16, 'Adobe'],
  ['6/12/2026', 'AMAZON WEB SERVI DES:INTERNET CO ID:9049016352 CCD', -2525.33, 'Amazon Web services'],
  ['6/1/2026', 'BUBBLE STARTER PLAN 06/01 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -751.96, 'Bubble Starter'],
  ['6/1/2026', 'Chargebee 06/01 PURCHASE 187-7303113 CA DEBIT CARD *4154', -175.00, 'Chargebee'],
  ['6/30/2026', 'SENTRY 06/29 PURCHASE SENTRY.IO CA DEBIT CARD *4154', -9000.00, 'Sentry'],
  ['6/8/2026', 'PADDLE.NET* EXTENSIONS 06/06 PURCHASE New York NY DEBIT CARD *4154', -10.61, 'Paddle Net'],
];

test.beforeEach(() => { written = []; aliasStore = {}; logStore = null; });

test('classifies an empty month cell as ready to apply', async () => {
  const { status, body } = await invoke(previewHandler, {
    contentBase64: statementBuffer(STATEMENT).toString('base64'), filename: 'june.xlsx',
  });
  assert.strictEqual(status, 200);
  const adobe = body.proposals.find(p => p.app === 'Adobe' && p.month === '2026-06');
  assert.strictEqual(adobe.status, 'auto');
  assert.strictEqual(adobe.amount, 37.16);
  assert.strictEqual(adobe.current, null);
});

test('holds back a cell that already has a different figure', async () => {
  const { body } = await invoke(previewHandler, {
    contentBase64: statementBuffer(STATEMENT).toString('base64'), filename: 'june.xlsx',
  });
  // Chargebee's Jun-26 cell already holds 169.43; the statement says 175.
  const cb = body.proposals.find(p => p.app === 'Chargebee');
  assert.strictEqual(cb.status, 'review');
  assert.strictEqual(cb.reason, 'overwrite');
  assert.strictEqual(cb.current, 169.43);
});

test('holds back a figure far out of line with the trailing average', async () => {
  const { body } = await invoke(previewHandler, {
    contentBase64: statementBuffer(STATEMENT).toString('base64'), filename: 'june.xlsx',
  });
  // Sentry averages ~95/month; 9000 should not be written unattended.
  const sentry = body.proposals.find(p => p.app === 'Sentry.io');
  assert.strictEqual(sentry.status, 'review');
  assert.strictEqual(sentry.reason, 'swing');
});

test('reports an unrecognised vendor instead of guessing a row', async () => {
  const { body } = await invoke(previewHandler, {
    contentBase64: statementBuffer(STATEMENT).toString('base64'), filename: 'june.xlsx',
  });
  assert.ok(body.unmapped.some(u => /paddle/i.test(u.label)));
  assert.ok(!body.proposals.some(p => /paddle/i.test(p.app || '')));
});

test('a pending vendor mapping turns an unmapped label into a proposal', async () => {
  const { body } = await invoke(previewHandler, {
    contentBase64: statementBuffer(STATEMENT).toString('base64'),
    filename: 'june.xlsx',
    aliasOverrides: { paddlenet: 'Adobe' },
  });
  assert.ok(!body.unmapped.some(u => /paddle/i.test(u.label)), 'should no longer be unmapped');
  const adobe = body.proposals.find(p => p.app === 'Adobe');
  // 37.16 (Adobe) + 10.61 (the newly mapped Paddle line)
  assert.ok(Math.abs(adobe.amount - 47.77) < 0.005, `got ${adobe.amount}`);
});

test('writes each approved amount to the exact cell for that app and month', async () => {
  const { status, body } = await invoke(applyHandler, {
    cells: [
      { app: 'Adobe', month: '2026-06', amount: 37.16 },
      { app: 'AWS', month: '2026-06', amount: 2525.33 },
      { app: 'Google cloud', month: '2026-01', amount: 20001 },
    ],
    statementName: 'june.xlsx',
  });
  assert.strictEqual(status, 200);
  assert.strictEqual(body.written, 3);

  const at = app => written.find(w => w.app === app).address;
  // Header row 2 puts Jun-26 in column M; Adobe is row 3, AWS row 4.
  assert.strictEqual(at('Adobe'), 'M3');
  assert.strictEqual(at('AWS'), 'M4');
  // Google cloud is row 8; Jan-26 is column H.
  assert.strictEqual(at('Google cloud'), 'H8');
  // Nothing may land on the Total row (9) or outside the month columns.
  for (const w of written) {
    const row = Number(w.address.replace(/[A-Z]/g, ''));
    assert.notStrictEqual(row, 9, 'must never write to the Total row');
    assert.ok(row >= 3 && row <= 8, `row ${row} is outside the app rows`);
  }
  assert.strictEqual(sessionsOpened, sessionsClosed, 'every workbook session must be closed');
});

test('records what each cell held before the write', async () => {
  await invoke(applyHandler, {
    cells: [{ app: 'Chargebee', month: '2026-06', amount: 175 }],
    statementName: 'june.xlsx',
  });
  assert.strictEqual(written[0].before, 169.43);
  assert.strictEqual(written[0].value, 175);
  const entry = logStore.entries[logStore.entries.length - 1];
  assert.strictEqual(entry.cells[0].before, 169.43);
  assert.strictEqual(entry.cells[0].after, 175);
  assert.strictEqual(entry.statement, 'june.xlsx');
});

test('refuses a cell whose app or month is not in the sheet', async () => {
  const { status, body } = await invoke(applyHandler, {
    cells: [{ app: 'Not An App', month: '2026-06', amount: 10 }],
  });
  assert.strictEqual(status, 400);
  assert.strictEqual(written.length, 0, 'nothing may be written');
  assert.ok(body.rejected.some(r => /no row/.test(r.error)));
});

test('writes only the approved cells, leaving every other row alone', async () => {
  await invoke(applyHandler, { cells: [{ app: 'Adobe', month: '2026-06', amount: 37.16 }] });
  assert.strictEqual(written.length, 1);
  assert.strictEqual(written[0].address, 'M3');
});

test('only saves a vendor mapping that points at a real app row', async () => {
  await invoke(applyHandler, {
    cells: [{ app: 'Adobe', month: '2026-06', amount: 37.16 }],
    aliasUpdates: { paddlenet: 'Adobe', bogus: 'Nonexistent App' },
  });
  assert.strictEqual(aliasStore.paddlenet, 'Adobe');
  assert.strictEqual(aliasStore.bogus, undefined);
});

test.after(() => { excel.writeCells = realWriteCells; });
