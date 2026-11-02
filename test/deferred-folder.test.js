// Run with: node --test
//
// Two quite different things stop a month's cell being written, and for a while
// the dashboard called them the same thing.
//
// DBT Cloud's Nov-26 folder holds two invoices — 106.25 and 212.50, both of
// which read perfectly — and their two payment receipts. A run with the parse
// budget nearly spent read the first and stopped. The report then said:
//
//   Not written — an invoice in the month could not be used, so the total is
//   not known:
//   • DBT Cloud 2026-11: readable invoices come to 106.25
//
// which is not what happened. Nothing in that folder is unreadable; the run
// simply ended early, and the next one would have finished it. The owner went
// looking for a broken PDF that does not exist.
//
// So the two are reported apart: a folder the run did not finish says so and
// asks for nothing, and "an invoice could not be read" is kept for when one
// genuinely could not be.

const test = require('node:test');
const assert = require('node:assert');

const graph = require('../lib/graph');
const invoiceAmount = require('../lib/invoice-amount');

// The live Nov-26 folder.
const FILES = [
  'Invoice-C70CD8FE-0007.pdf',
  'Invoice-NV7SZL90-0009.pdf',
  'Receipt-2389-9778.pdf',
  'Receipt-2791-8444.pdf',
];

let downloads = [];
graph.graphFetch = async url => {
  const u = decodeURIComponent(String(url));
  if (u.includes(':/children')) {
    return { ok: true, status: 200, json: async () => ({ value: FILES.map(name => ({ name, file: {} })) }) };
  }
  downloads.push(u.slice(u.indexOf('root:/') + 6).replace(/:\/content$/, ''));
  return { ok: true, status: 200, arrayBuffer: async () => new ArrayBuffer(8), text: async () => '' };
};

// Every one of them reads. That is the whole point of the fixture.
const BY_NAME = {
  'Invoice-C70CD8FE-0007.pdf': { amount: 106.25, ref: 'C70CD8FE-0007' },
  'Invoice-NV7SZL90-0009.pdf': { amount: 212.50, ref: 'NV7SZL90-0009' },
  'Receipt-2389-9778.pdf': { amount: 106.25, ref: 'C70CD8FE-0007' },
  'Receipt-2791-8444.pdf': { amount: 212.50, ref: 'NV7SZL90-0009' },
};
// lib/mail-sync destructures these when it loads, so the stub is installed once
// here and reads mutable state — reassigning it later would not take. Files are
// listed and fetched in order, so the queue stays in step with them.
let queue = [];
let current = null;
invoiceAmount.readInvoiceTotal = async () => {
  current = queue.shift();
  return { text: '', error: null, reader: 'pdfjs', total: { amount: current.amount, currency: 'USD', usable: true, note: '' } };
};
invoiceAmount.extractInvoiceRef = () => current.ref;

const { sumFolderInvoices, planAmountCells, parseBudget, formatRunReport } = require('../lib/mail-sync');

function armReads() { queue = FILES.map(n => BY_NAME[n]); }

const FOLDER = 'Invoices/DBT Cloud/Nov-26';

test('a run that stops mid-folder reports what it did not reach, not a bad file', async () => {
  armReads();
  downloads = [];
  const budget = parseBudget(Date.now() + 60_000);
  budget.maxParse = 1;                       // room for the first invoice only

  const res = await sumFolderInvoices('tok', 'drive-1', FOLDER, new Map(), budget);

  assert.strictEqual(downloads.length, 1, 'the budget stopped it after one');
  assert.strictEqual(res.total, 106.25);
  assert.strictEqual(res.deferred, 3, 'three files were never looked at');
  assert.strictEqual(res.unread.length, 0, 'and none of them is unreadable');
  assert.strictEqual(res.unusable.length, 0);
});

test('a folder the run gets all the way through defers nothing', async () => {
  armReads();
  downloads = [];
  const res = await sumFolderInvoices('tok', 'drive-1', FOLDER, new Map(), parseBudget(Date.now() + 60_000));

  assert.strictEqual(downloads.length, 4);
  assert.strictEqual(res.deferred, 0);
  // The receipts are the same two charges, so the month is 106.25 + 212.50.
  assert.strictEqual(res.total, 318.75);
  assert.strictEqual(res.counted, res.pdfCount, 'nothing missing, so the cell may be written');
});

// --- what the owner is told -------------------------------------------------

const GRID = {
  apps: [{ name: 'DBT Cloud', rowIdx: 0 }],
  monthCols: { '2026-11': 1 },
};
const USED = { values: [['DBT Cloud', null]], start: { col: 0, row: 0 } };
const cellValue = (values, r, c) => values[r][c];

test('an unfinished folder holds the cell without blaming a PDF', () => {
  const { write, updated, skippedFilled } = planAmountCells(
    { 'DBT Cloud': { '2026-11': 106.25 } }, GRID, USED, cellValue, {},
    { 'DBT Cloud||2026-11': 'deferred' }
  );
  assert.strictEqual(write.length + updated.length, 0, 'a part-read folder is still not a total');
  assert.strictEqual(skippedFilled[0].reason, 'folder-total-deferred');
});

test('an unreadable invoice still reports as one', () => {
  const { skippedFilled } = planAmountCells(
    { 'DBT Cloud': { '2026-11': 106.25 } }, GRID, USED, cellValue, {},
    { 'DBT Cloud||2026-11': true }
  );
  assert.strictEqual(skippedFilled[0].reason, 'folder-total-incomplete');
});

test('the report tells the two apart', () => {
  const report = formatRunReport({
    // A report is only sent when the run did something; the held cells are
    // what it has to say about what it did not do.
    filed: 1, perApp: { Zapier: 1 },
    amountsSkippedFilled: [
      { app: 'DBT Cloud', month: '2026-11', invoiceTotal: 106.25, reason: 'folder-total-deferred' },
      { app: 'Apollo', month: '2026-08', invoiceTotal: 85, reason: 'folder-total-incomplete' },
    ],
  }, []);

  assert.match(report, /ran out of time/, 'the unfinished folder says the run was short, not the folder');
  assert.match(report, /DBT Cloud 2026-11: \$?106\.25 read so far/);
  assert.match(report, /could not be used/, 'and a genuinely unreadable invoice still says so');
  assert.match(report, /Apollo 2026-08/);

  const deferredLine = report.split('\n').find(l => l.includes('DBT Cloud 2026-11'));
  assert.doesNotMatch(deferredLine, /could not be/, 'nothing in that folder could not be read');
});
