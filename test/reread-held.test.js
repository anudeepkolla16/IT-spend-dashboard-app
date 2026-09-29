// Run with: node --test
//
// rereadHeld used to re-read only the invoices held as UNREADABLE, so a held
// invoice kept whatever figure the reader of the day gave it. Docusign's
// September invoice sat in the queue reading 143.44 against a real 2,828.58
// long after the reader was fixed — it was held for want of a filing rule, not
// for being unreadable — and answering it would have written the stale figure
// into the sheet.
//
// Now every held invoice is re-read, but only once per PARSE_VERSION, so a
// reader fix sweeps the queue once and a steady-state run downloads nothing.

const test = require('node:test');
const assert = require('node:assert');

// Stub the Graph and PDF layers before lib/mail-sync captures their exports.
const graph = require('../lib/graph');
const invoiceAmount = require('../lib/invoice-amount');

let downloads = [];
graph.graphFetch = async url => {
  downloads.push(decodeURIComponent(String(url)));
  return { ok: true, status: 200, arrayBuffer: async () => new ArrayBuffer(8), text: async () => '' };
};

// The Docusign summary, read correctly.
const TEXT = 'Docusign, Inc. Invoice\nInvoice #: 111100734908\nSubTotal 2,685.14\nTax Total* 143.44\nTotal 2,828.58\nCurrency USD';
let readResult = { text: TEXT, error: null, total: { amount: 2828.58, currency: 'USD', usable: true, note: '' } };
invoiceAmount.readInvoiceTotal = async () => readResult;

const mailSync = require('../lib/mail-sync');
const rulesLib = require('../lib/invoices/rules');
// `classify` is looked up on the module at call time, so it can be stubbed here.
// No rule names this sender — which is exactly why the invoice is held.
rulesLib.classify = () => ({
  app: null, vendor: null, options: [], currency: null, period: null,
  reason: 'no-rule', question: 'Which app is this invoice for?',
});

const DOCUSIGN = () => ({
  id: 'P61',
  heldPath: 'Invoices/_Pending/P61-111100734908.pdf',
  file: '111100734908_20260929070134_1.pdf',
  from: 'invoicing@erp.docusign.com',
  subject: 'Invoice 111100734908 is attached',
  reason: 'no-rule',
  question: 'Which app is this invoice for?',
  amount: 143.44,          // what the old reader gave it
  currency: 'USD',
  receivedMonth: '2026-09',
  parsed: 2,               // read under the patterns that produced 143.44
});

const run = held => mailSync.rereadHeld('tok', 'drive-1', {
  root: { path: 'Invoices' }, rules: { vendors: [] }, appNames: ['Docusign'],
  held, deadline: Date.now() + 60 * 1000,
});

test('a held invoice keeps no stale figure once the reader has moved on', async () => {
  downloads = [];
  const held = { items: [DOCUSIGN()] };
  const out = await run(held);

  assert.strictEqual(downloads.length, 1, 'the held PDF is fetched again');
  assert.strictEqual(held.items[0].amount, 2828.58, 'the tax figure is replaced by the real total');
  assert.strictEqual(out.changed, true, 'so the queue is written back to disk');
  assert.strictEqual(out.answers.length, 0, 'still no rule names the sender, so nothing is filed');

  const report = out.items.find(i => i.id === 'P61');
  assert.ok(report, 'a corrected figure is worth reporting');
  assert.match(report.note, /2,828\.58/);
  assert.match(report.note, /was 143\.44/);
});

test('a second run re-reads nothing', async () => {
  const held = { items: [DOCUSIGN()] };
  await run(held);
  downloads = [];
  const out = await run(held);
  assert.strictEqual(downloads.length, 0, 'the item is stamped with the version it was read under');
  assert.strictEqual(out.items.length, 0, 'and a quiet run says nothing');
});

test('an unchanged figure is re-read but not reported', async () => {
  // The sweep must not list the whole queue every time the version moves.
  downloads = [];
  const held = { items: [{ ...DOCUSIGN(), amount: 2828.58 }] };
  const out = await run(held);
  assert.strictEqual(downloads.length, 1, 'it is still read, to find out');
  assert.strictEqual(out.items.length, 0, 'but there is nothing to say about it');
});

test('a sheet-figure question is never fetched', async () => {
  // kind: 'cell' items park no PDF; asking Graph for their heldPath would be a
  // request for `undefined`.
  downloads = [];
  const held = { items: [{ id: 'P7', kind: 'cell', app: 'AWS', month: '2026-08', reason: 'amount-differs', question: 'keep or use?' }] };
  const out = await run(held);
  assert.strictEqual(downloads.length, 0);
  assert.strictEqual(out.changed, false);
});

test('an invoice held as unreadable is still retried on every run', async () => {
  // Its fix is a new reader, not a new total, so the version stamp must not
  // park it — it is tried again until it reads.
  const previous = readResult;
  readResult = { text: '', error: 'no text could be read (it may be a scan)', total: null };
  try {
    const held = { items: [{ ...DOCUSIGN(), reason: 'unreadable', parsed: undefined }] };
    await run(held);
    downloads = [];
    const out = await run(held);
    assert.strictEqual(downloads.length, 1, 'tried again rather than stamped and forgotten');
    assert.strictEqual(out.items[0].result, 'still unreadable');
  } finally {
    readResult = previous;
  }
});
