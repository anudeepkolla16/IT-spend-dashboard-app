// Run with: node --test
//
// The backfill re-files invoices that were archived before the billing-period
// rule existed. It moves files in someone's SharePoint and rewrites cells in the
// spend sheet, so what it declines to touch matters as much as what it moves.

const test = require('node:test');
const assert = require('node:assert');
const { planMove } = require('../lib/invoices/period-backfill');

const BASE = 'Desktop/Anudeep files/Procurment bills';

// The Cumul(Luzmo) folder as it stands: an Aug-26 folder, no Sep-26 folder yet.
const luzmoFolder = (extra) => ({
  app: 'Cumul(Luzmo)',
  vendorFolder: 'Cumul',
  month: '2026-08',
  monthFolderNames: ['Jul-26', 'Aug-26'],
  ...(extra || {}),
});

const luzmoFile = (extra) => ({
  name: '20260826_20260258.pdf',
  path: `${BASE}/Cumul/Aug-26/20260826_20260258.pdf`,
  relPath: 'Aug-26',
  currentMonth: '2026-08',
  read: true,
  periodStart: '2026-08-26',
  periodEnd: '2026-09-26',
  amount: 557.28,
  currency: 'USD',
  usable: true,
  ...(extra || {}),
});

test('the invoice that started this moves to the month it bills for', () => {
  const v = planMove(luzmoFile(), luzmoFolder(), BASE);
  assert.ok(v && v.move, 'expected a move');
  assert.strictEqual(v.move.fromMonth, '2026-08');
  assert.strictEqual(v.move.toMonth, '2026-09');
  assert.strictEqual(v.move.fromPath, `${BASE}/Cumul/Aug-26/20260826_20260258.pdf`);
  assert.strictEqual(v.move.amount, 557.28);
});

test('the dashboard copy moves with it, or the checklist keeps showing the old month', () => {
  const v = planMove(luzmoFile(), luzmoFolder(), BASE);
  assert.strictEqual(v.move.mirrorFromPath, 'Invoices/Cumul(Luzmo)/Aug-26/20260826_20260258.pdf');
  assert.strictEqual(v.move.mirrorToFolderPath, 'Invoices/Cumul(Luzmo)/Sep-26');
});

test('an unmapped vendor moves in the archive only', () => {
  const v = planMove(luzmoFile(), luzmoFolder({ app: null }), BASE);
  assert.ok(v.move);
  assert.strictEqual(v.move.mirrorFromPath, null);
  assert.strictEqual(v.move.mirrorToFolderPath, null);
});

test('the destination reuses what this vendor already calls that month', () => {
  // Cursor's folders are spelled "July", "August" — the move must not drop a
  // second "Sep-26" folder in beside "September".
  const v = planMove(
    luzmoFile(),
    luzmoFolder({ vendorFolder: 'Cursor', monthFolderNames: ['August', 'September'] }),
    BASE
  );
  assert.strictEqual(v.move.toFolderPath, `${BASE}/Cursor/September`);
});

test('a month folder that does not exist yet is named in the house style', () => {
  const v = planMove(luzmoFile(), luzmoFolder({ monthFolderNames: ['Aug-26'] }), BASE);
  assert.strictEqual(v.move.toFolderPath, `${BASE}/Cumul/Sep-26`);
});

// --- What it refuses to touch ---------------------------------------------

test('an invoice already in the right month is left alone', () => {
  const file = luzmoFile({ relPath: 'Sep-26', currentMonth: '2026-09', path: `${BASE}/Cumul/Sep-26/x.pdf` });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('an invoice that states no period is left alone', () => {
  // Most invoices state none. Their arrival month is the best anyone knows, and
  // moving them on a guess would be worse than leaving them.
  const file = luzmoFile({ periodStart: null, periodEnd: null });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('a PDF nobody could read is reported, never moved', () => {
  const v = planMove(luzmoFile({ read: false, note: 'scanned image' }), luzmoFolder(), BASE);
  assert.strictEqual(v.skip, 'scanned image');
  assert.ok(!v.move);
});

test('a file not in a month folder has nothing to be moved from', () => {
  const file = luzmoFile({ relPath: '', currentMonth: null });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('a file nested deeper than {vendor}/{month} is left for a human', () => {
  const v = planMove(luzmoFile({ relPath: '2026/Aug' }), luzmoFolder(), BASE);
  assert.match(v.skip, /not a plain month folder/);
  assert.ok(!v.move);
});

test('a period far from where the file sits is treated as a misread', () => {
  // A contract term read out of the wrong line would otherwise fling an invoice
  // a year across the archive. Two months is the limit, as in the mail sync.
  const file = luzmoFile({ periodStart: '2027-06-01', periodEnd: '2027-06-30' });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('an annual period keeps the invoice in the month it starts', () => {
  const file = luzmoFile({ periodStart: '2026-08-01', periodEnd: '2027-07-31' });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('an invoice whose total could not be used still moves, without an amount', () => {
  // An INR invoice cannot fill a USD cell, but it is still filed in the wrong
  // month. The file moves; the month it lands in is then reported as one the
  // sheet cannot be totalled for.
  const v = planMove(luzmoFile({ usable: false, currency: 'INR', amount: 150591.6 }), luzmoFolder(), BASE);
  assert.ok(v.move);
  assert.strictEqual(v.move.amount, null);
});

// --- Wiring ---------------------------------------------------------------

test('the scan writes nothing and the apply only touches what was approved', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'invoices', 'period-backfill.js'), 'utf8');

  // The scan's only write is the cache of what it read — no moves, no cells.
  const scan = src.slice(src.indexOf('async function scanPeriods'), src.indexOf('// The month totals the moves imply'));
  assert.ok(!/moveItem\(/.test(scan), 'the scan must never move a file');
  assert.ok(!/writeCells\(/.test(scan), 'the scan must never write to the sheet');

  // The apply writes a month only if the caller ticked it.
  assert.match(src, /approved\.has\(`\$\{app\}\|\|\$\{month\}`\)/,
    'a cell must be written only when it was approved');
  // And the figure written is re-derived from the folder, not taken on trust.
  assert.match(src, /sumFolderInvoices\(token, driveId, folder, cache, budget\)/,
    'the value written must be totalled from the folder after the moves');

  const ui = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.match(ui, /post\('periods'\)/, 'the dashboard must be able to run the scan');
  assert.match(ui, /post\('periods-apply', \{ moves, cells \}\)/,
    'and must send back the very moves and cells the scan proposed');
  assert.match(ui, /Move the files and update the sheet\?/, 'nothing moves without a confirmation');
});
