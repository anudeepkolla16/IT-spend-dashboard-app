// Run with: node --test
//
// The backfill re-files invoices that were archived before the billing-period
// rule existed. It moves files in someone's SharePoint and rewrites cells in the
// spend sheet, so what it declines to touch matters as much as what it moves.

const test = require('node:test');
const assert = require('node:assert');
const { planMove } = require('../lib/invoices/period-backfill');

// The archive is one folder, located at run time (lib/graph.js). These paths
// are built from whatever root the caller resolved, so the fixture names one.
const BASE = 'Desktop/Anudeep files/Invoices';

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
  folderMonth: '2026-08',   // it sits in a month folder
  nameMonth: null,
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

test('every path is built from the archive root it was given', () => {
  // The archive has been renamed once already, and hardcoding its old name is
  // what broke the checklist. A move must follow the root the caller resolved.
  const other = 'Some/Other/Archive';
  const v = planMove(luzmoFile({ path: `${other}/Cumul/Aug-26/x.pdf` }), luzmoFolder(), other);
  assert.strictEqual(v.move.toFolderPath, `${other}/Cumul/Sep-26`);
});

test('a vendor folder with no app row still moves', () => {
  const v = planMove(luzmoFile(), luzmoFolder({ app: null }), BASE);
  assert.ok(v.move);
  assert.strictEqual(v.move.app, null);
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

test('a file not in a month folder, with nothing in its name, is left alone', () => {
  const file = luzmoFile({ relPath: '', currentMonth: null, folderMonth: null, nameMonth: null });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

// Much of the archive keeps invoices flat in the app folder with the month in
// the name — "jan 26.pdf", "Aug 2026.pdf" — which is how the checklist dates
// them. There is no month folder to move such a file out of, and renaming
// somebody's files to impose one is a different job from this one.

test('a loose invoice whose name agrees with its period is left alone', () => {
  const file = luzmoFile({
    name: 'Aug 2026.pdf', relPath: '', currentMonth: '2026-08',
    folderMonth: null, nameMonth: '2026-08',
    periodStart: '2026-08-01', periodEnd: '2026-08-31',
  });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('a loose invoice whose period disagrees with its name is reported, not moved', () => {
  // The checklist dates this one August from the name; it bills September. That
  // is worth knowing, and it is the owner's call whether to rename or re-file.
  const file = luzmoFile({
    name: 'Aug 2026.pdf', relPath: '', currentMonth: '2026-08',
    folderMonth: null, nameMonth: '2026-08',
  });
  const v = planMove(file, luzmoFolder(), BASE);
  assert.ok(v.skip, 'expected it to be reported');
  assert.ok(!v.move, 'a loose file is never moved');
  assert.match(v.skip, /name reads as 2026-08, but it bills 2026-09/);
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

test('a file the run had no time to read is not judged', () => {
  // Pending is not "unreadable": the next run reads it. Reporting it either way
  // would put a guess in front of somebody about to move their files.
  const file = luzmoFile({ pending: true, read: false, periodStart: null, periodEnd: null });
  assert.strictEqual(planMove(file, luzmoFolder(), BASE), null);
});

test('reading is bounded by the run clock, not a fixed count', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'invoices', 'period-backfill.js'), 'utf8');
  // 460 invoices at 25 a run is eighteen clicks. Reads go through a pool and
  // stop on the deadline, with time in hand to total and cache what was read.
  assert.match(src, /READ_POOL/, 'PDFs must be read in parallel');
  assert.match(src, /deadline - READ_RESERVE_MS/, 'reading must stop with time left to finish the run');

  const ui = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.match(ui, /scan\.unread > 0/, 'the dashboard must keep going until the archive is read');
  assert.match(ui, /next\.unread >= scan\.unread/, 'and must stop when a round makes no progress');
});

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
  // A reply without a `periods` payload is not this feature answering, and
  // rendering it printed "Read undefined invoices across undefined folders".
  assert.match(ui, /if \(!json\.periods\) throw new Error\(/,
    'a reply from an older deployment must be named, not rendered as undefined');
});
