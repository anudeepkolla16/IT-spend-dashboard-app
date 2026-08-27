// Run with: node --test
//
// Covers the invoice checklist's archive read: turning the hand-made month
// subfolder names in the archive into real months, and joining parsed amounts
// back onto the mirrored copies of the files.

const test = require('node:test');
const assert = require('node:assert');

const {
  parseMonthFolder, monthFromRelPath, monthFromFileName, detectDayFirst,
  amountsByFileName, monthIndexFromWord, buildInventory,
} = require('../lib/invoices/inventory');

const AUG_26 = '2026-08-20T10:00:00Z';

test('month folders: every shape the archive actually uses', () => {
  // Real folder names from Procurment bills/: inconsistent by vendor.
  assert.equal(parseMonthFolder('Aug-26', AUG_26), '2026-08');
  assert.equal(parseMonthFolder('July', AUG_26), '2026-07');
  assert.equal(parseMonthFolder('Aug', AUG_26), '2026-08');
  assert.equal(parseMonthFolder('2026-08', AUG_26), '2026-08');
  assert.equal(parseMonthFolder('August 2026', AUG_26), '2026-08');
  assert.equal(parseMonthFolder('Aug 26', AUG_26), '2026-08');
  assert.equal(parseMonthFolder('Jan-2026', AUG_26), '2026-01');
  assert.equal(parseMonthFolder('01-2026', AUG_26), '2026-01');
  assert.equal(parseMonthFolder('2026.3', AUG_26), '2026-03');
});

test('month folders: anything that is not a month stays unclaimed', () => {
  for (const name of ['Invoices', 'Old', 'Misc', '', 'Q1', '2026', '13-2026', '2026-13']) {
    assert.equal(parseMonthFolder(name, AUG_26), null, `"${name}" should not read as a month`);
  }
});

test('a bare month name takes its year from when the file landed', () => {
  // A December invoice filed in January belongs to the year before, not ahead.
  assert.equal(parseMonthFolder('December', '2026-01-09T00:00:00Z'), '2025-12');
  // September filed in August is next month's -- close enough to be a genuine
  // early filing rather than last year's.
  assert.equal(parseMonthFolder('September', AUG_26), '2026-09');
  // A month already past in the same year stays in that year. The mirror stamps
  // every copied file with the copy date, so a January invoice first mirrored in
  // August must not be pushed back to the year before.
  assert.equal(parseMonthFolder('March', AUG_26), '2026-03');
  assert.equal(parseMonthFolder('January', AUG_26), '2026-01');
  // Only a month well ahead of the file date rolls back.
  assert.equal(parseMonthFolder('November', AUG_26), '2025-11');
});

test('month abbreviations shorter than three letters are refused', () => {
  // "Ma" is March or May and there is no way to tell; guessing would file
  // invoices under the wrong month silently.
  assert.equal(monthIndexFromWord('ma'), -1);
  assert.equal(monthIndexFromWord('mar'), 2);
  assert.equal(monthIndexFromWord('may'), 4);
  assert.equal(monthIndexFromWord('jun'), 5);
  assert.equal(monthIndexFromWord('jul'), 6);
});

test('the month is found however deep the subfolder chain goes', () => {
  assert.equal(monthFromRelPath('Aug-26', AUG_26), '2026-08');
  assert.equal(monthFromRelPath('2026/Aug', AUG_26), '2026-08');
  assert.equal(monthFromRelPath('Paid/July/scanned', AUG_26), '2026-07');
  assert.equal(monthFromRelPath('', AUG_26), null);
  assert.equal(monthFromRelPath('Misc', AUG_26), null);
});

test('amounts join to the mirrored copies by file name', () => {
  // The index records the procurement-folder path; the dashboard archive is a
  // mirror at a different path, so a path join would find nothing.
  const map = amountsByFileName({
    amounts: [
      { path: 'Desktop/Anudeep files/Procurment bills/Bubble/Aug-26/Invoice-A1.pdf', amount: 64, currency: 'USD', usable: true },
    ],
    entries: [
      { file: 'Tata-June.pdf', amount: 150591.6, currency: 'INR', amountUsable: false },
    ],
  });
  assert.deepEqual(map.get('invoice-a1.pdf'), { amount: 64, currency: 'USD', usable: true });
  assert.deepEqual(map.get('tata-june.pdf'), { amount: 150591.6, currency: 'INR', usable: false });
});

test('a file name carrying two different amounts is dropped, not guessed', () => {
  // Vendors reuse generic names ("invoice.pdf"). Attributing one invoice's
  // total to another is worse than showing no total at all.
  const map = amountsByFileName({
    amounts: [
      { path: 'a/Aug/invoice.pdf', amount: 10, currency: 'USD', usable: true },
      { path: 'b/Aug/invoice.pdf', amount: 20, currency: 'USD', usable: true },
      { path: 'c/Aug/invoice.pdf', amount: 10, currency: 'USD', usable: true },
    ],
  });
  assert.equal(map.has('invoice.pdf'), false);
});

test('the same amount recorded twice for one name is still usable', () => {
  const map = amountsByFileName({
    amounts: [
      { path: 'a/Aug/dup.pdf', amount: 42, currency: 'USD', usable: true },
      { path: 'a/Aug/dup.pdf', amount: 42, currency: 'USD', usable: true },
    ],
  });
  assert.deepEqual(map.get('dup.pdf'), { amount: 42, currency: 'USD', usable: true });
});

test('an empty or missing index yields an empty map rather than throwing', () => {
  assert.equal(amountsByFileName(null).size, 0);
  assert.equal(amountsByFileName({}).size, 0);
  assert.equal(amountsByFileName({ amounts: 'nonsense' }).size, 0);
});

test('the month is read from the file name when there is no month subfolder', () => {
  // Verbatim file names from the archive. Most vendors are filed flat in the app
  // folder with the month in the name, so reading only the subfolder left 220 of
  // 472 invoices undated -- shown on the checklist as months charged with no
  // invoice, while the PDF sat right there.
  const cases = {
    // Adobe
    'jan 26.pdf': '2026-01', 'feb 26.pdf': '2026-02', 'March 26.pdf': '2026-03',
    'apr 26.pdf': '2026-04', 'May 26.pdf': '2026-05', 'June 26.pdf': '2026-06',
    'july 26.pdf': '2026-07', 'Aug 2026.pdf': '2026-08',
    // AWS
    'Jan-26.pdf': '2026-01', 'feb-26.pdf': '2026-02', 'mar-26.pdf': '2026-03',
    'Apr-26.pdf': '2026-04', 'May-26.pdf': '2026-05', 'June-26.pdf': '2026-06',
    // Chargebee
    'April 26.pdf': '2026-04', 'may 26.pdf': '2026-05',
  };
  for (const [name, want] of Object.entries(cases)) {
    assert.equal(monthFromFileName(name, AUG_26, null), want, name);
  }
});

test('a month never matches inside a longer word', () => {
  // "Marchant", "Augustine", "Decision" are not months, and a vendor name that
  // happens to contain one must not date an invoice.
  for (const name of ['Marchant Ltd 1234.pdf', 'Augustine Partners.pdf', 'Decision-support.pdf', 'Janus Systems.pdf']) {
    assert.equal(monthFromFileName(name, AUG_26, null), null, name);
  }
});

test('a file name with no date at all stays undated', () => {
  for (const name of ['Invoice-1234.pdf', 'receipt.pdf', 'Signed Delivery Report.pdf', '']) {
    assert.equal(monthFromFileName(name, AUG_26, null), null, name);
  }
});

test('the upload timestamp this app prefixes is not read as the month', () => {
  // A hand-uploaded January invoice must not be dated by the day it was uploaded.
  assert.equal(monthFromFileName('2026-08-27T09-15-00-000Z_jan 26.pdf', AUG_26, null), '2026-01');
});

test('an ISO date in the name is read year-first', () => {
  assert.equal(monthFromFileName('invoice_2026-03-14.pdf', AUG_26, null), '2026-03');
  assert.equal(monthFromFileName('statement 2026-11.pdf', AUG_26, null), '2026-11');
  assert.equal(monthFromFileName('bad 2026-13-01.pdf', AUG_26, null), null);
});

test("a folder's date order is settled by its own unambiguous names", () => {
  // Real Anthropic file names. "27-01-2026" can only be day-first, and that
  // settles "02-03-2026", which on its own could be either.
  const names = ['Claude Purchase 02-03-2026.pdf', 'Claude Purchase 27-01-2026.pdf',
    'Claude Purchase 03-12-2025.pdf', 'Claude Purchase 09-01-2026.pdf'];
  assert.equal(detectDayFirst(names), true);
  assert.equal(monthFromFileName(names[0], AUG_26, true), '2026-03');
  assert.equal(monthFromFileName(names[1], AUG_26, true), '2026-01');
  assert.equal(monthFromFileName(names[2], AUG_26, true), '2025-12');
});

test('a month-first folder is detected too', () => {
  const names = ['inv 01-13-2026.pdf', 'inv 02-05-2026.pdf'];
  assert.equal(detectDayFirst(names), false);
  assert.equal(monthFromFileName('inv 02-05-2026.pdf', AUG_26, false), '2026-02');
});

test('an ambiguous numeric date with nothing to settle it is refused', () => {
  // Filing an invoice under the wrong month is worse than leaving it undated:
  // a wrong tick claims a gap is covered when it is not.
  assert.equal(detectDayFirst(['inv 03-12-2025.pdf', 'inv 05-06-2026.pdf']), null);
  assert.equal(monthFromFileName('inv 03-12-2025.pdf', AUG_26, null), null);
});

test('a bare month name in the file name takes the year from the file date', () => {
  assert.equal(monthFromFileName('adobe july.pdf', AUG_26, null), '2026-07');
  assert.equal(monthFromFileName('december statement.pdf', '2026-01-09T00:00:00Z', null), '2025-12');
});

/* ---------- buildInventory against a stubbed Graph ---------- */

// Mirrors the Graph shapes the real calls return, including the fields
// listFilesRecursive selects. Getting these wrong in a fixture is how the
// header-row bug shipped, so the stub answers the same URLs the code builds.
const ARCHIVE_PATH = 'Desktop/Anudeep files/Invoices';

function stubGraph(tree) {
  const original = global.fetch;
  const calls = [];
  global.fetch = async (url) => {
    calls.push(String(url));
    const u = String(url);
    if (u.includes('_invoice-index.json')) {
      return { ok: true, status: 200, json: async () => tree.index || {} };
    }
    // The archive-root probe: only the real archive path resolves, so the
    // earlier candidates 404 exactly as they do on the live drive.
    const probe = u.match(/\/root:\/(.+?)\?/);
    if (probe) {
      const path = decodeURIComponent(probe[1]).split('/').map(decodeURIComponent).join('/');
      if (path === ARCHIVE_PATH) return { ok: true, status: 200, json: async () => ({ id: 'root-inv', folder: {} }) };
      return { ok: false, status: 404, text: async () => 'itemNotFound' };
    }
    const m = u.match(/items\/([^/]+)\/children/);
    if (m) {
      const id = decodeURIComponent(m[1]);
      // A folder the tree does not declare fails its listing, the way a deleted
      // or permission-denied folder does in Graph.
      if (!(id in tree.children)) return { ok: false, status: 403, text: async () => 'accessDenied' };
      return { ok: true, status: 200, json: async () => ({ value: tree.children[id] }) };
    }
    return { ok: false, status: 404, text: async () => 'not found' };
  };
  return { calls, restore: () => { global.fetch = original; } };
}

// resolveArchiveRoot caches per driveId, so every test needs its own.
let driveN = 0;
const nextDrive = () => `drive-inv-${++driveN}`;

const folder = (id, name) => ({ id, name, folder: {} });
const file = (id, name, extra) => ({
  id, name, file: {}, size: 1024, createdDateTime: AUG_26,
  webUrl: `https://example/${name}`, ...(extra || {}),
});

test('buildInventory flattens the archive and dates each file by its folder', async () => {
  const stub = stubGraph({
    index: { amounts: [{ path: 'src/Bubble/Aug-26/bubble-aug.pdf', amount: 64, currency: 'USD', usable: true }] },
    children: {
      'root-inv': [folder('f-bubble', 'Bubble Starter'), folder('f-adobe', 'Adobe'), folder('f-un', '_Unmatched')],
      'f-bubble': [folder('f-bubble-aug', 'Aug-26'), file('x-loose', 'loose.pdf'), file('x-flat', 'jan 26.pdf')],
      'f-bubble-aug': [file('x1', 'bubble-aug.pdf')],
      'f-adobe': [folder('f-adobe-jul', 'July')],
      'f-adobe-jul': [file('x2', 'adobe-jul.pdf')],
    },
  });
  try {
    const inv = await buildInventory('tok', nextDrive(), { pool: 2 });

    assert.deepEqual(inv.apps, ['Adobe', 'Bubble Starter']);
    assert.equal(inv.fileCount, 4);
    assert.deepEqual(inv.months, ['2026-01', '2026-07', '2026-08']);

    // A flat file dated only by its name is picked up alongside the foldered ones.
    const flat = inv.files.find(f => f.name === 'jan 26.pdf');
    assert.equal(flat.month, '2026-01');
    assert.equal(flat.monthSource, 'filename');
    assert.equal(inv.datedFromNameCount, 1);

    const bubbleAug = inv.files.find(f => f.name === 'bubble-aug.pdf');
    assert.equal(bubbleAug.app, 'Bubble Starter');
    assert.equal(bubbleAug.month, '2026-08');
    assert.equal(bubbleAug.monthSource, 'folder');
    assert.equal(bubbleAug.amount, 64);
    assert.equal(bubbleAug.amountUsable, true);

    const adobe = inv.files.find(f => f.name === 'adobe-jul.pdf');
    assert.equal(adobe.month, '2026-07');
    assert.equal(adobe.amount, null);

    // A file sitting straight under the app folder has no month to read.
    const loose = inv.files.find(f => f.name === 'loose.pdf');
    assert.equal(loose.month, null);
    assert.equal(inv.undatedCount, 1);
  } finally { stub.restore(); }
});

test('the sync\'s own bookkeeping folders are not reported as applications', async () => {
  const stub = stubGraph({
    children: {
      'root-inv': [folder('f-a', 'Adobe'), folder('f-u', '_Unmatched'), folder('f-t', '_temp')],
      'f-a': [file('x', 'a.pdf')],
      'f-u': [file('y', 'mystery.pdf')],
      'f-t': [],
    },
  });
  try {
    const inv = await buildInventory('tok', nextDrive());
    assert.deepEqual(inv.apps, ['Adobe']);
    assert.equal(inv.fileCount, 1);
  } finally { stub.restore(); }
});

test('a missing Invoices folder reports empty rather than failing', async () => {
  const original = global.fetch;
  global.fetch = async () => ({ ok: false, status: 404, text: async () => 'not found' });
  try {
    const inv = await buildInventory('tok', nextDrive());
    assert.equal(inv.empty, true);
    assert.deepEqual(inv.files, []);
  } finally { global.fetch = original; }
});

test('one unreadable folder is reported but does not lose the rest', async () => {
  const stub = stubGraph({
    children: {
      'root-inv': [folder('f-a', 'Adobe'), folder('f-bad', 'Broken')],
      'f-a': [file('x', 'a.pdf')],
      // 'f-bad' is absent from the tree, so its listing 404s.
    },
  });
  try {
    const inv = await buildInventory('tok', nextDrive(), { pool: 1 });
    assert.equal(inv.fileCount, 1);
    assert.equal(inv.errors.length, 1);
    assert.match(inv.errors[0], /^Broken: /);
  } finally { stub.restore(); }
});

test('a scan that runs past its deadline says so instead of reporting false gaps', async () => {
  const stub = stubGraph({
    children: {
      'root-inv': [folder('f-a', 'Adobe'), folder('f-b', 'Bubble')],
      'f-a': [file('x', 'a.pdf')],
      'f-b': [file('y', 'b.pdf')],
    },
  });
  try {
    const inv = await buildInventory('tok', nextDrive(), { pool: 1, deadline: Date.now() - 1 });
    assert.equal(inv.truncated, true);
  } finally { stub.restore(); }
});
