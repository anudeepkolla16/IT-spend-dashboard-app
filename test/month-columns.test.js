// Run with: node --test
//
// On 2 November 2026 the Monthly breakdown ran Jan 2026 to Oct 2026. November
// was missing — not one app had a November figure yet, because the month's
// invoices were still being read — and the columns were built from the months
// the rows happened to mention. So there was nowhere to see the current month,
// no cell to type a figure into, and no "month so far" footer.
//
// The same gap would swallow any quiet month inside the range, which is worse
// than it looks: two columns sitting side by side would not be consecutive
// months, and nobody reading across a row would know.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
function lift(names, from, to, extra) {
  const a = html.indexOf(from), b = html.indexOf(to, a);
  assert.ok(a > -1 && b > a, `${from} must sit where the test expects`);
  return new Function(`${extra || ''}\n${html.slice(a, b)}\n; return {${names.join(',')}};`)();
}

const { monthRange } = lift(['monthRange'], '// Every month from the first on record', 'let ALL_MONTHS = [];');

test('the current month has a column even when nothing is charged to it yet', () => {
  const seen = ['2026-08', '2026-09', '2026-10'];
  assert.deepStrictEqual(monthRange(seen, '2026-11'), ['2026-08', '2026-09', '2026-10', '2026-11']);
});

test('a quiet month in the middle still gets its column', () => {
  // Otherwise Feb and Apr sit next to each other and read as consecutive.
  assert.deepStrictEqual(
    monthRange(['2026-01', '2026-02', '2026-04'], '2026-04'),
    ['2026-01', '2026-02', '2026-03', '2026-04']
  );
});

test('the range crosses a year end', () => {
  assert.deepStrictEqual(
    monthRange(['2026-11'], '2027-02'),
    ['2026-11', '2026-12', '2027-01', '2027-02']
  );
});

test('budgeted months past the current one are kept', () => {
  // The sheet pre-fills figures for months still ahead; they were always shown
  // and must stay shown.
  assert.deepStrictEqual(
    monthRange(['2026-10', '2027-01'], '2026-11'),
    ['2026-10', '2026-11', '2026-12', '2027-01']
  );
});

test('an empty sheet gets no columns rather than one stray month', () => {
  assert.deepStrictEqual(monthRange([], '2026-11'), []);
});

test('a bad date does not render ten thousand columns', () => {
  // One row with a nonsense month would otherwise span to the year 9999.
  const seen = ['2026-01', '9999-01'];
  assert.deepStrictEqual(monthRange(seen, '2026-11'), seen, 'falls back to the months themselves');
});

test('the breakdown and its footer both read the same month list', () => {
  // ALL_MONTHS is what the header, every row, the footer, the export and the
  // modal chart iterate. Building it from anything but monthRange is the bug.
  assert.match(html, /ALL_MONTHS = monthRange\(\[\.\.\.new Set\(all\.filter\(r=>r\.month\)\.map\(r=>r\.month\)\)\]\.sort\(\), realMonthKey\(\)\);/,
    'setupTable must build the columns as a continuous range through the current month');
});

test('a month with no figures shows a dash, not a nought', () => {
  // An empty November must not read as "we were charged $0" — the footer
  // counts the apps that have a figure and prints an em dash when none does.
  const { breakdownTotals } = lift(
    ['breakdownTotals'],
    '// Column sums for the monthly breakdown',
    'function draw(){',
    `const ALL_MONTHS = ['2026-10','2026-11'];`
  );
  const foot = breakdownTotals([{ '2026-10': 425, total: 425 }]);
  assert.strictEqual(foot.countByMonth['2026-11'], 0, 'nothing is charged to November yet');
  assert.strictEqual(foot.byMonth['2026-11'], 0);
  assert.strictEqual(foot.grand, 425, 'and the grand total is unchanged by the empty column');
});
