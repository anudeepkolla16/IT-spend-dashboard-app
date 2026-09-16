// Run with: node --test
//
// The 16 September review found three ways the dashboard's numbers could not be
// trusted, and all three were real:
//   A  "Annual cost" $0 and "Total (all-time)" $120 for the same app, because
//      the sheet keeps actual charges and budgeted months in the same columns.
//   B  An Applications search matching nothing turned every recorded charge into
//      an apparent missing charge in Invoices: coverage read "0 of 0 · 100%".
//   C  A part-billed September was reported as a 62% saving against a whole
//      August, in green, on the 16th.
// These pin the fixes against the page source, which is where they live.

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

const TODAY = '2026-09';
const M = lift(
  ['spendIn', 'actualSpend', 'closedSpend', 'mtdSpend', 'plannedSpend', 'billedRows'],
  '/* ---------------- what "spend" means',
  'function latestRowsPerApp(',
  `const realMonthKey = () => '${TODAY}';`
);

const row = (name, month, usd, extra) => ({ name, month, usd, kind: 'Apps', dept: 'IT', cat: 'Monthly', poc: '', renewalDate: '', ...(extra || {}) });
// One app charged through August, one whose only figure is a budgeted future month.
const ROWS = [
  row('AWS', '2026-07', 2260.67), row('AWS', '2026-08', 2374.66), row('AWS', '2026-09', 900),
  row('CANVA', '2026-11', 120, { cat: 'Annual' }),
  { ...row('Laptops', '2026-06', 3493.17), kind: 'Laptops' },
  row('Nothing', null, 0),
];

// --- A. actual, planned and month-to-date are separate measures --------------

test('the four measures split actual charges from budgeted months', () => {
  assert.strictEqual(M.actualSpend(ROWS), 2260.67 + 2374.66 + 900 + 3493.17, 'through this month, this month included');
  assert.strictEqual(M.closedSpend(ROWS), 2260.67 + 2374.66 + 3493.17, 'months that have finished');
  assert.strictEqual(M.mtdSpend(ROWS), 900);
  assert.strictEqual(M.plannedSpend(ROWS), 120, 'a month still ahead is a budget, never spend');
  assert.strictEqual(M.actualSpend(ROWS) + M.plannedSpend(ROWS), ROWS.reduce((s, r) => s + (r.month ? r.usd : 0), 0),
    'actual and planned account for every dated figure exactly once');
  assert.strictEqual(M.closedSpend(ROWS) + M.mtdSpend(ROWS), M.actualSpend(ROWS));
});

test('billedRows drops budgeted months and rows with no month at all', () => {
  assert.deepStrictEqual(M.billedRows(ROWS).map(r => `${r.name} ${r.month}`),
    ['AWS 2026-07', 'AWS 2026-08', 'AWS 2026-09', 'Laptops 2026-06']);
  assert.strictEqual(M.billedRows(ROWS).length, ROWS.length - 2);
  assert.strictEqual(M.actualSpend([]), 0);
  assert.strictEqual(M.actualSpend(null), 0);
});

test('the app card and the detail panel report the same pair of measures', () => {
  // Canva: nothing charged, $120 budgeted ahead. The table's 12-month column and
  // the modal must not disagree about that.
  assert.match(html, /Spend · last 12 mo/, 'the column names its measure');
  assert.match(html, /Actual charges in the twelve months ending with the current one/);
  assert.match(html, /<span class="cap">Actual spend · to date<\/span>/, 'the modal reports actual');
  assert.match(html, /<span class="cap">Budgeted ahead<\/span>/, 'and planned, apart');
  assert.ok(!/cap">Total \(all-time\)/.test(html), 'the label that mixed them is no longer rendered');
  assert.match(html, /if \(m <= modalToday\) modalActual \+= v; else modalPlanned \+= v;/);
  // The KPI strip and the donut say which basis they are on.
  assert.match(html, /Actual spend · to date<\/div><div class="val">\$\{fmtUSD\(toDate\)\}/);
  assert.match(html, /budgeted ahead, not yet billed/);
  assert.match(html, /Last 12<br>months by<br>department/);
  // Charts of "spend" count charges only.
  assert.match(html, /const g=\{\}; billedRows\(apps\)\.forEach/);
  assert.match(html, /groupSum\(billedRows\(all\),'cat'\)/);
});

// --- B. reconciliation is a property of the data, not of a filter ------------

const P = lift(['buildPivot'], 'function buildPivot(all, filt){', '\nfunction draw()', `
  const ALL_MONTHS = ['2026-07','2026-08','2026-09','2026-11'];
  const window = { SHEET_ROWS: [
    { name:'AWS', kind:'Apps', dept:'Engineering', cycle:'Monthly', cur:'USD', poc:'Ajay', renewalDate:'' },
    { name:'CANVA', kind:'Apps', dept:'Marketing', cycle:'Annual', cur:'USD', poc:'Bhavana', renewalDate:'' },
    { name:'Laptops', kind:'Laptops', dept:'IT', cycle:'Monthly', cur:'USD', poc:'', renewalDate:'' },
  ] };`);

test('buildPivot with no filter is the whole sheet; a filter only narrows the copy', () => {
  const all = P.buildPivot(ROWS, {});
  assert.deepStrictEqual(all.map(g => g.name).sort(), ['AWS', 'CANVA', 'Laptops', 'Nothing']);
  assert.strictEqual(all.find(g => g.name === 'AWS')['2026-08'], 2374.66);
  assert.strictEqual(all.find(g => g.name === 'CANVA')['2026-11'], 120);

  assert.deepStrictEqual(P.buildPivot(ROWS, { d: 'Marketing' }).map(g => g.name), ['CANVA']);
  assert.deepStrictEqual(P.buildPivot(ROWS, { t: 'Laptops' }).map(g => g.name), ['Laptops']);
  assert.deepStrictEqual(P.buildPivot(ROWS, { q: 'aws' }).map(g => g.name), ['AWS']);
  // The case that broke reconciliation: a search matching nothing.
  assert.deepStrictEqual(P.buildPivot(ROWS, { q: 'zzzz' }), []);
  // …and the unfiltered pivot is untouched by it.
  assert.strictEqual(P.buildPivot(ROWS, {}).length, 4);
});

test('the checklist and the drill-down read the unfiltered pivot', () => {
  assert.match(html, /function invBuildRows\(\)\{\s*\/\/[^\n]*\n\s*const pivoted = window\.PIVOT_ALL \|\| window\.PIVOTED \|\| \[\];/);
  assert.match(html, /const g = \(window\.PIVOT_ALL\|\|window\.PIVOTED\|\|\[\]\)\.find\(x=>x\.name===name\)/,
    'a renewal or invoice cell still opens its app while an unrelated search is active');
  assert.match(html, /window\.PIVOT_ALL = buildPivot\(all, \{\}\);/);
  assert.match(html, /let pivoted = \(t \|\| d \|\| q\) \? buildPivot\(all, \{ t, d, q \}\) : window\.PIVOT_ALL\.slice\(\);/);
});

test('no charged month reports no coverage, never 100%', () => {
  assert.match(html, /const pct = charged \? Math\.round\(covered\/charged\*100\) : null;/);
  assert.match(html, /'No charged months to reconcile'/);
  assert.ok(!/: 100;/.test(html), 'the empty denominator no longer reads as full coverage');
});

// --- C. an incomplete period is never reported as a saving -------------------

test('the current month is labelled month-to-date and compared against nothing', () => {
  assert.match(html, /Spend · \$\{monthLabel\(thisKey\)\} <span class="mtd">month to date<\/span>/);
  assert.match(html, /const thisMonth = mtdSpend\(all, thisKey\);/);
  // The trend it does report is the last two months that finished.
  assert.match(html, /const lastClosed = totals\[prevKey\] \|\| 0, prevClosed = totals\[prev2Key\] \|\| 0;/);
  assert.match(html, /closed at \$\{fmtUSD\(lastClosed\)\}/);
  assert.ok(!/mom>=0\?'▲':'▼'/.test(html), 'the month-to-date arrow is gone');
});

test('the spend chart is solid to the last closed month and dashed after it', () => {
  assert.match(html, /const closedEnd = months\.reduce\(\(n, m\) => \(m < todayKey \? n \+ 1 : n\), 0\);/);
  assert.match(html, /segment: \{ borderDash: \(ctx\) => \(ctx\.p0DataIndex >= closedEnd - 1 \? \[5, 4\] : undefined\) \}/);
  assert.match(html, /'budgeted, not yet billed'/);
  assert.match(html, /the last month that finished; dashed after it/);
  // The Forecast tab's own actual series is indexed over its own months.
  assert.match(html, /line\('Apps spend \(actual\)'[^\n]*, true\)/);
});
