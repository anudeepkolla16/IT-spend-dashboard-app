// Run with: node --test
//
// The breakdown gains a column every month and never drops one, on a table
// that already scrolls sideways: eleven columns in November 2026, twenty-three
// a year later. So the table shows the last twelve months, with a toggle for
// the whole history.
//
// The window is the TABLE'S, and nothing else's. Every measure on the page —
// run-rate, annual cost, stage, forecast, the charts, the app modal — reads
// ALL_MONTHS, and a column you chose not to look at must never move a figure.
//
// The one thing a window can get wrong is arithmetic. If the Total column goes
// on meaning the whole history while the cells beside it are twelve months,
// then every row on screen fails to add up, and a reader who checks one is
// right and the table is wrong.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

// 14 months, so the window actually bites.
const MONTHS = ['2025-10','2025-11','2025-12','2026-01','2026-02','2026-03','2026-04',
                '2026-05','2026-06','2026-07','2026-08','2026-09','2026-10','2026-11'];

function win(months) {
  const a = html.indexOf('const MONTH_WINDOW = 12;');
  const b = html.indexOf('function renderWindowBtn(){', a);
  assert.ok(a > -1 && b > a, 'the window state must sit where the test expects');
  return new Function(`
    const ALL_MONTHS = ${JSON.stringify(months)};
    ${html.slice(a, b)}
    ; return { shownMonths, windowedTotal, MONTH_WINDOW, setWindow: n => { monthWindow = n; } };`)();
}

test('the table shows the last twelve months of fourteen', () => {
  const W = win(MONTHS);
  const shown = W.shownMonths();
  assert.strictEqual(shown.length, 12);
  assert.strictEqual(shown[0], '2025-12', 'the two oldest drop off');
  assert.strictEqual(shown[11], '2026-11', 'and the current month is always the last');
});

test('the toggle shows every month on record', () => {
  const W = win(MONTHS);
  W.setWindow(0);
  assert.deepStrictEqual(W.shownMonths(), MONTHS);
});

test('a sheet shorter than the window is shown whole', () => {
  // November 2026 has eleven months on record, so the window is a no-op and
  // nothing the owner sees today changes.
  const W = win(MONTHS.slice(-11));
  assert.strictEqual(W.shownMonths().length, 11);
});

test('the Total column adds up the cells beside it, not the ones hidden', () => {
  const W = win(MONTHS);
  const g = { total: 1400 };               // 14 months at 100
  for (const m of MONTHS) g[m] = 100;
  assert.strictEqual(W.windowedTotal(g, W.shownMonths()), 1200, 'twelve columns, twelve hundred');
  W.setWindow(0);
  assert.strictEqual(W.windowedTotal(g, W.shownMonths()), 1400, 'unwindowed it is the row total');
});

test('unwindowed, the Total still catches a charge whose month could not be read', () => {
  // g.total counts every charge; the month cells only count the ones with a
  // month. Summing the columns instead would quietly drop the odd one.
  const W = win(MONTHS);
  W.setWindow(0);
  const g = { total: 140, '2026-11': 100 };
  assert.strictEqual(W.windowedTotal(g, W.shownMonths()), 140);
});

test('the footer total follows the window too', () => {
  const a = html.indexOf('function breakdownTotals(rows, months){');
  const b = html.indexOf('\nfunction draw()', a);
  const breakdownTotals = new Function(`
    const ALL_MONTHS = ${JSON.stringify(MONTHS)};
    ${html.slice(a, b)}
    ; return breakdownTotals;`)();

  const g = { name: 'AWS', total: 1400 };
  for (const m of MONTHS) g[m] = 100;
  assert.strictEqual(breakdownTotals([g], MONTHS.slice(-12)).grand, 1200,
    'the figure under the Total column must equal that column');
  assert.strictEqual(breakdownTotals([g], MONTHS).grand, 1400);
});

// --- what the window must NOT touch -----------------------------------------

test('every measure still reads the full range', () => {
  // A toggle that moved the run-rate or the annual cost would be a reporting
  // bug wearing a feature's clothes.
  for (const line of [
    /const annual = ALL_MONTHS\.filter/,
    /const months = ALL_MONTHS\.filter\(m => g\[m\] != null && m <= todayKey\)/,
    /for \(const m of ALL_MONTHS\) if \(m >= from && m <= todayKey && g\[m\] != null\) spend \+= g\[m\]/,
    /data:\{labels:ALL_MONTHS\.map\(monthLabel\)/,            // the app modal's chart
  ]) assert.match(html, line, `${line} must keep reading ALL_MONTHS`);
});

test('the spike flags are worked out over the whole history', () => {
  // A jump against the month before it, even when that month is off screen.
  const a = html.indexOf('g._spikes = {};');
  const b = html.indexOf('return pivoted;', a);
  assert.match(html.slice(a, b), /ALL_MONTHS\.forEach/, 'not the window');
});

// --- the wiring -------------------------------------------------------------

test('the header is rebuilt when the window changes, not just on a sync', () => {
  // Columns in the header and cells in the body are built in different
  // functions; a toggle that redrew only one would misalign the table.
  assert.match(html, /renderHead\(\);\s*\n\s*draw\(\);\s*\n\}/, 'setupTable must build the head then draw');
  const a = html.indexOf("$('#monthWindowBtn').onclick");
  const b = html.indexOf('};', a);
  const handler = html.slice(a, b);
  assert.match(handler, /renderHead\(\);/);
  assert.match(handler, /draw\(\);/);
});

test('sorting by a month that leaves the table does not leave the order unexplained', () => {
  const a = html.indexOf("$('#monthWindowBtn').onclick");
  const handler = html.slice(a, html.indexOf('};', a));
  assert.match(handler, /ALL_MONTHS\.includes\(sortKey\) && !shownMonths\(\)\.includes\(sortKey\)/);
  assert.match(handler, /sortKey = 'total'/);
});

test('the choice is remembered', () => {
  assert.match(html, /localStorage\.setItem\(MONTH_WINDOW_KEY/);
  assert.match(html, /localStorage\.getItem\(MONTH_WINDOW_KEY\)/);
  // and never throws where storage is blocked
  assert.match(html, /try \{ const v = localStorage\.getItem\(MONTH_WINDOW_KEY\); if \(v !== null\) monthWindow = Number\(v\) \|\| 0; \} catch\(e\)\{\}/);
});

test('the export carries the columns that are on screen', () => {
  const a = html.indexOf("$('#exportBtn').onclick");
  const ex = html.slice(a, html.indexOf('XLSX.writeFile', a));
  assert.match(ex, /const months = shownMonths\(\);/);
  assert.match(ex, /\.\.\.months\.map\(monthLabel\)/);
  assert.match(ex, /windowedTotal\(g, months\)/);
  assert.doesNotMatch(ex, /ALL_MONTHS\.map\(m =>/, 'a hidden column must not reappear in the file');
});

test('the button only shows on the tab whose columns it picks', () => {
  assert.match(html, /\$\('#monthWindowBtn'\)\.classList\.toggle\('hidden', tab !== 'monthly'\);/);
});
