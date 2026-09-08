// Run with: node --test
//
// The Forecast tab used to draw a straight line through every month's total,
// and the part-billed current month plus the budgeted months after it dragged
// that line towards zero. It is now built from what the sheet knows: recurring
// apps at their average over the last three complete months, renewals in the
// month they fall at the last charge on record, laptops and one-offs left out.
// The function is pulled out of the page source and run as-is.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

function pageFunction() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const slice = (from, to) => {
    const a = html.indexOf(from), b = html.indexOf(to, a);
    assert.ok(a > 0 && b > a, `${from} must sit where the test expects`);
    return html.slice(a, b);
  };
  const src = [
    slice('function cycleMonths(', '/* Renewal Date text'),
    slice('const MONTH_IDX =', "let renewalsTab = 'yearly';"),
    slice('function monthsBetween(', 'function monthlyRunRate('),
    'const realMonthKey = () => { throw new Error("today must be passed in"); };',
    slice('/* ---------------- forecast:', '/* ---------------- end forecast'),
  ].join('\n');
  return new Function(src + '; return buildForecast;')();
}
function runRate() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const slice = (from, to) => html.slice(html.indexOf(from), html.indexOf(to, html.indexOf(from)));
  const src = [slice('function cycleMonths(', '/* Renewal Date text'), slice('function monthsBetween(', 'function monthlyRunRate('), slice('function addMonths(', 'function buildForecast(')].join('\n');
  return new Function(src + '; return trailingRunRateByApp;')();
}

const row = (name, month, usd, cycle, extra) => ({ kind: 'Apps', name, month, usd, cycle: cycle || 'Monthly', dept: 'Eng', renewalDate: '', ...(extra || {}) });

const ROWS = [
  row('Bubble Starter', '2026-06', 500), row('Bubble Starter', '2026-07', 520), row('Bubble Starter', '2026-08', 540),
  row('Bubble Starter', '2026-09', 100),                      // the current month, part-billed: not a basis
  row('Cursor pro', '2026-08', 1133.53),                      // started in August
  row('Zapier', '2026-05', 31.86),                            // nothing since May: gone
  row('Sprinto', '2025-09', 4800, 'Annual', { renewalDate: '19-09-2026' }),
  row('Adobe', '2026-01', 37.16, 'Annual', { renewalDate: '1/21/2027' }),
  row('Keka', '2026-08', 1200, 'Half-Yearly'),                // no date in the sheet: last charge plus a cycle
  row('Sentry', '2026-07', 300, 'Quarterly'),
  row('Antivirus', '2026-07', 1086.92, 'One-time'),
  row('Wingman', '2026-03', 12189, 'Annual', { renewalDate: '2026-03-01' }), // renews March 2027, outside the window
  { kind: 'Laptops', name: 'Laptops Procurement', month: '2026-06', usd: 5000, cycle: 'One-time', dept: 'IT' },
  row('Bubble Starter', '2026-10', 600),                      // budgeted ahead in the sheet: a plan, not a record
];
const SHEET_ROWS = [
  { name: 'Posthog', kind: 'Apps', cycle: 'Annual', renewalDate: '2026-11-05', dept: 'Product' }, // never charged yet
  { name: 'Laptops Procurement', kind: 'Laptops' },
];

test('only complete months feed the estimate, and the window starts with the current month', () => {
  const f = pageFunction()(ROWS, SHEET_ROWS, '2026-09');
  assert.deepStrictEqual(f.basis, ['2026-06', '2026-07', '2026-08']);
  assert.ok(!f.complete.includes('2026-09') && !f.complete.includes('2026-10'));
  assert.deepStrictEqual(f.months.map(m => m.month), ['2026-09', '2026-10', '2026-11', '2026-12', '2027-01', '2027-02']);
});

test('recurring apps cost their recent average; a new one the months it has run; a lapsed one nothing', () => {
  const f = pageFunction()(ROWS, SHEET_ROWS, '2026-09');
  const by = Object.fromEntries(f.recurring.map(r => [r.name, r]));
  assert.strictEqual(by['Bubble Starter'].perMonth, 520);
  assert.deepStrictEqual(by['Bubble Starter'].basis, ['2026-06', '2026-07', '2026-08']);
  assert.strictEqual(by['Cursor pro'].perMonth, 1133.53, 'not averaged with months before it existed');
  assert.deepStrictEqual(by['Cursor pro'].basis, ['2026-08']);
  assert.ok(!by['Zapier'], 'two complete months without a charge: cancelled');
  assert.ok(!by['Antivirus'] && !by['Laptops Procurement'], 'one-offs and laptops are not projected');
  assert.strictEqual(f.months[0].recurring, 520 + 1133.53);
  assert.strictEqual(f.months[5].recurring, 520 + 1133.53, 'the run-rate is flat across the window');
});

test('renewals land in the month they fall, at the last charge on record', () => {
  const f = pageFunction()(ROWS, SHEET_ROWS, '2026-09');
  const due = f.renewals.map(r => `${r.month} ${r.name} ${r.cost}`).sort();
  assert.deepStrictEqual(due, [
    '2026-09 Sprinto 4800',          // the sheet's date, 19-09-2026
    '2026-10 Sentry 300',            // quarterly from July
    '2026-11 Posthog 0',             // dated in the sheet, never charged: listed with no cost
    '2027-01 Adobe 37.16',           // 1/21/2027
    '2027-01 Sentry 300',            // and again three months on
    '2027-02 Keka 1200',             // half-yearly from August
  ]);
  assert.ok(!due.some(d => /Wingman/.test(d)), 'March 2027 is outside the window');
  assert.strictEqual(f.renewals.find(r => r.name === 'Sprinto').dated, true);
  assert.strictEqual(f.renewals.find(r => r.name === 'Keka').dated, false);
});

test('each month\'s estimate is its recurring run-rate plus its renewals', () => {
  const f = pageFunction()(ROWS, SHEET_ROWS, '2026-09');
  const sep = f.months[0];
  assert.strictEqual(sep.renewalTotal, 4800);
  assert.strictEqual(sep.total, 520 + 1133.53 + 4800);
  assert.deepStrictEqual(sep.renewals.map(r => r.name), ['Sprinto']);
  assert.strictEqual(f.months[3].renewalTotal, 0, 'December: recurring only');
  assert.strictEqual(f.months[4].total, 520 + 1133.53 + 37.16 + 300);
});

test('a renewal date already passed rolls forward by whole cycles', () => {
  const rows = [row('Wingman', '2025-03', 12189, 'Annual', { renewalDate: '2026-03-01' })];
  const f = pageFunction()(rows, [], '2026-09');
  assert.deepStrictEqual(f.renewals, [], 'next is March 2027, past the window');
  const soon = pageFunction()([row('Sprinto', '2025-09', 4800, 'Annual', { renewalDate: '2025-09-19' })], [], '2026-09');
  assert.deepStrictEqual(soon.renewals.map(r => `${r.month} ${r.cost}`), ['2026-09 4800']);
});

test('with no history at all the estimate is empty rather than wrong', () => {
  const f = pageFunction()([], [], '2026-09');
  assert.deepStrictEqual(f.complete, []);
  assert.deepStrictEqual(f.basis, ['2026-06', '2026-07', '2026-08'], 'the three complete months before today, whatever they hold');
  assert.strictEqual(f.months.length, 6);
  assert.ok(f.months.every(m => m.total === 0));
});

test('the page no longer draws a straight-line trend', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.ok(!/linreg/.test(html));
  assert.match(html, /buildForecast\(all, window\.SHEET_ROWS \|\| \[\], realMonthKey\(0\)\)/);
});

// --- The run-rate column, the KPI tile and the forecast share one figure ------
//
// "This monthly run rate is also wrong": Cursor read $94 (labelled Yearly,
// charged monthly, last charge divided by twelve), dbt 388 (its part-billed
// September averaged in), Claude nothing at all (its 0.00 months dropped).

const SHEET = [
  // Cursor pro: labelled Yearly, charged every month, August a real 0.00 (paid from credits)
  row('Cursor pro', '2026-05', 616.57, 'Annual'), row('Cursor pro', '2026-06', 815.04, 'Annual'), row('Cursor pro', '2026-07', 1133.53, 'Annual'),
  // DBT: steady, then a part-billed September
  row('DBT Cloud', '2026-06', 531.25), row('DBT Cloud', '2026-07', 531.25), row('DBT Cloud', '2026-08', 531.25), row('DBT Cloud', '2026-09', 102.71),
  // Claude Ai: June, then 0.00 every month (the sheet's zeros are no rows)
  row('Claude Ai', '2026-05', 14513.01), row('Claude Ai', '2026-06', 10092.19),
  // Sprinto: annual, billed this month
  row('Sprinto', '2026-09', 25000, 'Annual'),
  // Hubspot: quarterly
  row('Hubspot', '2026-01', 3796.38, 'Quarterly'), row('Hubspot', '2026-08', 3796.38, 'Quarterly'),
  // Wingman: annual in January, still within a cycle
  row('Wingman', '2026-01', 12189, 'Annual'),
  // Zapier: labelled monthly, nothing since May
  row('Zapier', '2026-05', 31.86),
  // Keepa: one-time
  row('Keepa', '2026-04', 56.66, 'One-time'),
  // Google cloud: monthly, growing
  row('Google cloud', '2026-06', 38686.27), row('Google cloud', '2026-07', 41225.25), row('Google cloud', '2026-08', 34633.23),
];

test('the run-rate reads how an app bills from its charges, over complete months only', () => {
  const by = Object.fromEntries(runRate()(SHEET, '2026-09').map(x => [x.name, x]));
  assert.strictEqual(by['Cursor pro'].pattern, 'monthly', 'charged monthly whatever the label says');
  assert.strictEqual(Math.round(by['Cursor pro'].m * 100) / 100, Math.round((815.04 + 1133.53 + 0) / 3 * 100) / 100, 'August with no charge counts as nothing');
  assert.strictEqual(by['DBT Cloud'].m, 531.25, 'the part-billed current month is not averaged in');
  assert.deepStrictEqual(by['DBT Cloud'].basis, ['2026-06', '2026-07', '2026-08']);
  assert.strictEqual(Math.round(by['Claude Ai'].m * 100) / 100, Math.round(10092.19 / 3 * 100) / 100, 'June, then two months of nothing');
  assert.strictEqual(by['Claude Ai'].pattern, 'monthly');
  assert.strictEqual(by['Google cloud'].m, (38686.27 + 41225.25 + 34633.23) / 3);
});

test('an app on a cycle is its latest charge spread over the cycle, the current month included', () => {
  const by = Object.fromEntries(runRate()(SHEET, '2026-09').map(x => [x.name, x]));
  assert.strictEqual(by['Sprinto'].pattern, 'cycle');
  assert.strictEqual(Math.round(by['Sprinto'].m), Math.round(25000 / 12));
  assert.strictEqual(Math.round(by['Hubspot'].m), Math.round(3796.38 / 3));
  assert.strictEqual(Math.round(by['Wingman'].m), Math.round(12189 / 12));
  assert.ok(!by['Zapier'], 'labelled monthly, nothing for three complete months: lapsed');
  assert.ok(!by['Keepa'], 'a one-time purchase has no run-rate');
  // A cycle overdue by more than a cycle and a month is gone.
  assert.ok(!runRate()([row('Old', '2025-01', 1200, 'Annual')], '2026-09').length);
});

test('the forecast\'s recurring half is that same run-rate', () => {
  const f = pageFunction()(SHEET, [], '2026-09');
  const by = Object.fromEntries(runRate()(SHEET, '2026-09').filter(x => x.pattern === 'monthly').map(x => [x.name, x.m]));
  assert.deepStrictEqual(Object.fromEntries(f.recurring.map(r => [r.name, r.perMonth])), by);
  assert.ok(!f.renewals.some(r => r.name === 'Cursor pro'), 'charged monthly, so not a renewal');
  assert.ok(f.renewals.some(r => r.name === 'Hubspot' && r.month === '2026-11'), 'quarterly from August');
});
