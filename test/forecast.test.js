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
    slice('function monthsBetween(', 'function trailingRunRateByApp('),
    slice('/* ---------------- forecast:', '/* ---------------- end forecast'),
  ].join('\n');
  return new Function(src + '; return buildForecast;')();
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
  assert.deepStrictEqual(f.basis, []);
  assert.strictEqual(f.months.length, 6);
  assert.ok(f.months.every(m => m.total === 0));
});

test('the page no longer draws a straight-line trend', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.ok(!/linreg/.test(html));
  assert.match(html, /buildForecast\(all, window\.SHEET_ROWS \|\| \[\], realMonthKey\(0\)\)/);
});
