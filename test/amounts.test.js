// Run with: node --test
//
// The important claim these tests make is the reconciliation one: parsing a
// statement laid out like the real June 2026 file and aggregating it must
// reproduce the amounts a human already typed into the live spend sheet.

const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('xlsx');

const { parseStatement, monthFor, monthsCovered } = require('../lib/statement');
const { buildResolver } = require('../lib/vendor-map');
const { locateGrid, cellAddress, colLetter, normMonthHeader } = require('../lib/excel');

// App names exactly as the live sheet spells them, trailing spaces included.
const APPS = [
  'Adobe', 'AWS', 'Anthropic(Api Console)', 'Apify', 'Bubble Starter', 'Chargebee',
  'Claude Ai Max 6 Accounts', 'Cursor pro', 'DBT Cloud', 'Envato', 'ElevenLabs',
  'Filestack', 'Github', 'GOOGLE ADS', 'Google Workspace', 'Google cloud',
  'Granola Business', 'Helpjuice', 'Hex', 'Linkedin', 'MetalPriceAPI', 'Pagerduty',
  'Phantombuster', 'Product Fruits', 'Prosp AI', 'Render ', 'Sentry.io', 'Superhuman',
  'TMobile', 'Twilo Sendgrid', 'Typeform', 'WEBFLOW', 'windsurf pro', 'ZOHO Books', 'ZOOM',
];

function buildJuneStatement() {
  const s1 = [['BOA ', 'Date', 'Month', 'Description', 'Amount', 'Comments']];
  const add = (d, desc, amt, cmt) => s1.push(['BOA-0218', d, '26-Jun', desc, amt, cmt]);

  add('6/2/2026', 'ADOBE *ADOBE 06/01 PURCHASE 408-536-6000 CA DEBIT CARD *4154', -37.16, 'Adobe');
  add('6/12/2026', 'AMAZON WEB SERVI DES:INTERNET ID:043000092186264 CO ID:9049016352 CCD', -2525.33, 'Amazon Web services');
  // Bubble Starter totals 751.96 in the live sheet across 11 rows, two of which
  // carry a May purchase date in the descriptor.
  add('6/1/2026', 'BUBBLE STARTER PLAN 05/29 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/1/2026', 'BUBBLE STARTER PLAN 05/31 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -431.96, 'Bubble Starter');
  add('6/11/2026', 'BUBBLE STARTER PLAN 06/10 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'bubble Starter');
  add('6/11/2026', 'BUBBLE STARTER PLAN 06/10 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/11/2026', 'BUBBLE STARTER PLAN 06/10 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/15/2026', 'BUBBLE STARTER PLAN 06/12 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/24/2026', 'BUBBLE STARTER PLAN 06/23 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/25/2026', 'BUBBLE STARTER PLAN 06/25 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/25/2026', 'BUBBLE STARTER PLAN 06/25 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/29/2026', 'BUBBLE STARTER PLAN 06/27 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  add('6/30/2026', 'BUBBLE STARTER PLAN 06/29 PURCHASE BUBBLE.IO NY DEBIT CARD *4154', -32, 'Bubble Starter');
  // DBT Cloud totals 531.25
  add('6/2/2026', 'DBT CLOUD SUBSCRIPTION 06/01 PURCHASE GETDBT.COM PA DEBIT CARD *4154', -212.5, 'DBT Cloud');
  add('6/2/2026', 'DBT CLOUD SUBSCRIPTION 06/01 PURCHASE GETDBT.COM PA DEBIT CARD *4154', -106.25, 'DBT Cloud');
  add('6/2/2026', 'DBT CLOUD SUBSCRIPTION 06/01 PURCHASE GETDBT.COM PA DEBIT CARD *4154', -106.25, 'DBT Cloud');
  add('6/2/2026', 'DBT CLOUD SUBSCRIPTION 06/02 PURCHASE GETDBT.COM PA DEBIT CARD *4154', -106.25, 'DBT Cloud');
  add('6/1/2026', 'Chargebee 05/31 PURCHASE 187-7303113 CA DEBIT CARD *4154', -158.91, 'Chargebee');
  add('6/15/2026', 'LinkedIn*P3032739084 06/12 PURCHASE Sunnyvale CA DEBIT CARD *4154', -18.85, 'Likedin');
  add('6/17/2026', 'LinkedIn SN P3035681257 06/16 PURCHASE 855-6535653 CA DEBIT CARD *4154', -127.49, 'Linkedin');
  add('6/30/2026', 'SENTRY 06/29 PURCHASE SENTRY.IO CA DEBIT CARD *4154', -78.16, 'Sentry');
  add('6/2/2026', 'ZOHO* ZOHO-BOOKS 06/01 PURCHASE WWW.ZOHO.COM CA DEBIT CARD *4154', -101.77, 'Zoho books');
  add('6/8/2026', 'PADDLE.NET* EXTENSIONS 06/06 PURCHASE New York NY DEBIT CARD *4154', -10.61, 'Paddle Net');
  add('6/10/2026', 'ENVATO 06/09 PURCHASE 613-837-6628 UT DEBIT CARD *4154', -33, 'Envato');
  add('6/20/2026', 'ENVATO REFUND 06/19 PURCHASE 613-837-6628 UT DEBIT CARD *4154', 12, 'Envato');

  // Second layout: positive amounts, "Statment Period", no Comments column.
  const s2 = [['Transaction Code', 'Statment Period', 'Date', 'Description', 'Amount', 'Cost Sheet', 'Bill Status', 'Transaction Status', 'Transaction  type']];
  const add2 = (date, desc, amt) => s2.push(['BOA/2026/8292/0277', '26-Jun', date, desc, amt, '', 'Bill required', 'Recorded', 'Apps & Subscriptions']);
  add2('6/2/2026', 'ANTHROPIC* CLAUDE SUB SAN FRANCISCO CA', 118);
  add2('5/12/2026', 'GOOGLE *CLOUD 9Q4THP 6502530000 CA', 10000);
  add2('6/1/2026', 'Google CLOUD L9CwWG 6502530000 CA', 8686.27);
  add2('6/11/2026', 'Google CLOUD M2JBHZ 6502530000 CA', 10000);
  add2('5/26/2026', 'Google CLOUD R68HwJ 6502530000 CA', 10000);
  add2('5/19/2026', 'WEBFLOW.COM 41 59640555 CA', 497.25);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s1), 'Sheet1');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s2), 'Sheet2');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function aggregate(buffer, attribution) {
  const { txns } = parseStatement(buffer);
  const resolve = buildResolver({}, APPS);
  const totals = {};
  const unmapped = {};
  for (const t of txns) {
    const r = resolve(t.vendorLabel, t.description);
    const month = monthFor(t, attribution);
    if (!r.app) {
      unmapped[t.vendorLabel || t.description] = (unmapped[t.vendorLabel || t.description] || 0) + t.amount;
      continue;
    }
    const key = `${r.app}@${month}`;
    totals[key] = (totals[key] || 0) + t.amount;
  }
  return { totals, unmapped, txns };
}

test('reproduces the amounts already in the live spend sheet', () => {
  const { totals } = aggregate(buildJuneStatement(), 'statement');
  // Every figure below was read off the live sheet's Jun-26 column.
  const expected = {
    'Adobe@2026-06': 37.16,
    'AWS@2026-06': 2525.33,
    'Bubble Starter@2026-06': 751.96,
    'DBT Cloud@2026-06': 531.25,
    'Google cloud@2026-06': 38686.27,
    'Claude Ai Max 6 Accounts@2026-06': 118,
  };
  for (const [key, want] of Object.entries(expected)) {
    assert.ok(totals[key] !== undefined, `${key} missing from the parsed totals`);
    assert.ok(Math.abs(totals[key] - want) < 0.005, `${key}: expected ${want}, got ${totals[key]}`);
  }
});

test('merges hand-typed vendor label variants onto one app', () => {
  const { totals } = aggregate(buildJuneStatement(), 'statement');
  // "Bubble Starter" and "bubble Starter" are one app; "Likedin" is "Linkedin".
  assert.ok(Math.abs(totals['Bubble Starter@2026-06'] - 751.96) < 0.005);
  assert.ok(Math.abs(totals['Linkedin@2026-06'] - 146.34) < 0.005);
});

test('nets refunds down instead of counting them as spend', () => {
  const { totals } = aggregate(buildJuneStatement(), 'statement');
  // A 33.00 charge and a 12.00 refund net to 21.00, not 45.00.
  assert.ok(Math.abs(totals['Envato@2026-06'] - 21) < 0.005, `got ${totals['Envato@2026-06']}`);
});

test('maps the sheet that has no Comments column from the bank descriptor', () => {
  const { totals } = aggregate(buildJuneStatement(), 'statement');
  // "ANTHROPIC* CLAUDE SUB" is the seat subscription, a different row from the
  // API console — the more specific descriptor rule has to win.
  assert.strictEqual(totals['Claude Ai Max 6 Accounts@2026-06'], 118);
  assert.strictEqual(totals['Anthropic(Api Console)@2026-06'], undefined);
});

test('leaves an unrecognised vendor unmapped rather than guessing', () => {
  const { unmapped } = aggregate(buildJuneStatement(), 'statement');
  assert.ok(Object.keys(unmapped).some(k => /Paddle/i.test(k)), 'Paddle Net should be unmapped');
});

test('attribution rule changes which month a late-posting charge lands in', () => {
  const statement = aggregate(buildJuneStatement(), 'statement').totals;
  const transaction = aggregate(buildJuneStatement(), 'transaction').totals;
  // Under the statement rule everything on the June statement is June.
  assert.ok(Math.abs(statement['Bubble Starter@2026-06'] - 751.96) < 0.005);
  // Under the transaction rule the two May-dated purchases move to May.
  assert.ok(Math.abs(transaction['Bubble Starter@2026-05'] - 463.96) < 0.005);
  assert.ok(Math.abs(transaction['Bubble Starter@2026-06'] - 288) < 0.005);
});

test('reports the months a statement covers', () => {
  const { txns } = aggregate(buildJuneStatement(), 'statement');
  const months = monthsCovered(txns, 'statement');
  assert.strictEqual(months[0].month, '2026-06');
});

test('locates the grid and computes cell addresses from the sheet layout', () => {
  const values = [
    ['Spendings'],
    ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Renewal data', 'Recurring/Onetime', 'FREQUENCY', 'Payment Method', 'Jan-26', 'Feb-26', 'Mar-26'],
    ['Adobe', 'Marketing', 'Bhavana', '', 'Recurring', 'Monthly', 'US Debit Card', 37.16, 37.16, 37.16],
    ['AWS', 'Engineering', 'Ajay', '', 'Recurring', 'Monthly', 'US Debit Card', 4552.64, 2673.27, 2116.53],
    ['Total', '', '', '', '', '', '', 60241.72, 89287.23, 99363],
  ];
  const grid = locateGrid(values);
  assert.strictEqual(grid.headerRowIdx, 1);
  assert.strictEqual(grid.nameCol, 0);
  assert.strictEqual(grid.apps.length, 2, 'the Total row must not be treated as an app');
  assert.strictEqual(grid.monthCols['2026-03'], 9);
  // AWS x Mar-26 -> column J, row 4.
  assert.strictEqual(cellAddress({ col: 0, row: 0 }, grid.apps[1].rowIdx, grid.monthCols['2026-03']), 'J4');
});

test('handles month headers stored as real dates', () => {
  assert.strictEqual(normMonthHeader('Jan-26'), '2026-01');
  assert.strictEqual(normMonthHeader('01/01/2026'), '2026-01');
  assert.strictEqual(normMonthHeader('not a month'), null);
  assert.strictEqual(colLetter(26), 'AA');
});
