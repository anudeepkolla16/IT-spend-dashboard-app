// Run with: node --test
//
// The Password page shows the workbook's "Invoices mail id" sheet. These tests
// pin how that sheet is found and read: columns by header text (so the sheet
// can be reordered), rows by application name, blank rows and the Total row
// dropped, and the page wired to fetch it only on demand.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const { parseLoginSheet, pickLoginSheet } = require('../api/spend-data.js');

function workbook(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
const SPEND = ['Spendings', [['APPLICATION / SW / LICENSE', 'Department', 'Jan-26'], ['AWS', 'Eng', 1]]];
const HEADER = ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Invoice mail id ', 'Status ', 'Login mail Id ', 'password '];

test('reads every app row of the Invoices mail id sheet with its login', () => {
  const buf = workbook([SPEND, ['Invoices mail id', [
    HEADER,
    ['Adobe', 'Marketing', 'Rakesh', '', '', 'rakesh@example.com', 'Design@Saras26'],
    ['AWS', 'Engineering ', 'Ajay/Santhosh', 'Invoices@example.com', 'Mail Forword From Krishna', '', ''],
    ['Claude Ai', 'org', 'Anudeep', 'Invoices@example.com', '', 'Anudeep mail sso ', ''],
    ['', '', '', '', '', '', ''],
    ['Dovetail'],
    ['Total'],
  ]]]);
  const out = parseLoginSheet(buf);
  assert.strictEqual(out.sheet, 'Invoices mail id');
  assert.deepStrictEqual(out.columns, ['app', 'dept', 'poc', 'invoiceMail', 'status', 'loginMail', 'password']);
  assert.deepStrictEqual(out.rows.map(r => r.app), ['Adobe', 'AWS', 'Claude Ai', 'Dovetail']);
  assert.deepStrictEqual(out.rows[0], {
    app: 'Adobe', dept: 'Marketing', poc: 'Rakesh', invoiceMail: '', status: '',
    loginMail: 'rakesh@example.com', password: 'Design@Saras26',
  });
  assert.strictEqual(out.rows[1].status, 'Mail Forword From Krishna');
  assert.strictEqual(out.rows[1].dept, 'Engineering', 'cells are trimmed');
  assert.strictEqual(out.rows[2].loginMail, 'Anudeep mail sso');
  assert.deepStrictEqual(out.rows[3], { app: 'Dovetail', dept: '', poc: '', invoiceMail: '', status: '', loginMail: '', password: '' });
});

test('finds columns by header text, so a reordered sheet still reads', () => {
  const buf = workbook([SPEND, ['Invoices mail id', [
    ['Note row above the header'],
    ['password', 'Login mail Id', 'APPLICATION / SW / LICENSE', 'Extra'],
    ['Secret1', 'x@example.com', 'Apify', 'ignored'],
  ]]]);
  const out = parseLoginSheet(buf);
  assert.deepStrictEqual(out.columns, ['app', 'loginMail', 'password']);
  assert.deepStrictEqual(out.rows, [{ app: 'Apify', dept: '', poc: '', invoiceMail: '', status: '', loginMail: 'x@example.com', password: 'Secret1' }]);
});

test('the sheet is found by its name, a mail-id-ish name, or LOGIN_SHEET_NAME', () => {
  assert.strictEqual(pickLoginSheet(['Spendings', 'Invoices mail id', 'Invoices tracker']), 'Invoices mail id');
  assert.strictEqual(pickLoginSheet(['Spendings', 'Invoices tracker', 'App logins']), 'App logins');
  assert.strictEqual(pickLoginSheet(['Spendings', 'Invoices tracker']), null);
  process.env.LOGIN_SHEET_NAME = 'Credentials';
  try {
    assert.strictEqual(pickLoginSheet(['Spendings', 'credentials', 'Invoices mail id']), 'credentials');
  } finally {
    delete process.env.LOGIN_SHEET_NAME;
  }
});

test('a workbook without the sheet fails with a message that names the tabs', () => {
  const buf = workbook([SPEND, ['Invoices tracker', [['APPLICATION / SW / LICENSE', 'Jan-26'], ['AWS', 'TRUE']]]]);
  assert.throws(() => parseLoginSheet(buf), /no "Invoices mail id" sheet .*Spendings, Invoices tracker/);
});

test('the dashboard has a Password sidebar item that fetches the sheet on demand', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.match(html, /<button class="nav-item" data-goto="loginsCard">[^<]*<span class="ico">🔑<\/span>Password<\/button>/);
  assert.match(html, /<div class="card" id="loginsCard"/);
  assert.match(html, /\/api\/spend-data\?sheet=logins/);
  // Only on demand: the poll that feeds the dashboard must not carry logins.
  assert.match(html, /fetch\(`\/api\/spend-data\$\{manual \? '\?refresh=1' : ''\}`/);
  // Passwords are masked until revealed, and copy works without revealing.
  assert.match(html, /'••••••••'/);
  assert.match(html, /navigator\.clipboard\.writeText\(r\.password\)/);
});
