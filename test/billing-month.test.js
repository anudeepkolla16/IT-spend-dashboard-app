// Run with: node --test
//
// "So many Aug invoices filed under Sep just because the invoice was generated
// in Sep — completely wrong." Every one of those invoices said which month it
// was for: AWS in a sentence, Render under a label, Anthropic, dbt and Apollo
// on their line items. These tests pin the texts as they came out of the
// archive, and the rule for an invoice that truly states no period: the
// vendor's remembered billing convention, or a question — never the date alone.

const test = require('node:test');
const assert = require('node:assert');

const { extractBillingPeriod, previousMonth } = require('../lib/invoice-period');
const { normalizeRules, learnPeriod, conventionFor, conventionOf } = require('../lib/invoices/rules');
const { mergeEntries } = require('../lib/mail-sync');
const { cachedReads, PERIOD_VERSION } = require('../lib/invoices/period-backfill');

const AWS = 'Invoice Summary  Invoice Number:   2789999141  Invoice Date:   September 1 , 2026  TOTAL AMOUNT DUE ON September 11 , 2026   USD 2,374.66  This invoice is for the billing period August 1 - August 31 , 2026  Greetings from Amazon Web Services';
const RENDER = 'Invoice number   IXPFOJDB   0005  Date of issue   September 1, 2026 Date due   September 1, 2026 Billing period   Aug 1 - Aug 31, 2026 Team name   My Workspace  Render Aug 1, 2026 - Aug 31, 2026  Workspace Subscription';
const ANTHROPIC = 'Date of issue   September 1, 2026 Date due   October 1, 2026  Anthropic, PBC  Bill to  Saras analytics INC 92 Ruggles St Westborough, Massachusetts 01581   2121 United States  $16,376.29 USD due October 1, 2026  Description   Qty   Unit price   Tax   Amount  Claude Haiku 4.5 Usage Aug 1   Aug 31, 2026 1   $775.48983425   6.25%   $775.49';
const DBT_MIXED = 'Date of issue   September 1, 2026 Date due   September 1, 2026  dbt Labs  915 Spring Garden St Suite 500 Philadelphia  Semantic Layer Usage - Queried Metrics, Tier 1     0 - 5000    Aug 1   Aug 31, 2026 347   $0.00   $0.00  Seats - Developer License Sep 1   Sep 30, 2026 2   $100.00   $200.00';
const DBT_SEATS = 'Date of issue   September 2, 2026 Date due   September 2, 2026  Sales Tax   1   $6.04   $6.04  Seats - Developer License (prorated) Sep 2   Sep 30, 2026 1   $96.67';
const APOLLO = 'Date of issue   September 2, 2026 Date due   September 2, 2026  ZenLeads Inc. (dba Apollo.io)  440 N Barranca Ave #4750 Covina, California 91723  Basic Seat Sep 2, 2026 – Sep 2, 2027 1   $588.00   $588.00  Credits Sep 2, 2026 – Sep 2, 2027 10,000';
const ZAPIER = 'PAID  Zapier Inc. 548 Market St #62411 San Francisco, CA 94104-5401  Invoice  1 September 2026  Description   Amount  Zapier - Pro 750 (monthly)   $29.99 USD  Total   $31.86 USD  Card:   visa   ****   4154  Reference Code:   01a05c9b-83b0-15a3-5398-0b02ecbc718b';

test('the September invoices that were August\'s say so, each in its own way', () => {
  assert.strictEqual(extractBillingPeriod(AWS).start, '2026-08-01', 'AWS: "August 31 , 2026" with a space before the comma');
  assert.strictEqual(extractBillingPeriod(RENDER).start, '2026-08-01', 'Render: a labelled period');
  assert.strictEqual(extractBillingPeriod(ANTHROPIC).start, '2026-08-01', 'Anthropic: usage lines, dash lost');
  assert.strictEqual(extractBillingPeriod(DBT_MIXED).start, '2026-08-01', 'dbt: last month\'s usage comes before next month\'s seats');
});

test('the September invoices that really are September\'s stay there', () => {
  assert.strictEqual(extractBillingPeriod(DBT_SEATS).start, '2026-09-02');
  const apollo = extractBillingPeriod(APOLLO);
  assert.strictEqual(apollo.start, '2026-09-02', 'an annual term starts in September');
  assert.strictEqual(apollo.end, '2027-09-02');
});

test('an invoice that states no period at all reads as none', () => {
  assert.strictEqual(extractBillingPeriod(ZAPIER), null);
});

test('two dates that are not a range are never read as one', () => {
  // A PDF's date-of-issue and due-date columns, side by side.
  assert.strictEqual(extractBillingPeriod('Date of issue Date due September 1, 2026 October 1, 2026 Total $5'), null);
  assert.strictEqual(extractBillingPeriod('Invoice  Date: Aug 26, 2026  Due date: Sep 09, 2026  Total $557.28'), null);
  assert.strictEqual(extractBillingPeriod('Invoice date 09/01/2026 Due date 10/01/2026 Net 30'), null);
  assert.strictEqual(extractBillingPeriod('Call 410-555-1234 or 1-800-123-4567 Suite 500 - 1 Page 1 of 2 Account 12-34-5678 ZIP 01581-2121'), null);
});

test('previousMonth steps back across a year end', () => {
  assert.strictEqual(previousMonth('2026-09'), '2026-08');
  assert.strictEqual(previousMonth('2027-01'), '2026-12');
  assert.strictEqual(previousMonth('nope'), null);
});

// --- The vendor's billing convention ---------------------------------------

const RULES = { version: 1, vendors: [
  { name: 'Zapier', domains: ['mail.zapier.com'], subject: [], app: 'Zapier' },
  { name: 'Anthropic', domains: ['anthropic.com'], subject: ['Anthropic'], period: 'usage', apps: [{ app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'] }] },
  { name: 'AWS', domains: ['amazon.com'], subject: [], app: 'AWS', period: 'arrears' },
] };

test('the rules keep advance and arrears, and usage says nothing about undated invoices', () => {
  const r = normalizeRules({ vendors: [{ name: 'X', domains: ['x.com'], app: 'X', period: 'ARREARS' }, { name: 'Y', domains: ['y.com'], app: 'Y', period: 'whenever' }] });
  assert.strictEqual(r.vendors[0].period, 'arrears');
  assert.strictEqual(r.vendors[1].period, undefined);
  assert.strictEqual(conventionOf('usage'), null);
  assert.strictEqual(conventionOf('advance'), 'advance');
  assert.strictEqual(conventionFor(RULES, 'AWS'), 'arrears');
  assert.strictEqual(conventionFor(RULES, 'Anthropic(Api Console)'), null);
  assert.strictEqual(conventionFor(RULES, 'Zapier'), null);
  assert.strictEqual(conventionFor(RULES, 'Nobody'), null);
});

test('the owner\'s month for an undated invoice is remembered as how the vendor bills', () => {
  const item = { vendor: 'Zapier', invoiceDate: '2026-09-01', periodStart: null };
  const advance = learnPeriod(RULES, item, '2026-09');
  assert.strictEqual(advance.rules.vendors[0].period, 'advance');
  assert.match(advance.learned, /Zapier bills the month ahead/);

  const arrears = learnPeriod(RULES, { vendor: 'Anthropic', invoiceDate: '2026-09-01', periodStart: null }, '2026-08');
  assert.strictEqual(arrears.rules.vendors[1].period, 'arrears', 'replaces "usage", which said nothing about this case');
  assert.match(arrears.learned, /Anthropic bills in arrears/);

  // Not learned: a month that is neither, an invoice with a period, no date, an unknown vendor, or nothing new.
  assert.strictEqual(learnPeriod(RULES, item, '2026-06').learned, null);
  assert.strictEqual(learnPeriod(RULES, { ...item, periodStart: '2026-09-01' }, '2026-09').learned, null);
  assert.strictEqual(learnPeriod(RULES, { ...item, invoiceDate: null }, '2026-09').learned, null);
  assert.strictEqual(learnPeriod(RULES, { ...item, vendor: 'Ghost' }, '2026-09').learned, null);
  assert.strictEqual(learnPeriod(RULES, { vendor: 'AWS', invoiceDate: '2026-09-01' }, '2026-08').learned, null);
  // The original rules are untouched.
  assert.strictEqual(RULES.vendors[0].period, undefined);
});

// --- Housekeeping the fix leans on ------------------------------------------

test('the filing log keeps one line per file and month, the latest', () => {
  const merged = mergeEntries(
    [{ app: 'DBT Cloud', month: '2026-09', file: 'a.pdf', monthVia: 'invoice-date' }, { app: 'Zapier', month: '2026-09', file: 'z.pdf' }],
    [{ app: 'DBT Cloud', month: '2026-09', file: 'a.pdf', monthVia: 'answered' }, null, { app: 'DBT Cloud', month: '2026-08', file: 'a.pdf' }]
  );
  assert.deepStrictEqual(merged.map(e => `${e.app}|${e.month}|${e.file}|${e.monthVia || ''}`), [
    'Zapier|2026-09|z.pdf|', 'DBT Cloud|2026-09|a.pdf|answered', 'DBT Cloud|2026-08|a.pdf|',
  ]);
  assert.deepStrictEqual(mergeEntries(undefined, undefined), []);
});

test('Recheck Periods re-reads a cached record that found no period under the old reader', () => {
  const periods = cachedReads({ periods: [
    { path: 'a.pdf', read: true, periodStart: '2026-08-01', periodEnd: '2026-08-31' },   // old read, period found: stands
    { path: 'b.pdf', read: true, periodStart: null, invoiceMonth: '2026-09' },          // old read, none found: read again
    { path: 'c.pdf', read: false, note: 'scanned' },                                     // old failure: read again
    { path: 'd.pdf', read: true, periodStart: null, pv: PERIOD_VERSION },                // current reader: stands
  ] });
  assert.deepStrictEqual([...periods.keys()].sort(), ['a.pdf', 'd.pdf']);
  assert.strictEqual(PERIOD_VERSION, 2);
});

test('the sync passes the vendor\'s convention, and the backfill gets the rules', () => {
  const fs = require('fs');
  const path = require('path');
  const sync = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(sync, /monthForInvoice\(pdfText, receivedMonth, \{ convention: rulesLib\.conventionOf\(verdict\.period\), vendor: verdict\.vendor \}\)/);
  assert.match(sync, /rulesLib\.learnPeriod\(rules, item, month\)/);
  assert.ok(!/usageRange/.test(sync), 'the usage opt-in is gone: every vendor\'s line items are read');
  const cron = fs.readFileSync(path.join(__dirname, '..', 'api', 'invoices', 'sync-cron.js'), 'utf8');
  assert.match(cron, /scanPeriods\(token, targetDriveId, \{[^}]*rules \}\)/);
});
