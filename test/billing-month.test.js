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

// --- One vendor, rows that bill differently ---------------------------------
//
// Anthropic's seeded rule says `period: "usage"` — read the line items, and ask
// if they say nothing. One answer about one of its rows replaced that with a
// flat vendor-wide "advance", and from then on every Anthropic invoice without
// a readable period was filed as the month it was dated. On 2 October the API
// console's September usage, invoiced on 1 October, went into October. Its
// Claude seats genuinely are billed a month ahead; the console is not, and a
// single field could not hold both.

const { classify } = require('../lib/invoices/rules');

const ANTHROPIC_ROWS = { version: 1, vendors: [{
  name: 'Anthropic', domains: ['anthropic.com'], subject: ['Anthropic'], period: 'usage',
  apps: [
    { app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'], period: 'arrears' },
    { app: 'Claude Ai', text: ['2FSKIDHO'], period: 'advance' },
    { app: 'Claude Ai Max 6 Accounts', text: ['XQRYKLO3'] },
  ],
}] };
const ROW_APPS = ['Anthropic(Api Console)', 'Claude Ai', 'Claude Ai Max 6 Accounts'];

test('a row keeps its own billing convention, whatever the vendor says', () => {
  assert.strictEqual(conventionFor(ANTHROPIC_ROWS, 'Anthropic(Api Console)'), 'arrears');
  assert.strictEqual(conventionFor(ANTHROPIC_ROWS, 'Claude Ai'), 'advance');
  // No convention of its own: the vendor's is the fallback, and "usage" is not
  // a convention — so this row is still asked about, which is the safe answer.
  assert.strictEqual(conventionFor(ANTHROPIC_ROWS, 'Claude Ai Max 6 Accounts'), null);
});

test('a vendor-wide convention still covers the rows that have none', () => {
  const mixed = { version: 1, vendors: [{
    name: 'Anthropic', domains: ['anthropic.com'], subject: [], period: 'advance',
    apps: [{ app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'], period: 'arrears' }, { app: 'Claude Ai', text: ['2FSKIDHO'] }],
  }] };
  assert.strictEqual(conventionFor(mixed, 'Anthropic(Api Console)'), 'arrears', 'the row that was told apart');
  assert.strictEqual(conventionFor(mixed, 'Claude Ai'), 'advance', 'and the vendor for the rest');
});

test('the rules keep a row\'s convention and drop anything else', () => {
  const r = normalizeRules({ vendors: [{ name: 'V', domains: ['v.com'], apps: [
    { app: 'A', text: ['a'], period: 'ARREARS' },
    { app: 'B', text: ['b'], period: 'usage' },   // not a convention for an undated invoice
    { app: 'C', text: ['c'], period: 'whenever' },
  ] }] });
  assert.strictEqual(r.vendors[0].apps[0].period, 'arrears');
  assert.strictEqual(r.vendors[0].apps[1].period, undefined);
  assert.strictEqual(r.vendors[0].apps[2].period, undefined);
});

test('filing an invoice carries the row\'s convention, not the vendor\'s', () => {
  const signals = { address: 'billing@anthropic.com', subject: 'Your Anthropic invoice', attachmentNames: ['Invoice-Q8MUNTUC-0205.pdf'] };
  const v = classify(ANTHROPIC_ROWS, signals, 'Invoice Q8MUNTUC 0205', ROW_APPS);
  assert.strictEqual(v.app, 'Anthropic(Api Console)');
  assert.strictEqual(v.period, 'arrears', 'the console meters usage; it is not billed a month ahead');

  const seats = classify(ANTHROPIC_ROWS, { ...signals, attachmentNames: ['Invoice-2FSKIDHO-0012.pdf'] }, 'Invoice 2FSKIDHO 0012', ROW_APPS);
  assert.strictEqual(seats.app, 'Claude Ai');
  assert.strictEqual(seats.period, 'advance');
});

test('an answer about one row is not learned for the whole vendor', () => {
  // This is the bug: September usage invoiced on 1 October is the console's
  // September. Learning that must not move the Claude seats with it.
  // As the live rule stood before the answer: the console not yet told apart.
  const before = { version: 1, vendors: [{
    name: 'Anthropic', domains: ['anthropic.com'], subject: [], period: 'usage',
    apps: [
      { app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'] },
      { app: 'Claude Ai', text: ['2FSKIDHO'], period: 'advance' },
    ],
  }] };
  const item = { vendor: 'Anthropic', app: 'Anthropic(Api Console)', invoiceDate: '2026-10-01', periodStart: null };
  const out = learnPeriod(before, item, '2026-09');
  const vendor = out.rules.vendors[0];
  assert.strictEqual(vendor.apps[0].period, 'arrears', 'learned on the row');
  assert.strictEqual(vendor.period, 'usage', 'and the vendor is left as it was seeded');
  assert.strictEqual(vendor.apps[1].period, 'advance', 'the seats are untouched');
  assert.match(out.learned, /Anthropic\(Api Console\) is billed in arrears/);
  assert.match(out.learned, /other rows are unaffected/);
});

test('a vendor that bills one row still learns at the vendor', () => {
  // Nothing to tell apart, so the answer belongs to the vendor as before.
  const out = learnPeriod(RULES, { vendor: 'Zapier', app: 'Zapier', invoiceDate: '2026-09-01', periodStart: null }, '2026-09');
  assert.strictEqual(out.rules.vendors[0].period, 'advance');
  assert.match(out.learned, /Zapier bills the month ahead/);
});

test('the live rule, as the bug left it, is repaired by the upgrade', () => {
  const { upgradeRules } = require('../lib/invoices/rules');
  // _vendor-rules.json on 2 October 2026: the seeded "usage" replaced by a flat
  // vendor-wide "advance", learned from one answer, with no row told apart.
  const live = { version: 1, locks: [], locksSeeded: 2, vendors: [{
    name: 'Anthropic', domains: ['anthropic.com'], subject: ['Anthropic'], period: 'advance',
    apps: [
      { app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'] },
      { app: 'Claude Ai', text: ['2FSKIDHO'] },
      { app: 'Claude Ai Max 6 Accounts', text: ['XQRYKLO3'] },
    ],
  }] };
  const { rules, changed } = upgradeRules(live);
  assert.strictEqual(changed, true);
  const vendor = rules.vendors[0];
  assert.strictEqual(vendor.apps[0].period, 'arrears', 'the console gets its own convention from the seed');
  assert.strictEqual(vendor.period, 'advance', "and the owner's vendor-level value is left alone");
  // Which is what actually matters: the console no longer inherits "advance",
  // so a 1 October invoice stating no period is September's, not October's.
  assert.strictEqual(conventionFor(rules, 'Anthropic(Api Console)'), 'arrears');
  assert.strictEqual(conventionFor(rules, 'Claude Ai'), 'advance', 'the seats keep billing a month ahead');
});

test('the upgrade never overwrites a convention already on a row', () => {
  const { upgradeRules } = require('../lib/invoices/rules');
  const live = { version: 1, locks: [], locksSeeded: 2, vendors: [{
    name: 'Anthropic', domains: ['anthropic.com'], subject: [], period: 'usage',
    apps: [{ app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'], period: 'advance' }],
  }] };
  assert.strictEqual(upgradeRules(live).rules.vendors[0].apps[0].period, 'advance', "the owner's value is theirs");
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

test('Recheck Periods re-reads every record an older reader left behind', () => {
  // A record that HAD found a period used to stand regardless of version, on
  // the assumption that a reader only ever finds more. Version 3 broke that:
  // an invoice bracketed by its own previous and next cycles now names no
  // period of its own, so a cached period can be the stale answer too. Only
  // the current version's records stand.
  const periods = cachedReads({ periods: [
    { path: 'a.pdf', read: true, periodStart: '2026-08-01', periodEnd: '2026-08-31' },   // old read, period found: read again
    { path: 'b.pdf', read: true, periodStart: null, invoiceMonth: '2026-09' },           // old read, none found: read again
    { path: 'c.pdf', read: false, note: 'scanned' },                                      // old failure: read again
    { path: 'd.pdf', read: true, periodStart: null, pv: PERIOD_VERSION },                 // current reader: stands
    { path: 'e.pdf', read: true, periodStart: '2026-10-15', pv: PERIOD_VERSION },         // current reader: stands
  ] });
  assert.deepStrictEqual([...periods.keys()].sort(), ['d.pdf', 'e.pdf']);
  assert.ok(PERIOD_VERSION >= 3, 'the Apify records have to be read again');
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

/* ---------- an invoice that settles one cycle and buys the next ----------
 *
 * "Apify invoice is for nov month."
 *
 * Apify's 15 November invoice names two billing periods and says what each one
 * is, relative to itself:
 *
 *   Platform usage in the previous billing period from 2026-10-15 to 2026-11-14
 *   Prepayment for the next billing period    from 2026-11-15 to 2026-12-14
 *
 * The extractor took the first range it found and filed the invoice under
 * October. Reading order was never a reason to prefer one: the invoice sits
 * BETWEEN those periods and names neither as its own, so the row's convention
 * decides — and Apify bills in advance, which puts it in November.
 */

const APIFY_NOV = 'Page 1 of 2  Invoice #202611151493  Date: November 15, 2026  Supplier  Apify Technologies s.r.o. '
  + 'Services  Regular subscription payment Platform usage in the previous billing period from 2026-10-15 to 2026-11-14  '
  + 'Platform usage from 2026-10-15 to 2026-11-14   $41.08  Usage prepaid by the subscription plan   -$29.00  '
  + 'Prepayment for the next billing period from 2026-11-15 to 2026-12-14  Starter plan (monthly)   $29.00  '
  + 'Total amount   41.08 USD  Breakdown of platform usage from 2026-10-15 to 2026-11-14';

const { monthForInvoice } = require('../lib/mail-sync');

test('an invoice bracketed by its own two cycles names no period of its own', () => {
  assert.strictEqual(extractBillingPeriod(APIFY_NOV), null,
    'the first range in the text is not the invoice\'s period');
});

test('so the vendor convention places it, and Apify bills in advance', () => {
  const placed = monthForInvoice(APIFY_NOV, '2026-11', { convention: 'advance', vendor: 'Apify' });
  assert.strictEqual(placed.month, '2026-11', 'November, where the owner files it');
  assert.match(placed.via, /advance/);
});

test('only a document that brackets itself declines to say', () => {
  // One label is a statement, not an ambiguity: an invoice that has just
  // measured a cycle is a plain arrears bill, and one that buys the cycle
  // ahead is a plain prepayment. Each names its period and is read as before.
  const arrearsOnly = 'Invoice  Date: November 15, 2026  Usage in the previous billing period from 2026-10-01 to 2026-10-31  Total 50.00 USD';
  assert.strictEqual(extractBillingPeriod(arrearsOnly).start, '2026-10-01');

  const prepayOnly = 'Invoice  Date: November 15, 2026  Prepayment for the next billing period from 2026-11-15 to 2026-12-14  Total 29.00 USD';
  assert.strictEqual(extractBillingPeriod(prepayOnly).start, '2026-11-15');
});

test('line items with different service windows are not that, and keep their month', () => {
  // dbt's September invoice lists August usage above September seats and says
  // nothing about a previous or next cycle. The test above pins it to August;
  // this one guards against the Apify rule reaching it.
  assert.strictEqual(extractBillingPeriod(DBT_MIXED).start, '2026-08-01');
  // And dbt's November invoice, the same shape, is likewise untouched.
  const dbtNov = 'Invoice number   C70CD8FE   0007  Date of issue   November 1, 2026  '
    + 'Usage - Successful Models Built, Tier 1     0 - 15000    Oct 1   Oct 31, 2026 217   $0.00   $0.00  '
    + 'Seats - Developer License Nov 1   Nov 30, 2026 1   $100.00   $100.00  Total   $106.25';
  assert.ok(extractBillingPeriod(dbtNov), 'it still reads a period');
});

test('a period repeated through one invoice is one period, not several', () => {
  // Apify prints its usage window four times — summary, line item, breakdown
  // heading, breakdown body. Counting repeats as separate periods would make
  // every such invoice look ambiguous.
  const repeated = 'Invoice  Date: November 15, 2026  Period from 2026-10-15 to 2026-11-14  '
    + 'usage from 2026-10-15 to 2026-11-14  Breakdown from 2026-10-15 to 2026-11-14  Total 41.08 USD';
  const p = extractBillingPeriod(repeated);
  assert.ok(p, 'one cycle written three times still names a period');
  assert.strictEqual(p.start, '2026-10-15');
});
