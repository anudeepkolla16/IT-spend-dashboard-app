// Run with: node --test
//
// Which month an invoice belongs to. The fixtures are the text of real invoices
// in the live archive, checked against the month the charge actually falls in.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { extractBillingPeriod, extractInvoiceDate, monthOfPeriod, invoiceMonth } = require('../lib/invoice-period');

// Invoices/Cumul(Luzmo)/20260826_20260258.pdf. Dated Aug 26, due Sep 09, and the
// line item states the cycle it pays for: six days of it fall in August and
// twenty-five in September. The mail arrived on the 26th, which is the only
// reason the sync used to call it Aug-26.
const LUZMO = `Luzmo  77 Sands St · Office 9035 · Brooklyn · NY 11201  Luzmo Inc · support@luzmo.com
Invoice  Date:   Aug 26, 2026  Due date:   Sep 09, 2026  Currency:   USD  Invoice nb.:   US-20260258
Saras Analytics INC. (Market Defense)  2 Corning Fairbanks Way  Unit 207  01581 Westborough  MA  United States
Description   Quantity   Unit price   Sales Tax   Total price
1   Monthly Elite license fee (period from 2026-08-26 until 2026-09-26)   1   $524.50   6.25%   $557.28
2   Additional used monthly active viewers (last month)   0   $0.00   6.25%   $0.00
Total excl. Sales Tax:   $524.50  Total Sales Tax:   $32.78  TOTAL INCL. SALES TAX:   $557.28`;

// Invoices/Adobe/June 26.pdf — a calendar-aligned term, which must not move.
const ADOBE = `INVOICE  Item Details  Service Term: 01-JUN-2026 to 30-JUN-2026
Invoice Information  3474509584 Invoice Number  01-JUN-2026 Invoice Date  USD Currency
65182902   Creative Cloud Pro   1   EA   69.99   34.97   6.25%   2.19   37.16`;

// Invoices/Bubble Starter — nine of these arrive each month, one per application,
// each on its own mid-month cycle, and none of them labels that cycle a period.
// They were reconciled against the month they arrived in and stay there.
const BUBBLE = `Bubble Group, Inc.  Invoice:   2026-08-1218140  Application:   adverio  Date:   8/12/26
DETAIL  Starter Web Plan   8/12/26 - 9/12/26   $32  TOTAL PAID   $32  Amounts are in usd`;

// Invoices/Anthropic(Api Console) — metered usage, stated as bare line dates.
const ANTHROPIC = `Invoice number 8A2F  Date of issue July 31, 2026  Claude Opus 4.8 Usage
Jul 2   Jul 31, 2026   1   13,479.42   13,479.42  Total   13,479.42  Amount due   $13,479.42`;

test('a cycle that straddles two months is filed in the month it starts', () => {
  // The owner's rule: "period is 01-08-2026 to 31-08-2026 and invoice came on
  // 1st Sep but this bill belongs to Aug" — and a 26 Aug → 26 Sep cycle is
  // August's too. The start date decides, however the days fall.
  const r = invoiceMonth(LUZMO, '2026-08');
  assert.strictEqual(r.month, '2026-08');
  assert.strictEqual(r.via, 'period-start');
  assert.strictEqual(r.period.start, '2026-08-26');
  assert.strictEqual(r.period.end, '2026-09-26');
});

test('a period inside one calendar month stays in that month', () => {
  const r = invoiceMonth(ADOBE, '2026-06');
  assert.strictEqual(r.month, '2026-06');
  assert.strictEqual(r.via, 'period-start');
});

test('a line-item date range is the period when nothing labels one', () => {
  // Bubble's "8/12/26 - 9/12/26" is not called a period, but it is one, and
  // the month it starts in is the invoice's — the same month as before.
  assert.strictEqual(extractBillingPeriod(BUBBLE).start, '2026-08-12');
  assert.strictEqual(invoiceMonth(BUBBLE, '2026-08').month, '2026-08');
});

test('metered usage dates are the period, for every vendor', () => {
  // Issued 31 July for "Jul 2 – Jul 31" usage: July's. The same invoice issued
  // on 1 August would still be July's, which is the whole point.
  assert.strictEqual(extractBillingPeriod(ANTHROPIC).start, '2026-07-02');
  assert.strictEqual(invoiceMonth(ANTHROPIC, '2026-07').month, '2026-07');
  assert.strictEqual(invoiceMonth(ANTHROPIC, '2026-08').month, '2026-07');
});

test('an invoice date beside a due date is never read as a period', () => {
  const text = 'Invoice  Date: Aug 26, 2026  Due date: Sep 09, 2026  Currency: USD  Total $557.28';
  assert.strictEqual(extractBillingPeriod(text), null);
  assert.strictEqual(invoiceMonth(text, '2026-08').month, '2026-08');
});

test('an invoice with no readable text keeps the month the mail arrived in', () => {
  assert.strictEqual(invoiceMonth('', '2026-08').month, '2026-08');
  assert.strictEqual(invoiceMonth('', '2026-08').via, 'received');
  assert.strictEqual(invoiceMonth(null, '2026-08').month, '2026-08');
});

// --- Reading the period ---------------------------------------------------

test('the period is read in every shape vendors write it', () => {
  const cases = [
    ['period from 2026-08-26 until 2026-09-26', '2026-08-26', '2026-09-26'],
    ['Billing period: Aug 26 - Sep 26, 2026', '2026-08-26', '2026-09-26'],
    ['Billing Period 26/08/2026 to 26/09/2026', '2026-08-26', '2026-09-26'],
    ['Service period 08/26/2026 - 09/26/2026', '2026-08-26', '2026-09-26'],
    ['Subscription period 1 Aug 2026 through 31 Aug 2026', '2026-08-01', '2026-08-31'],
    ['Service Term: 01-JUN-2026 to 30-JUN-2026', '2026-06-01', '2026-06-30'],
    ['Usage period Jul 1, 2026 – Jul 31, 2026', '2026-07-01', '2026-07-31'],
  ];
  for (const [text, start, end] of cases) {
    const p = extractBillingPeriod(text);
    assert.ok(p, `no period read from "${text}"`);
    assert.strictEqual(p.start, start, text);
    assert.strictEqual(p.end, end, text);
  }
});

test('a range that states its year once carries it back across December', () => {
  const p = extractBillingPeriod('Billing period Dec 20 - Jan 20, 2027');
  assert.strictEqual(p.start, '2026-12-20');
  assert.strictEqual(p.end, '2027-01-20');
});

test('a day-first date is recognised when the day proves it', () => {
  const p = extractBillingPeriod('Billing period 26/08/2026 to 26/09/2026');
  assert.strictEqual(p.start, '2026-08-26'); // not 2026-26-08
});

test('either end of a numeric range can settle the day/month order', () => {
  // 08/09 alone is ambiguous; the 26 at the other end can only be a day, and
  // both ends of one range are written the same way.
  const p = extractBillingPeriod('Billing period 08/09/2026 to 26/09/2026');
  assert.strictEqual(p.start, '2026-09-08');
  assert.strictEqual(p.end, '2026-09-26');
});

test('nonsense dates are left alone rather than guessed at', () => {
  assert.strictEqual(extractBillingPeriod('Billing period 99/99/2026 to 88/88/2026'), null);
  assert.strictEqual(extractBillingPeriod('Period: on receipt to net 30'), null);
});

// --- Choosing the month ---------------------------------------------------

test('a straddling cycle goes to the month it starts in, however the days fall', () => {
  // Aug 26 → Sep 26: six days in August, twenty-five in September — and it is
  // still August's invoice. The breakdown is reported, not decided by.
  const r = monthOfPeriod({ start: '2026-08-26', end: '2026-09-26' });
  assert.strictEqual(r.month, '2026-08');
  assert.strictEqual(r.rule, 'period-start');
  assert.strictEqual(r.days['2026-08'], 6);
  assert.strictEqual(r.days['2026-09'], 25);
});

test('a bill for last month that arrives on the 1st belongs to last month', () => {
  // "period is 01-08-2026 to 31-08-2026 and invoice came on 1st sep but this
  // bill belong to Aug" — the case the rule was written for.
  const r = invoiceMonth('Billing period: 01-08-2026 to 31-08-2026  Total $50.00', '2026-09');
  assert.strictEqual(r.month, '2026-08');
  assert.strictEqual(r.via, 'period-start');
});

test('an even split goes to the month the period starts in', () => {
  const r = monthOfPeriod({ start: '2026-01-17', end: '2026-02-16' });
  assert.strictEqual(r.days['2026-01'], 15);
  assert.strictEqual(r.days['2026-02'], 15);
  assert.strictEqual(r.month, '2026-01');
});

test('a quarterly or annual cycle stays in the month it starts', () => {
  // A year has no majority month, so spreading it is meaningless — an annual
  // invoice is one charge, in the month it was raised.
  assert.strictEqual(monthOfPeriod({ start: '2026-01-01', end: '2026-12-31' }).month, '2026-01');
  assert.strictEqual(monthOfPeriod({ start: '2026-02-15', end: '2026-05-15' }).month, '2026-02');
  assert.strictEqual(invoiceMonth('Subscription period 2026-01-01 to 2026-12-31', '2026-01').month, '2026-01');
});

test('a period far from the mail is treated as a misread, not a cycle', () => {
  // A contract term or a renewal-date pair can read like a period. Anything more
  // than two months from the mail is not this month's invoice cycle.
  const r = invoiceMonth('Contract term 2027-01-01 to 2027-01-31', '2026-08');
  assert.strictEqual(r.month, '2026-08');
  assert.strictEqual(r.via, 'received');
  assert.ok(r.ignoredPeriod);
});

test('the period still wins when the mail arrives late', () => {
  // Luzmo's August invoice forwarded in September still belongs to August.
  assert.strictEqual(invoiceMonth(LUZMO, '2026-09').month, '2026-08');
});

// --- Wiring ---------------------------------------------------------------

test('the mail sync files, ticks and totals by the resolved month', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /monthForInvoice\(pdfText, receivedMonth, /, 'the sync must ask which month the invoice is for');
  // Never the mail's month: the period start, else the invoice date, else ask.
  assert.match(src, /const attMonth = doc\.month;/);
  assert.match(src, /if \(!doc\.app \|\| !doc\.month \|\| doc\.holdWithGroup\) \{/, 'an invoice with no month is held and asked about, not filed by arrival');
  // Nothing downstream may keep using the mail's month: the folder, the tracker
  // tick and the folder total all have to agree with where the PDF went.
  assert.match(src, /placeFor\(attVendor, attMonth\)/, 'the PDF must be filed under the resolved month');
  assert.match(src, /marks\.push\(\{ app: attApp, month: attMonth \}\)/, 'the tracker must tick the resolved month');
  assert.match(src, /touched\.set\(markKey, \{ app: attApp, month: attMonth/, 'the folder total must follow the PDF');
});

test('a copy left in the old month folder is reported, not silently doubled', () => {
  // A rescan re-files an invoice under its period's month, but the copy an
  // earlier run put in the mail's month stays where it is — and both folders
  // are totalled. Nothing is deleted automatically; the leftover is reported.
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /alsoStillAt/, 'a leftover copy in the old month must be reported');
  const ui = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  assert.match(ui, /reroutedByPeriod/, 'the run summary must show what moved month');
  assert.match(ui, /alsoStillAt/, 'and must name the leftover copy to delete');
});


/* ---------------- the date an invoice was issued ---------------- */
//
// For an invoice with no month in its name and no month folder, and which
// states no billing period, the issue date is the only thing that says which
// month it belongs to. It is read narrowly on purpose: the neighbouring due
// date is a different month on Net 30 terms, and a line item's term can span a
// year. A wrong month is worse than no month, so an unrecognised shape yields
// null and the file is left where it is.

test('the invoice date is read, and the due date beside it is not', () => {
  assert.strictEqual(extractInvoiceDate('Invoice date: August 26, 2026  Due date: September 9, 2026').month, '2026-08');
  assert.strictEqual(extractInvoiceDate('Invoice date: 5/31/2025 Due date: 6/30/2025').month, '2025-05');
  assert.strictEqual(extractInvoiceDate('Date paid September 2, 2025').month, '2025-09');
});

test('a due date on its own is never taken as the invoice date', () => {
  assert.strictEqual(extractInvoiceDate('Due date: 6/30/2025'), null);
  assert.strictEqual(extractInvoiceDate('Payment due 6/30/2025 Net 30'), null);
});

test("a line item's own term is not the invoice's date", () => {
  // ClickUp's annual plan runs 5/26/2025 to 5/17/2026. Reading either end as
  // the invoice date would file a May bill under the wrong year.
  assert.strictEqual(extractInvoiceDate('Product name Term start Term end 5/26/2025 5/17/2026'), null);
});

test('a date with no year cannot be placed and is refused', () => {
  // A period borrows the missing year from the other end of its range; a single
  // date has no other end.
  assert.strictEqual(extractInvoiceDate('Invoice date: Aug 26'), null);
});

test('a flattened invoice table is paired by column order, never by adjacency', () => {
  // Real text from the archive's ClickUp invoices: the PDF's table collapses to
  // a run of labels followed by a run of values, so the date next to "Invoice
  // date" is the wrong column's. The k-th date label goes with the k-th date.
  const flattened = 'Invoice number:   Purchase order:   Invoice date:   Due date:   Terms:   Amount due:  '
    + 'INV50264   5/31/2025   6/30/2025   Net 30   USD $0.00';
  const r = extractInvoiceDate(flattened);
  assert.strictEqual(r.month, '2025-05');
  assert.strictEqual(r.paired, true);
});

test('Stripe and Google headers, which come out as labels then values or values then labels', () => {
  // The three invoices the first live run held for want of a month: each one
  // states its date, in a column the reader separated from its label.
  const stripe = 'Invoice  Invoice number  Date of issue  Date due  Q8MUNTUC-0204  September 2, 2026  September 2, 2026  Anthropic, PBC';
  assert.strictEqual(extractInvoiceDate(stripe).month, '2026-09');
  const google = 'Invoice  Invoice number: 5571764174  5571764174  Apr 30, 2026  7476-5812-7596  sarasanalytics.com  Details  Invoice number  Invoice date  Billing ID  Domain Name  $13.87';
  assert.strictEqual(extractInvoiceDate(google).month, '2026-04');
  // Labels in the other order still pair by position, not by nearest date.
  assert.strictEqual(extractInvoiceDate('Date due  Date of issue  September 9, 2026  September 2, 2026').month, '2026-09');
  assert.strictEqual(extractInvoiceDate('Date due  Date of issue  October 9, 2026  September 2, 2026').month, '2026-09');
});

test('usage lines with the dash lost are read as the period, with no opt-in', () => {
  const t = 'Date of issue August 1, 2026  Claude Haiku 4.5 Usage Jul 2   Jul 31, 2026 1   $425.42';
  const r = invoiceMonth(t, '2026-08');
  assert.strictEqual(r.month, '2026-07');
  assert.strictEqual(r.period.start, '2026-07-02');
});

test('Google\'s "Summary for" is a billing period', () => {
  assert.strictEqual(invoiceMonth('Summary for Apr 1, 2026 - Apr 30, 2026  Subtotal in USD', '2026-05').month, '2026-04');
});

test('an invoice with no date at all yields nothing', () => {
  assert.strictEqual(extractInvoiceDate('Thanks for your business. Total 42.00'), null);
  assert.strictEqual(extractInvoiceDate(''), null);
  assert.strictEqual(extractInvoiceDate(null), null);
});
