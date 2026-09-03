// Run with: node --test
//
// Fixtures are verbatim text extracted from real invoices in the live archive,
// not invented samples. Each one is checked against the figure that is actually
// in the spend sheet for that app and month.

const test = require('node:test');
const assert = require('node:assert');
const { extractInvoiceTotal, detectCurrency } = require('../lib/invoice-amount');

// Invoices/Bubble Starter — sheet has no Aug-26 figure yet; invoice says $32.
const BUBBLE = `Bubble Group, Inc.  22 W 21st Street  2nd Floor  New York, NY 10010  United States of America
Bubble Support Center  bubble.io  Invoice:   2026-08-1218140  Application:   adverio  Date:   8/12/26
Saras Analytics LLC  92 Ruggles Street Massachusetts Westborough, MA 01581  audit@sarasanalytics.com
DETAIL  Starter Web Plan   8/12/26 - 9/12/26   $32  TOTAL PAID   $32
Amounts are in usd - payment by credit card ending in 4154.  Thank you - we really appreciate your business.`;

// Invoices/Adobe/June 26.pdf — the sheet's Adobe Jun-26 cell holds 37.16.
// Note the pre-tax "NET AMOUNT (USD) 34.97" appears BEFORE the grand total.
const ADOBE = `Bill To  Harsha Anathneni  MA 01581  INVOICE  Item Details  Service Term: 01-JUN-2026 to 30-JUN-2026
Invoice Information  3474509584 Invoice Number  01-JUN-2026 Invoice Date  Credit Card Payment Terms
USD Currency  Adobe Inc.  345 Park Avenue  San Jose CA 95110-2704  United States
PRODUCT NUMBER   PRODUCT DESCRIPTION   QUANTITY   UNIT   UNIT PRICE   TOTAL DISCOUNT AMOUNT/UNIT   NET AMOUNT   TAX RATE   TAXES   TOTAL
65182902   Creative Cloud Pro   1   EA   69.99   (35.02)   34.97   6.25%   2.19   37.16
Invoice Total   NET AMOUNT (USD)   34.97  TAXES (SEE DETAILS FOR RATES)   2.19  GRAND TOTAL (USD)   37.16  Comments:`;

// Invoices/MICROSOFT(Tata Tele)/June 2026.pdf — invoice is INR 150591.60, while
// the sheet correctly holds 1574.40 (the converted figure). Writing the face
// value would be a ~95x error.
const TATA = `TATA TELE NXTGEN SOLUTIONS LTD  TAX INVOICE - ORIGINAL  BILL DETAILS
Invoice No.   :   20260636I182669 Invoice Date   :   28-06-2026 Payment Due Date   :   16-07-2026
Billed To:  SARAS SOLUTIONS INDIA PRIVATE LIMITED
S. No.   Item   HSN   Price (INR)   UoM   Quantity   Sub Total (INR)   Total (INR)
1   Microsoft 365 Business Standard with Email Security   997331   604.00   Monthly   186   112344.00   112344.00
2   Microsoft 365 Business basic with Email Security   997331   114.00   Monthly   134   15276.00   15276.00
Sub Total (INR)   127620.00 IGST @ 18.00% (INR)   22971.60 Net Payable (INR)   150591.60`;

// Invoices/Cumul(Luzmo)/20260826_20260258.pdf — the invoice that exposed this.
// Its totals block is three lines that all begin "Total", and only the last is
// what is owed. Which of them a bare /total/ pattern reached first depended on
// how the PDF's table came out as text, so all three layouts are checked.
const LUZMO_TOTALS = {
  inline: 'Total excl. Sales Tax: $524.50  Total Sales Tax: $32.78  TOTAL INCL. SALES TAX: $557.28',
  stacked: 'Total excl. Sales Tax:\n$ 524.50\nTotal Sales Tax:\n$ 32.78\nTOTAL INCL. SALES TAX:\n$ 557.28',
  columns: 'Total excl. Sales Tax:   $ 524.50\nTotal Sales Tax:   $ 32.78\nTOTAL INCL. SALES TAX:   $ 557.28\n',
};

test('the payable total wins over the pre-tax subtotal and the tax itself', () => {
  for (const [layout, text] of Object.entries(LUZMO_TOTALS)) {
    const r = extractInvoiceTotal(text);
    assert.strictEqual(r.amount, 557.28, `${layout} layout read ${r.amount}`);
    assert.strictEqual(r.usable, true, layout);
  }
});

test('the sales tax on its own is never read as a total', () => {
  // 32.78 against a real 557.28 is a 17x error, and it would have gone into the
  // sheet as that month's Luzmo spend.
  const r = extractInvoiceTotal('Total Sales Tax: $32.78');
  assert.strictEqual(r.amount, null);
  assert.strictEqual(r.usable, false);
});

test('an invoice that only states a pre-tax total still yields that', () => {
  // Refusing the tax line must not refuse everything: some invoices print no
  // other figure, and the one they do print beats reporting nothing.
  const r = extractInvoiceTotal('Total excl. Sales Tax: $524.50');
  assert.strictEqual(r.amount, 524.5);
  assert.strictEqual(r.via, 'total excl. tax');
});

test('reads the total from a simple USD invoice', () => {
  const r = extractInvoiceTotal(BUBBLE);
  assert.strictEqual(r.amount, 32);
  assert.strictEqual(r.currency, 'USD');
  assert.strictEqual(r.usable, true);
});

test('prefers the grand total over the pre-tax net amount', () => {
  const r = extractInvoiceTotal(ADOBE);
  // 37.16 is what the sheet holds; 34.97 is the pre-tax figure printed first.
  assert.strictEqual(r.amount, 37.16, `picked ${r.amount} — the pre-tax 34.97 must not win`);
  assert.strictEqual(r.currency, 'USD');
  assert.strictEqual(r.usable, true);
});

test('refuses to offer an INR invoice as a USD figure', () => {
  const r = extractInvoiceTotal(TATA);
  // It still reads the number, so a human can see it — but never as usable.
  assert.strictEqual(r.amount, 150591.60);
  assert.strictEqual(r.currency, 'INR');
  assert.strictEqual(r.usable, false, 'an INR total must never be written into a USD sheet');
  assert.match(r.note, /INR/);
});

test('declared currency beats a stray symbol', () => {
  assert.strictEqual(detectCurrency('USD Currency ... total $50').code, 'USD');
  assert.strictEqual(detectCurrency('Net Payable (INR) 150591.60').code, 'INR');
  assert.strictEqual(detectCurrency('Amounts are in usd - payment by credit card').code, 'USD');
});

test('reports an unreadable or scanned PDF rather than guessing', () => {
  const r = extractInvoiceTotal('');
  assert.strictEqual(r.usable, false);
  assert.strictEqual(r.amount, null);
  assert.match(r.note, /scan/i);
});

test('reports an invoice with no recognisable total', () => {
  const r = extractInvoiceTotal('Thank you for your business. Your subscription is active. USD');
  assert.strictEqual(r.amount, null);
  assert.strictEqual(r.usable, false);
  assert.strictEqual(r.via, 'no-total');
});

test('does not mistake a thousands separator for a decimal point', () => {
  const r = extractInvoiceTotal('Currency USD\nGrand Total   12,345.67');
  assert.strictEqual(r.amount, 12345.67);
});

test('will not offer a figure when the currency is ambiguous', () => {
  const r = extractInvoiceTotal('Total Paid  £100  also shown as €120');
  assert.strictEqual(r.usable, false);
});

// --- Which Spendings cells an invoice total may fill ---------------------

const { planAmountCells } = require('../lib/mail-sync');
const { locateGrid, cellAddress } = require('../lib/excel');
const { cellValue } = require('../lib/spend-sheet');

// Shaped like the live Spendings sheet: date-serial headers, a Total row.
const VALUES = [
  ['Spendings'],
  ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Renewal data', 'Recurring/Onetime', 'FREQUENCY', 'Payment Method', 46023, 46174, 46235],
  ['Adobe', 'Marketing', 'Bhavana', '', 'Recurring', 'Monthly', 'US Debit Card', 37.16, 37.16, ''],
  ['Bubble Starter', 'Product', 'Ganesh', '', 'Recurring', 'Monthly', 'US Debit Card', 320, 751.96, ''],
  ['MICROSOFT(Tata Tele)', 'org', 'Anudeep', '', 'Recurring', 'Monthly', 'HDFC bank', 1651.42, 1574.4, ''],
  ['Total', '', '', '', '', '', '', 2008.58, 2363.52, ''],
];
const TEXT = VALUES.map((row, i) => row.map((c, j) => {
  if (i === 1 && j >= 7) return ['Jan-26', 'Jun-26', 'Aug-26'][j - 7];
  return c == null ? '' : String(c);
}));
const USED = { values: VALUES, start: { col: 0, row: 0 } };
const grid = () => locateGrid(VALUES, TEXT);

test('fills an empty month cell from the invoice total', () => {
  const { write, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-08': 256 } }, grid(), USED, cellValue
  );
  assert.strictEqual(write.length, 1);
  assert.strictEqual(write[0].address, 'J4'); // Bubble Starter row 4, Aug-26 column J
  assert.strictEqual(write[0].value, 256);
  assert.strictEqual(skippedFilled.length, 0);
});

test('updates a filled cell when the invoice total has grown', () => {
  // Jun-26 holds 751.96; a new invoice takes the folder total to 999.
  const { write, updated } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 999 } }, grid(), USED, cellValue, {}
  );
  assert.strictEqual(write.length, 0);
  assert.strictEqual(updated.length, 1, 'a grown total must reach the sheet');
  assert.strictEqual(updated[0].previous, 751.96);
  assert.strictEqual(updated[0].value, 999);
});

test('a lower invoice total is a question, never a write', () => {
  // The owner has said twice that a lowered cell was wrong: the archive was
  // short, not the sheet. So a lower total is held and asked about; only an
  // answer (or a lock) lowers a cell.
  const { write, updated, skippedFilled, lowered } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 500 } }, grid(), USED, cellValue, {}
  );
  assert.strictEqual(write.length + updated.length + skippedFilled.length, 0);
  assert.strictEqual(lowered.length, 1);
  assert.strictEqual(lowered[0].current, 751.96);
  assert.strictEqual(lowered[0].invoiceTotal, 500);
});

test('an answered "use the invoices" lets that one lowering through', () => {
  const { updated, lowered } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 500 } }, grid(), USED, cellValue, {}, {}, [], new Set(['Bubble Starter||2026-06'])
  );
  assert.strictEqual(lowered.length, 0);
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].value, 500);
  assert.strictEqual(updated[0].direction, 'down');
});

test('never writes to the Total row or an unknown app', () => {
  const { write } = planAmountCells(
    { 'Total': { '2026-08': 5 }, 'Nonexistent App': { '2026-08': 5 } }, grid(), USED, cellValue
  );
  assert.strictEqual(write.length, 0);
});

test('ignores a month the sheet has no column for', () => {
  const { write } = planAmountCells({ 'Adobe': { '2027-01': 40 } }, grid(), USED, cellValue);
  assert.strictEqual(write.length, 0);
});

test('ignores a zero or negative total', () => {
  const { write } = planAmountCells({ 'Adobe': { '2026-08': 0 } }, grid(), USED, cellValue);
  assert.strictEqual(write.length, 0);
});

test('an INR invoice never reaches the planner at all', () => {
  // extractInvoiceTotal marks it unusable, so it is never summed into `amounts`
  // — this asserts the boundary that keeps a 150591.60 out of a USD cell.
  const r = extractInvoiceTotal(TATA);
  assert.strictEqual(r.usable, false);
  const { write } = planAmountCells(
    r.usable ? { 'MICROSOFT(Tata Tele)': { '2026-08': r.amount } } : {},
    grid(), USED, cellValue
  );
  assert.strictEqual(write.length, 0);
});

// --- Reading the PDF itself ---------------------------------------------
//
// pdf-parse v2 wraps modern pdf.js and needs browser globals (DOMMatrix,
// Path2D) plus @napi-rs/canvas. On Vercel it threw ReferenceError at REQUIRE
// time, killing the whole function before any handler could catch it — invoice
// filing and the folder mirror went down with it. The library is pinned to 1.x,
// required by its inner path, and loaded lazily behind a guard.

const fs = require('node:fs');
const path = require('node:path');
const { readPdfText } = require('../lib/invoice-amount');

test('reads text out of a real PDF on a plain Node runtime', async () => {
  const sample = path.join(__dirname, '..', 'node_modules', 'pdf-parse', 'test', 'data', '01-valid.pdf');
  if (!fs.existsSync(sample)) return; // sample not shipped; nothing to assert
  const { text, error } = await readPdfText(fs.readFileSync(sample));
  assert.strictEqual(error, null, `reader failed: ${error}`);
  assert.ok(text.length > 0, 'expected some extracted text');
});

test('never throws on a corrupt PDF — reports and moves on', async () => {
  const { text, error } = await readPdfText(Buffer.from('this is definitely not a pdf'));
  assert.strictEqual(text, '');
  assert.ok(error, 'a corrupt PDF must report an error rather than throw');
});

test('an unreadable PDF yields no amount, so filing still proceeds', async () => {
  const { text, error } = await readPdfText(Buffer.from('nonsense'));
  const total = error ? { amount: null, usable: false } : extractInvoiceTotal(text);
  assert.strictEqual(total.amount, null);
  assert.strictEqual(total.usable, false);
});

test('does not require the PDF library at module load', () => {
  // Loading lib/invoice-amount.js must not pull pdf-parse in — that is what
  // made a library incompatibility fatal to the whole function.
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'invoice-amount.js'), 'utf8');
  const topLevel = src.split('\n').filter(l => /^\s*(const|let|var)\s.*require\(['"]pdf-parse/.test(l));
  assert.strictEqual(topLevel.length, 0, 'pdf-parse must only be required lazily, inside readPdfText');
});

// --- Totalling the folder, not just the mail ----------------------------
//
// A live run reported "Bubble Starter 2026-08: invoice says 64.00" against an
// actual 524.27. Only 2 of Bubble's 9 August charges arrived by email; the rest
// reached the archive by other routes. Summing only emailed attachments
// undercounts silently, and would have written 64.00 into an empty cell.
// The total now comes from every invoice in the app-month folder.

test('the message loop no longer accumulates amounts by itself', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  // The per-attachment accumulator was the bug; the folder pass replaces it.
  assert.ok(!/byMonth\[month\] = Math\.round\(\(\(byMonth\[month\] \|\| 0\) \+ total\.amount\)/.test(src),
    'amounts must not be summed per emailed attachment');
  assert.match(src, /sumFolderInvoices/, 'the folder-wide total must be what feeds the sheet');
});

test('the folder pass is bounded so one run cannot parse without limit', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /maxParse/, 'a parse budget must bound the run');
  assert.match(src, /budget\.deadline/, 'the folder pass must respect the run deadline');
});

test('parsed totals are cached by path so a PDF is read once', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /cache\.set\(key, entry\)/);
  assert.match(src, /amounts: existing\.concat\(budget\.fresh\)/,
    'fresh results must persist to the index for the next run');
});


// --- Topping up as later invoices arrive --------------------------------
//
// Invoices for a month arrive across it — Bubble's ninth August invoice lands on
// the 28th. Writing once and never revisiting would leave the cell permanently
// short while looking final. A cell this sync wrote, and which still holds
// exactly what it wrote, may be topped up; anything else must not be.

test('tops up a cell this sync wrote when later invoices arrive', () => {
  // We wrote 492.27 earlier; the ninth invoice takes the folder to 524.27.
  const prior = { 'Bubble Starter||2026-06': 492.27 };
  const values = VALUES.map(r => r.slice());
  values[3][8] = 492.27; // Bubble Starter, Jun-26 column
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 524.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, prior
  );
  assert.strictEqual(write.length, 0);
  assert.strictEqual(skippedFilled.length, 0);
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].previous, 492.27);
  assert.strictEqual(updated[0].value, 524.27);
});

test('a figure a human put there, above the invoices, is held and asked about', () => {
  const values = VALUES.map(r => r.slice());
  values[3][8] = 524.27;
  const { write, updated, lowered } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 492.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}
  );
  assert.strictEqual(write.length + updated.length, 0);
  assert.strictEqual(lowered.length, 1);
  assert.strictEqual(lowered[0].wasOurs, false);
});

test('a hand correction is still superseded by a higher invoice total', () => {
  // We wrote 492.27; someone corrected it to 500; the ninth invoice takes the
  // folder to 524.27. The author asked for amounts to keep up with invoices, so
  // the grown total wins — but `wasOurs` marks it as a figure we did not write,
  // so the summary and the audit log can show whose number was replaced.
  const prior = { 'Bubble Starter||2026-06': 492.27 };
  const values = VALUES.map(r => r.slice());
  values[3][8] = 500;
  const { updated } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 524.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, prior
  );
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].previous, 500);
  assert.strictEqual(updated[0].wasOurs, false, 'must be flagged as not our own figure');
});

test('a hand correction above the invoice total stands until the owner answers', () => {
  const prior = { 'Bubble Starter||2026-06': 492.27 };
  const values = VALUES.map(r => r.slice());
  values[3][8] = 600; // corrected upward by a human
  const { write, updated, lowered } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 524.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, prior
  );
  assert.strictEqual(write.length + updated.length, 0);
  assert.strictEqual(lowered.length, 1);
  assert.strictEqual(lowered[0].current, 600);
});

test('does not rewrite a cell that already holds the right total', () => {
  const prior = { 'Bubble Starter||2026-06': 524.27 };
  const values = VALUES.map(r => r.slice());
  values[3][8] = 524.27;
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 524.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, prior
  );
  assert.strictEqual(write.length + updated.length + skippedFilled.length, 0, 'nothing to do');
});

// --- Anthropic: one sender, two rows in the sheet ------------------------
//
// Verbatim text from the two July 2026 Anthropic invoices in the archive. Both
// arrive from Anthropic, PBC with near-identical subjects and filenames, so only
// the body distinguishes the API console from the Claude seat subscriptions.

const ANTHROPIC_API_JULY = `Page 1 of 2  Invoice  Invoice number   Q8MUNTUC   0203  Date of issue   August 1, 2026
Anthropic, PBC   (@anthropic)  548 Market Street San Francisco, California 94104  support@anthropic.com
Bill to  Saras analytics  $13,479.42 USD due August 31, 2026
Description   Qty   Unit price   Tax   Amount
Claude Haiku 4.5 Usage Jul 2   Jul 31, 2026 1   $425.4220351   6.25%   $425.42
Claude Opus 4.8 Usage Jul 2   Jul 31, 2026 1   $1,362.52184825   6.25%   $1,362.52
Claude Sonnet 5 Usage Jul 2   Jul 31, 2026 1   $5,863.7163777   6.25%   $5,863.72
Subtotal   $12,686.51 Total excluding tax   $12,686.51 Tax - Massachusetts 6.25% $792.91 Total   $13,479.42
Amount due   $13,479.42 USD`;

const ANTHROPIC_SEATS_JULY = `Page 1 of 2  Receipt  Invoice number   2FSKIDHO   0088  Receipt number   2699 1972 3418
Date paid   August 1, 2026  Anthropic, PBC   (@anthropic)  support@anthropic.com
Bill to  Saras Analytics  $0.00 paid on August 1, 2026
Description   Qty   Unit price   Tax   Amount
Extra usage units, Enterprise plan Jul 1   Jul 31, 2026 1   $4,007.391289   6.25%   $4,007.39
Chat Team Billing Adjustment Credit applied Jul 1   Jul 31, 2026 1   $0.75   6.25%   $0.75
Auto recharge extra usage, Team plan applied Jul 1   Jul 31, 2026 1   $113.727874   6.25%   $113.73
Subtotal   $3,799.90 Total excluding tax   $3,799.90 Tax - Massachusetts 6.25% $237.49 Total   $4,037.39
Applied balance     $4,037.39`;

test('reads the API console invoice total, matching the sheet', () => {
  const r = extractInvoiceTotal(ANTHROPIC_API_JULY);
  assert.strictEqual(r.amount, 13479.42, 'the sheet holds 13,479.42 for Jul-26');
  assert.strictEqual(r.usable, true);
});

test('reads nothing charged when an applied balance settled the invoice', () => {
  // The receipt prints Total $4,037.39 but $0.00 paid — the sheet holds 0.00.
  // Taking the Total here would be wrong by the entire amount.
  const r = extractInvoiceTotal(ANTHROPIC_SEATS_JULY);
  assert.strictEqual(r.amount, 0, `took ${r.amount}; the 4,037.39 total was settled from balance`);
  assert.match(r.via, /balance applied/);
});

test('an ordinary receipt is not read as zero on a stray "paid"', () => {
  // No applied balance, so the total stands.
  const r = extractInvoiceTotal('Bubble Group  TOTAL PAID   $32  Amounts are in usd  paid on 8/12/26');
  assert.strictEqual(r.amount, 32);
});

const { refineAnthropic } = require('../lib/vendor-map');
const ANTHROPIC_ROWS = ['Anthropic(Api Console)', 'Claude Ai', 'Claude Ai Max 6 Accounts', 'Adobe'];

test('routes a usage invoice to the API console row', () => {
  assert.strictEqual(refineAnthropic(ANTHROPIC_API_JULY, ANTHROPIC_ROWS), 'Anthropic(Api Console)');
});

test('routes a plan or seat invoice to the Claude row', () => {
  assert.strictEqual(refineAnthropic(ANTHROPIC_SEATS_JULY, ANTHROPIC_ROWS), 'Claude Ai');
});

test('falls back to the amount threshold only when the wording says nothing', () => {
  const vague = 'Anthropic, PBC   Invoice   Amount due   $12,000.00 USD';
  assert.strictEqual(refineAnthropic(vague, ANTHROPIC_ROWS), 'Anthropic(Api Console)');
  const small = 'Anthropic, PBC   Invoice   Amount due   $250.00 USD';
  assert.strictEqual(refineAnthropic(small, ANTHROPIC_ROWS), 'Claude Ai');
});

test('leaves non-Anthropic invoices alone', () => {
  assert.strictEqual(refineAnthropic(BUBBLE, ANTHROPIC_ROWS), null);
});

/* ---------------- the cell is the month's invoice total ---------------- *
 *
 * The rule the owner set: "calculate the total for all invoices in that month
 * and show that in sheet". The cell holds the sum of the month's invoices,
 * replacing whatever is there, and nothing is ever ADDED to a cell — that is
 * what makes a re-run harmless: the same folder writes the same number twice,
 * never twice the number. Bubble's June is the fixture: eight invoices
 * totalling 492.27, a ninth of 32.00 still to come.
 */

const bubble = (total) => planAmountCells(
  { 'Bubble Starter': { '2026-06': total } }, grid(), USED, cellValue
);

test('sum: the cell follows the invoices on file upward; a shortfall is a question', () => {
  // Eight of nine invoices on file; the cell holds 751.96 from somewhere else.
  const eight = bubble(492.27);
  assert.strictEqual(eight.write.length + eight.updated.length, 0);
  assert.strictEqual(eight.lowered.length, 1, 'held and asked, not lowered');

  // The folder passes the sheet: the cell is raised.
  const more = bubble(800);
  assert.strictEqual(more.updated.length, 1);
  assert.strictEqual(more.updated[0].value, 800);
  assert.strictEqual(more.updated[0].direction, 'up');
});

test('sum: a cell already holding the month\'s total is not written again', () => {
  const same = bubble(751.96);
  assert.strictEqual(same.write.length, 0, 'the cell is already correct');
  assert.strictEqual(same.updated.length, 0, 'and must not be written again');
  assert.strictEqual(same.skippedFilled.length, 0);
});

test('sum: re-running changes nothing, so a repeat run cannot inflate a cell', () => {
  const first = bubble(900);
  assert.strictEqual(first.updated[0].value, 900);

  // Second run, same folder, but the cell now holds what the first run wrote.
  const AFTER = { values: VALUES.map(r => [...r]), start: { col: 0, row: 0 } };
  AFTER.values[3][8] = 900;
  const second = planAmountCells(
    { 'Bubble Starter': { '2026-06': 900 } }, grid(), AFTER, cellValue
  );
  assert.strictEqual(second.write.length, 0);
  assert.strictEqual(second.updated.length, 0);
  assert.strictEqual(second.skippedFilled.length, 0);
});

/* ---------------- a partial folder total is not a total ---------------- *
 *
 * Apollo's August folder held two invoices and one of them would not parse.
 * The 85.00 that did parse looked bigger than the 53.12 already in the cell, so
 * the sync raised the cell to 85.00 — replacing a figure it had never written
 * with one that is missing an invoice. "Has the folder grown past the sheet?"
 * is not a question a lower bound can answer.
 */

const apolloGrid = () => grid();
const PARTIAL = { 'Adobe||2026-06': true };

test('a partial folder total never overwrites a figure already in the cell', () => {
  // Adobe's Jun-26 cell holds 37.16. A half-read folder totalling 85.00 is
  // higher, but that proves nothing.
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Adobe': { '2026-06': 85 } }, apolloGrid(), USED, cellValue, {}, PARTIAL
  );
  assert.strictEqual(write.length, 0);
  assert.strictEqual(updated.length, 0, 'the cell must not be raised from a lower bound');
  assert.strictEqual(skippedFilled.length, 1);
  assert.strictEqual(skippedFilled[0].reason, 'folder-total-incomplete');
  assert.strictEqual(skippedFilled[0].current, 37.16);
  assert.strictEqual(skippedFilled[0].invoiceTotal, 85);
});

test('the same total DOES apply once every invoice in the folder has been read', () => {
  // Nothing about the number changed — only whether it is the whole folder.
  const { updated } = planAmountCells(
    { 'Adobe': { '2026-06': 85 } }, apolloGrid(), USED, cellValue, {}, {}
  );
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].previous, 37.16);
  assert.strictEqual(updated[0].value, 85);
});

test('a partial total does not fill an empty cell either — it is held and asked about', () => {
  // A cell claims to be the month's total, and a lower bound is not that.
  // "Never guess": the month is reported with the file that would not read,
  // and the cell is written on the run after it is fixed.
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Adobe': { '2026-08': 85 } }, apolloGrid(), USED, cellValue, {}, { 'Adobe||2026-08': true }
  );
  assert.strictEqual(write.length + updated.length, 0);
  assert.strictEqual(skippedFilled.length, 1);
  assert.strictEqual(skippedFilled[0].reason, 'folder-total-incomplete');
  assert.strictEqual(skippedFilled[0].current, null);
});

test('a complete total is not flagged as partial', () => {
  const { write } = planAmountCells(
    { 'Adobe': { '2026-08': 85 } }, apolloGrid(), USED, cellValue, {}, {}
  );
  assert.strictEqual(write[0].partial, false);
});

/* -------- an invoice and its receipt are one charge, not two -------- */

const { extractInvoiceRef } = require('../lib/invoice-amount');

// Verbatim from Apollo's August folder. Both documents were filed, both total
// $85.00, and the receipt is the payment FOR that invoice. Summing every PDF in
// the folder makes Apollo's August 170.00 — double the truth. Only one of the
// two parsed, so the live run happened to report 85.00 and hid it.
const APOLLO_INVOICE = `Page 1 of 1  Invoice  Invoice number   A0589F17   0017  Date of issue   August 27, 2026 Date due   August 27, 2026  ZenLeads Inc. (dba Apollo.io)  Bill to  Saras Analytics INC  $85.00 USD due August 27, 2026  Subtotal   $85.00 Total   $85.00  Amount due   $85.00 USD`;
const APOLLO_RECEIPT = `Page 1 of 1  Receipt  Invoice number   A0589F17   0017  Receipt number   2601   5895 Date paid   August 27, 2026  ZenLeads Inc. (dba Apollo.io)  $85.00 paid on August 27, 2026  Subtotal   $85.00 Total   $85.00  Amount paid   $85.00`;

test('an invoice and its receipt resolve to the same reference', () => {
  assert.strictEqual(extractInvoiceRef(APOLLO_INVOICE), 'A0589F170017');
  assert.strictEqual(extractInvoiceRef(APOLLO_RECEIPT), 'A0589F170017',
    'the receipt carries the invoice number too, which is what pairs them');
});

test("a receipt's own receipt number is never used as the reference", () => {
  // "Receipt number 2601 5895" sits right after the invoice number on the page.
  // Capturing it would give the two documents different keys and defeat the
  // whole thing.
  assert.ok(!extractInvoiceRef(APOLLO_RECEIPT).includes('2601'));
});

test('both documents still read the same total, so either one is correct alone', () => {
  assert.strictEqual(extractInvoiceTotal(APOLLO_INVOICE).amount, 85);
  assert.strictEqual(extractInvoiceTotal(APOLLO_RECEIPT).amount, 85);
});

test('a reference is only used when one was really read', () => {
  // Without a reference, two files sharing an amount are two charges — a vendor
  // billing the same figure twice in a month is ordinary.
  assert.strictEqual(extractInvoiceRef('Receipt for your payment. Total $85.00'), null);
  assert.strictEqual(extractInvoiceRef('Invoice # 7'), null, 'too short to be a reference');
  assert.strictEqual(extractInvoiceRef(''), null);
});

test('common invoice-number formats are read', () => {
  assert.strictEqual(extractInvoiceRef('Invoice No: INV-2026-0042  Date: 1 Aug'), 'INV20260042');
  assert.strictEqual(extractInvoiceRef('Invoice Number: 38639008'), '38639008');
});

/* ---- a month's invoices are not all inside that month's folder ---- */

const { monthFromFileName, detectDayFirst } = require('../lib/invoices/inventory');

// Apollo's real August, from their billing page: two paid invoices,
// in_0U0dee5… for 53.12 on 4 Aug and in_0U91fV5… for 85.00 on 27 Aug — 138.12.
// On disk, only the 27th's invoice is in Aug-26/. The 4th's sits loose in the
// vendor folder as "Invoice-A0589F17-0016-Aug 2026.pdf". Totalling the month
// folder alone gave 85.00, which looked bigger than the 53.12 in the cell and
// overwrote it — losing the 4 August charge outright.
test('a loose invoice is dated to the month its name states', () => {
  const created = '2026-08-27T10:00:00Z';
  assert.strictEqual(
    monthFromFileName('Invoice-A0589F17-0016-Aug 2026.pdf', created, null), '2026-08');
  // The July one in the same folder must not be pulled into August.
  assert.strictEqual(
    monthFromFileName('Invoice-A0589F17-0015-july 2026.pdf', created, null), '2026-07');
});

test("the vendor folder's own date order settles its loose names", () => {
  // Apollo's names carry an explicit month word, so no numeric order is needed
  // and none is inferred — detectDayFirst has nothing unambiguous to go on.
  const names = ['Invoice-A0589F17-0015-july 2026.pdf', 'Invoice-A0589F17-0016-Aug 2026.pdf'];
  assert.strictEqual(detectDayFirst(names), null);
  assert.strictEqual(monthFromFileName(names[1], '2026-08-27T10:00:00Z', null), '2026-08');
});

test('the two Apollo invoices are distinct charges, not a duplicate pair', () => {
  // 0016 and 0017 are different invoice numbers, so the receipt-pairing rule
  // must not collapse them: August really is 53.12 + 85.00.
  const inv16 = 'Invoice  Invoice number   A0589F17   0016  Date of issue   August 4, 2026  Total   $53.12  Amount due   $53.12 USD';
  const inv17 = 'Invoice  Invoice number   A0589F17   0017  Date of issue   August 27, 2026  Total   $85.00  Amount due   $85.00 USD';
  assert.notStrictEqual(extractInvoiceRef(inv16), extractInvoiceRef(inv17));
  assert.strictEqual(extractInvoiceTotal(inv16).amount, 53.12);
  assert.strictEqual(extractInvoiceTotal(inv17).amount, 85);
  assert.strictEqual(
    Math.round((extractInvoiceTotal(inv16).amount + extractInvoiceTotal(inv17).amount) * 100) / 100,
    138.12, "Apollo's real August total");
});

test('the totaller accepts invoices from outside the month folder', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /async function looseFilesForMonth\(/,
    'loose invoices have to be found before they can be totalled');
  assert.match(src, /sumFolderInvoices\(token, driveId, folder, cache, budget, loose\)/,
    'and passed to the totaller, or the month is short by whatever sits loose');
});

/* ---------------- what the archive survey turned up ---------------- *
 *
 * Forty archived invoices had "no payable total found" and two were reported
 * as being in the currencies AND and ANY. Each shape below is verbatim from a
 * real file in the archive (as the connector reads it), and each was a gap in
 * the patterns rather than a bad PDF.
 */

test('Sentry: "Total $82.31 USD" — the figure is followed by its currency', () => {
  const r = extractInvoiceTotal('Total   $82.31 USD  Your subscription renews');
  assert.strictEqual(r.amount, 82.31);
  assert.strictEqual(r.usable, true);
});

test('Webflow: pdf-parse glues the label to the figure ("TotalUSD 816.00")', () => {
  const r = extractInvoiceTotal('Subtotal USD 768.00 Tax (6.25%) USD 48.00 TotalUSD 816.00 Amount DueUSD 816.00');
  assert.strictEqual(r.amount, 816);
  assert.strictEqual(r.currency, 'USD');
});

test('Google Voice: a column of figures beside a column of labels', () => {
  const r = extractInvoiceTotal('$10.00  $0.73  $0.19  $1.45  $1.50  $13.87  Subtotal in USD  State sales tax (6.25%)  Federal Regulatory Assessment Fee  Federal Universal Service Fund  State 911 Tax  Total in USD  Domain Name: x');
  assert.strictEqual(r.amount, 13.87);
  assert.strictEqual(r.via, 'total in CUR (columns)');
});

test('Google Ads statement: Total in USD is the month\'s spend, not the payments', () => {
  const r = extractInvoiceTotal('$598.69  $0.00  $598.69  Subtotal in USD  Tax (0%)  Total in USD  -$460.00 Total payments received in USD');
  assert.strictEqual(r.amount, 598.69);
  assert.strictEqual(extractInvoiceTotal('Subtotal in USD $10.00  Total in USD $13.87').amount, 13.87, 'and inline, when the reader keeps the row together');
});

test('a figure/label pairing that does not line up is not trusted', () => {
  // Three figures, two labels: whichever is missing, the last figure is not
  // necessarily the total.
  assert.strictEqual(extractInvoiceTotal('$1.00  $2.00  $3.00  Subtotal in USD  Total in USD').amount, null);
});

test('PostHog: a $0.00 invoice is a real total of nothing, not an unread one', () => {
  const r = extractInvoiceTotal('$0.00 USD due April 24, 2026  Subtotal $0.00 Total $0.00  Amount due $0.00 USD');
  assert.strictEqual(r.amount, 0);
  assert.strictEqual(r.usable, true);
  // But a nought never beats a real figure elsewhere on the invoice.
  assert.strictEqual(extractInvoiceTotal('Credit $0.00  Total $12.00  Amount due $12.00 USD').amount, 12);
});

test('"any currency" and "currency and" are not currencies', () => {
  assert.strictEqual(detectCurrency('Pay in any currency you like. Total $5.00').code, 'USD');
  assert.strictEqual(detectCurrency('Currency and payment terms. Total $5.00').code, 'USD');
  assert.strictEqual(detectCurrency('the errors were ours. Total $5.00').code, 'USD', 'lower-case "rs" is not rupees');
});

test('a vendor rule can say which of two stated currencies the sheet takes', () => {
  const both = 'Total EUR 50.00 (USD 55.00) Total 55.00';
  assert.strictEqual(extractInvoiceTotal(both).usable, false, 'ambiguous without a rule');
  const r = extractInvoiceTotal(both, { currency: 'USD' });
  assert.strictEqual(r.usable, true);
  assert.strictEqual(r.currency, 'USD');
  assert.strictEqual(extractInvoiceTotal('Net Payable (INR) 150591.60', { currency: 'USD' }).usable, false, 'a rule cannot turn rupees into dollars');
});

test('positioned text comes back as lines, with cells on a row kept together', () => {
  const { linesFromTextContent } = require('../lib/invoice-amount');
  const item = (str, x, y, w) => ({ str, transform: [10, 0, 0, 10, x, y], width: w });
  const lines = linesFromTextContent({ items: [
    item('Total', 20, 100, 25), item('USD 816.00', 200, 100.5, 50),
    item('Sub', 20, 120, 15), item('total', 35.5, 120, 25),
  ] });
  assert.deepStrictEqual(lines, ['Subtotal', 'Total   USD 816.00']);
});

/* ---------------- what the first live run exposed ---------------- */

test('a Stripe account prefix on its own is not an invoice number', () => {
  // Five Cursor invoices in one month all read as "WYO2F9CO" when the number
  // broke across a line, and four were dropped as duplicates of the first.
  const { extractInvoiceRef } = require('../lib/invoice-amount');
  assert.strictEqual(extractInvoiceRef('Invoice number WYO2F9CO\n0101\nDate of issue'), 'WYO2F9CO0101');
  assert.strictEqual(extractInvoiceRef('Invoice number   Q8MUNTUC   0200  Date of issue'), 'Q8MUNTUC0200');
  assert.strictEqual(extractInvoiceRef('Invoice number WYO2F9CO Date of issue July 1 Reference WYO2F9CO 0101'), 'WYO2F9CO0101', 'the sequence is picked up from wherever the page states it');
  assert.strictEqual(extractInvoiceRef('Invoice number INV50264'), 'INV50264');
});

test('a PDF holding several invoices is totalled as all of them', () => {
  // Anthropic's "July 26.pdf": three credit top-ups and the usage invoice in
  // one file. Reading it as one document took the first Total, 138.82.
  const four = 'Invoice number   Q8MUNTUC   0200  Date of issue   July 1, 2026 Total   $138.82  Amount due   $138.82 USD'
    + '  Invoice number   Q8MUNTUC   0201 Total   $138.31  Amount due   $138.31 USD'
    + '  Invoice number   Q8MUNTUC   0202 Total   $139.28  Amount due   $139.28 USD'
    + '  Invoice number   Q8MUNTUC   0203  Date of issue   August 1, 2026 Subtotal   $12,686.51 Total   $13,479.42  Amount due   $13,479.42 USD';
  const r = extractInvoiceTotal(four);
  assert.strictEqual(r.amount, 13895.83);
  assert.strictEqual(r.usable, true);
  assert.deepStrictEqual(r.refs, ['Q8MUNTUC0200', 'Q8MUNTUC0201', 'Q8MUNTUC0202', 'Q8MUNTUC0203']);
  assert.match(r.via, /4 invoices in one file/);
});

test('an invoice number read as the total is refused', () => {
  const r = extractInvoiceTotal('Invoice number: 5685830900  Total 5685830900');
  assert.strictEqual(r.usable, false);
  assert.match(r.note, /not a plausible charge/);
});

test('a month emptied by a move is cleared only when the figure was this sync\'s own', () => {
  const prior = { 'Bubble Starter||2026-06': 751.96 };
  const { updated } = planAmountCells({ 'Bubble Starter': { '2026-06': 0 } }, grid(), USED, cellValue, prior);
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].value, 0);
  assert.strictEqual(updated[0].emptied, true);
  const theirs = planAmountCells({ 'Bubble Starter': { '2026-06': 0 } }, grid(), USED, cellValue, {});
  assert.strictEqual(theirs.updated.length + theirs.write.length, 0, 'a figure not ours stays');
});

/* ---------------- the reader, end to end ---------------- *
 *
 * The first live run on Vercel read nothing: pdfjs loads its worker through
 * an eval'd require that no bundler follows, so the function shipped without
 * pdf.worker.js. The fixture is a small invoice-shaped PDF; this pins that the
 * worker resolves, that pdfjs is the reader that answers, and that its text
 * keeps a label and its figure on one line.
 */

test('pdfjs reads a PDF end to end, with the worker it needs on disk', async () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const { readPdfText } = require('../lib/invoice-amount');
  assert.ok(require.resolve('pdfjs-dist/legacy/build/pdf.worker.js'), 'the worker must be resolvable');
  const vercel = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'vercel.json'), 'utf8'));
  assert.match(String(vercel.functions['api/invoices/sync-cron.js'].includeFiles), /pdf\.worker\.js/, 'and shipped with the function');

  const r = await readPdfText(fs.readFileSync(path.join(__dirname, 'fixtures', 'sample-invoice.pdf')));
  assert.strictEqual(r.error, null);
  assert.strictEqual(r.reader, 'pdfjs');
  assert.match(r.text, /Total\s+USD 816\.00/);
  assert.strictEqual(extractInvoiceTotal(r.text).amount, 816);
});


/* ---------------- cells the owner has locked ---------------- *
 *
 * "never change these amounts in next runs, these are correct figures":
 * some months have invoices missing from the archive, and Claude Ai's and
 * Cursor's charges were paid from prepaid credits. A lock in the rules file
 * holds the owner's figure; the sync writes it back if the cell drifts and
 * ignores the folder total for that month.
 */

const LOCKS = [{ app: 'Bubble Starter', month: '2026-06', value: 745.89, note: 'invoices missing' }];

test('a locked cell is never set to the folder total', () => {
  const values = VALUES.map(r => r.slice());
  values[3][8] = 745.89;
  const { write, updated, skippedFilled, locked } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 192 } }, locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}, {}, LOCKS
  );
  assert.strictEqual(write.length + updated.length + skippedFilled.length, 0);
  assert.strictEqual(locked.length, 1);
  assert.strictEqual(locked[0].enforced, false);
  assert.strictEqual(locked[0].invoiceTotal, 192, 'the ignored total is reported beside the kept figure');
});

test('a locked cell that has drifted is put back to the locked figure', () => {
  // The cell holds 192 (a previous run's total); the lock says 745.89.
  const values = VALUES.map(r => r.slice());
  values[3][8] = 192;
  const { updated, locked } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 192 } }, locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}, {}, LOCKS
  );
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].value, 745.89);
  assert.strictEqual(updated[0].locked, true);
  assert.strictEqual(locked[0].enforced, true);
});

test('a lock of zero clears a figure the sync wrote, and holds it there', () => {
  const values = VALUES.map(r => r.slice());
  values[3][8] = 4473.45;
  const zero = [{ app: 'bubble starter', month: '2026-06', value: 0, note: 'paid from credits' }];
  const { updated } = planAmountCells({ 'Bubble Starter': { '2026-06': 4473.45 } }, locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}, {}, zero);
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].value, 0);
  const values2 = VALUES.map(r => r.slice());
  values2[3][8] = '';
  const again = planAmountCells({ 'Bubble Starter': { '2026-06': 4473.45 } }, locateGrid(values2, TEXT), { values: values2, start: { col: 0, row: 0 } }, cellValue, {}, {}, zero);
  assert.strictEqual(again.write.length + again.updated.length, 0, 'an empty cell already reads as zero');
});

test('locks are enforced even for a month no invoice touched this run', () => {
  const values = VALUES.map(r => r.slice());
  values[3][8] = 10;
  const { updated } = planAmountCells({}, locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}, {}, LOCKS);
  assert.strictEqual(updated.length, 1);
  assert.strictEqual(updated[0].value, 745.89);
});


/* ---------------- one file, several invoices, by page ---------------- *
 *
 * "Some time if there are 7 invoices for an app for a single month I merge
 * all 7 invoices in a single pdf; you are calculating a single page and
 * updating the first invoice amount". The reader yields the text page by
 * page; the file is split into its invoices and every one is added up.
 */

const PAGE = (n, of, ref, amount, extra) => `Page ${n} of ${of}  Invoice  Invoice number ${ref}  Date of issue July 1, 2026  ${extra || ''}  Total   $${amount}  Amount due   $${amount} USD`;

test('a merged file is the sum of every invoice in it', () => {
  const pages = [PAGE(1, 1, 'WYO2F9CO-0101', '241.33'), PAGE(1, 1, 'WYO2F9CO-0102', '225.79'), PAGE(1, 1, 'WYO2F9CO-0103', '222.53')];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.amount, 689.65);
  assert.strictEqual(r.usable, true);
  assert.match(r.via, /3 invoices in one file/);
  assert.deepStrictEqual(r.refs, ['WYO2F9CO0101', 'WYO2F9CO0102', 'WYO2F9CO0103']);
});

test('a two-page invoice inside a merged file stays one invoice', () => {
  // Anthropic's usage invoice runs to two pages; its total is on the second.
  const pages = [
    PAGE(1, 1, 'Q8MUNTUC-0200', '138.82'),
    'Page 1 of 2  Invoice  Invoice number Q8MUNTUC-0203  Date of issue August 1, 2026  Claude Opus 4.8 Usage Jul 2   Jul 31, 2026  $1,362.52',
    'Page 2 of 2  Claude Sonnet 5 Usage Jul 2   Jul 31, 2026  $5,863.72  Subtotal $12,686.51  Total   $13,479.42  Amount due   $13,479.42 USD',
  ];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.amount, 13618.24);
  assert.strictEqual(r.parts.length, 2);
});

test('without page numbers, an invoice header opens a new invoice and a total-less page is folded in', () => {
  const pages = [
    'Tax Invoice  Bill to Saras  Invoice ID 5001  Plan fee  Total $50.00',
    'Tax Invoice  Bill to Saras  Invoice ID 5002  Plan fee',
    'Line items continued  Total $70.00',
    'Tax Invoice  Bill to Saras  Invoice ID 5003  Total $30.00',
  ];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.amount, 150);
  assert.strictEqual(r.parts.length, 3);
});

test('an invoice and its receipt stapled together count once', () => {
  const pages = [PAGE(1, 1, 'A0589F17-0017', '85.00'), 'Page 1 of 1  Receipt  Receipt number 2601-5895  Invoice number A0589F17-0017  Amount paid $85.00', PAGE(1, 1, 'A0589F17-0018', '53.12')];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.amount, 138.12);
  assert.match(r.via, /1 receipt counted once/);
});

test('a merged file with one unreadable invoice is refused whole, not totalled short', () => {
  const pages = [PAGE(1, 1, 'X1', '10.00'), 'Page 1 of 1  Invoice  Invoice number X2  Date of issue July 1, 2026  (scanned, no figures)'];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.usable, false);
  assert.match(r.note, /one of the 2 invoices in this file could not be read/);
});

test('a single invoice is not mistaken for a merged file', () => {
  const pages = ['Page 1 of 2  Invoice number 77  Date of issue July 1, 2026  Item A $5', 'Page 2 of 2  Item B $5  Total $10.00  Amount due $10.00 USD'];
  const r = extractInvoiceTotal(pages.join('\n\n'), { pages });
  assert.strictEqual(r.amount, 10);
  assert.strictEqual(r.parts, undefined);
});


/* ---------------- a month of single and merged files together ---------------- *
 *
 * "They will be both: some have separate invoices and some have merged
 * invoices, look for both." Every invoice number in the month counts once,
 * whichever file it is seen in; a merged file contributes only the invoices
 * not already counted.
 */

const { tallyInvoices } = require('../lib/mail-sync');
const single = (name, ref, amount) => ({ name, entry: { usable: true, amount, currency: 'USD', ref } });
const merged = (name, parts) => ({ name, entry: { usable: true, amount: parts.reduce((s, p) => s + p.amount, 0), currency: 'USD', refs: parts.map(p => p.ref), parts } });

test('single and merged files in one month are added invoice by invoice', () => {
  const t = tallyInvoices([
    single('Invoice-0101.pdf', 'WYO2F9CO0101', 241.33),
    single('Invoice-0102.pdf', 'WYO2F9CO0102', 225.79),
    merged('rest of July.pdf', [{ ref: 'WYO2F9CO0103', amount: 222.53 }, { ref: 'WYO2F9CO0104', amount: 222.45 }, { ref: 'WYO2F9CO0105', amount: 221.43 }]),
  ]);
  assert.strictEqual(t.total, 1133.53);
  assert.strictEqual(t.counted, 3);
  assert.strictEqual(t.duplicates.length, 0);
});

test('a merged file that repeats an invoice already filed on its own counts the rest, not nothing', () => {
  const t = tallyInvoices([
    single('Invoice-0101.pdf', 'WYO2F9CO0101', 241.33),
    merged('July all.pdf', [{ ref: 'WYO2F9CO0101', amount: 241.33 }, { ref: 'WYO2F9CO0102', amount: 225.79 }]),
  ]);
  assert.strictEqual(t.total, 467.12, 'the repeated invoice once, the new one added');
  assert.strictEqual(t.duplicates.length, 1);
  assert.strictEqual(t.duplicates[0].ref, 'WYO2F9CO0101');
  // The other way round: the merged file first, then the single copy.
  const u = tallyInvoices([
    merged('July all.pdf', [{ ref: 'WYO2F9CO0101', amount: 241.33 }, { ref: 'WYO2F9CO0102', amount: 225.79 }]),
    single('Invoice-0101.pdf', 'WYO2F9CO0101', 241.33),
  ]);
  assert.strictEqual(u.total, 467.12);
});

test('a merged file holding only invoices already counted is a duplicate, and a receipt still pairs with its invoice', () => {
  const t = tallyInvoices([
    single('Invoice-A0589F17-0017.pdf', 'A0589F170017', 85),
    single('Receipt-2601-5895.pdf', 'A0589F170017', 85),
    merged('again.pdf', [{ ref: 'A0589F170017', amount: 85 }]),
    single('Invoice-0016.pdf', 'A0589F170016', 53.12),
    { name: 'scan.pdf', entry: { usable: false, amount: null, note: 'no text' } },
  ]);
  assert.strictEqual(t.total, 138.12);
  assert.strictEqual(t.counted, 2);
  assert.strictEqual(t.duplicates.length, 2);
  assert.strictEqual(t.unread.length, 1);
});

test('files with no invoice number are never treated as duplicates of each other', () => {
  const t = tallyInvoices([single('a.pdf', null, 32), single('b.pdf', null, 32), single('c.pdf', null, 32)]);
  assert.strictEqual(t.total, 96);
  assert.strictEqual(t.counted, 3);
});
