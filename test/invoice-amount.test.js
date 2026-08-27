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

test('leaves a filled cell alone when the invoice total is lower', () => {
  // A new invoice can only add, so a shortfall means invoices are missing from
  // the folder — not that less was spent. Overwriting here would replace a
  // correct figure with a wrong one.
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 500 } }, grid(), USED, cellValue, {}
  );
  assert.strictEqual(write.length + updated.length, 0);
  assert.strictEqual(skippedFilled.length, 1);
  assert.strictEqual(skippedFilled[0].current, 751.96);
  assert.strictEqual(skippedFilled[0].reason, 'invoice-total-lower');
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

test('never touches a figure a human or the statement put there', () => {
  // The cell holds 524.27 but we never wrote it — it is not ours to change.
  const values = VALUES.map(r => r.slice());
  values[3][8] = 524.27;
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 492.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, {}
  );
  assert.strictEqual(write.length, 0);
  assert.strictEqual(updated.length, 0);
  assert.strictEqual(skippedFilled.length, 1, 'must be reported, not written');
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

test('a hand correction survives an invoice total that is lower', () => {
  const prior = { 'Bubble Starter||2026-06': 492.27 };
  const values = VALUES.map(r => r.slice());
  values[3][8] = 600; // corrected upward by a human
  const { write, updated, skippedFilled } = planAmountCells(
    { 'Bubble Starter': { '2026-06': 524.27 } },
    locateGrid(values, TEXT), { values, start: { col: 0, row: 0 } }, cellValue, prior
  );
  assert.strictEqual(write.length + updated.length, 0, 'must not drag a correction back down');
  assert.strictEqual(skippedFilled.length, 1);
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
