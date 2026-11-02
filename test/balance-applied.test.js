// Run with: node --test
//
// Anthropic's September Claude invoice went into the sheet as $5,026.36 for a
// month in which Saras Analytics was charged nothing. The document says so
// plainly:
//
//   Extra usage units, Enterprise plan Sep 1 - Sep 30, 2026   $4,730.69
//   Subtotal                                                  $4,730.69
//   Tax - Massachusetts  6.25% on $4,730.69                     $295.67
//   Total                                                     $5,026.36
//   Applied balance                                           $5,026.36
//   Amount due                                                    $0.00 USD
//
// The reader already preferred what was actually paid over what was billed
// whenever a balance had been applied — but it looked only for the word
// "paid". That is how a RECEIPT puts it. An INVOICE states what is still owed,
// and here that is nothing. Finding no "paid", the branch fell through to the
// Total and the whole 5,026.36 was offered to the sheet as usable USD.

const test = require('node:test');
const assert = require('node:assert');

const { extractInvoiceTotal } = require('../lib/invoice-amount');

// The live file: Invoices/Claude Ai/sep/Invoice-2FSKIDHO-0090.pdf
const SEPTEMBER = `Page 1 of 1  Invoice  Invoice number   2FSKIDHO   0090  Date of issue   October 1, 2026 Date due   October 31, 2026 Organization ID   674ecb7a-83fc-4a2d-af60    3b1b90f5e945  Anthropic, PBC   @anthropic 548 Market Street PMB 90375 San Francisco, California 94104 United States support@anthropic.com  Bill to  Saras Analytics 92 Ruggles Street Westborough, Massachusetts 01581 United States krishna@sarasanalytics.com  $0.00 USD due October 31, 2026  Description   Qty   Unit price   Tax   Amount  Extra usage units, Enterprise plan Sep 1   Sep 30, 2026 1   $4,730.69   6.25%   $4,730.69  Subtotal   $4,730.69 Total excluding tax   $4,730.69 Tax - Massachusetts     6.25% on $4,730.69     $295.67 Total   $5,026.36 Applied balance     $5,026.36  Amount due   $0.00 USD`;

test('an invoice the applied balance covered costs nothing', () => {
  const r = extractInvoiceTotal(SEPTEMBER);
  assert.strictEqual(r.amount, 0, 'the balance covered it; nothing was charged in September');
  assert.strictEqual(r.usable, true, 'nought is a figure the sheet can hold');
  assert.match(r.via, /balance applied/);
  assert.notStrictEqual(r.amount, 5026.36, 'the Total is what was consumed, not what was paid');
});

test('a balance that covers only part of the bill leaves the rest owing', () => {
  // The rule has to generalise: what is still owed IS the month's charge,
  // whether that is all of the total, none of it, or something between.
  const part = SEPTEMBER
    .replace('Applied balance     $5,026.36', 'Applied balance     $3,000.00')
    .replace('Amount due   $0.00 USD', 'Amount due   $2,026.36 USD');
  const r = extractInvoiceTotal(part);
  assert.strictEqual(r.amount, 2026.36);
  assert.strictEqual(r.usable, true);
});

test('an ordinary invoice is still read as its total, not its balance', () => {
  // The due patterns are consulted ONLY once a document says a balance was
  // applied. Without that line the same invoice is worth its full total.
  const plain = SEPTEMBER
    .replace('Applied balance     $5,026.36', '')
    .replace('Amount due   $0.00 USD', 'Amount due   $5,026.36 USD')
    .replace('$0.00 USD due October 31, 2026', '$5,026.36 USD due October 31, 2026');
  const r = extractInvoiceTotal(plain);
  assert.strictEqual(r.amount, 5026.36);
  assert.strictEqual(r.usable, true);
  assert.doesNotMatch(r.via, /balance applied/);
});

test('a receipt that says "paid" is still read that way', () => {
  // PAID_PATTERNS come first, so the wording a receipt uses keeps its label.
  const receipt = `Anthropic, PBC Receipt  Total   $4,037.39  Applied balance   $4,037.39  $0.00 USD paid on July 1, 2026`;
  const r = extractInvoiceTotal(receipt);
  assert.strictEqual(r.amount, 0);
  assert.match(r.via, /paid on \(balance applied\)/);
});

// --- the guard that cost every AWS invoice its figure ------------------------

test('a due date is never read as the amount owed, balance or no balance', () => {
  // "DUE ON October 11" read as $11 across the whole archive. The due patterns
  // added here are phrased in exactly those words, so they carry the same
  // guards — a regression would be silent and archive-wide.
  const aws = `Amazon Web Services Invoice  Applied balance   $0.00  TOTAL AMOUNT DUE ON October 11 , 2026   USD 2,401.67  Charges $2,387.70  Tax $13.97  Total $2,401.67`;
  const r = extractInvoiceTotal(aws);
  assert.notStrictEqual(r.amount, 11, 'the day of the month is not the charge');
  assert.strictEqual(r.amount, 2401.67);
});

test('the day can sit before its month too', () => {
  const t = `Invoice  Credit applied   $5.00  Amount due by 15 March 2027   USD 840.00  Total   USD 840.00`;
  const r = extractInvoiceTotal(t);
  assert.notStrictEqual(r.amount, 15);
  assert.strictEqual(r.amount, 840);
});

// --- and the cache has to be told ------------------------------------------

test('every invoice on record is read again under the new rule', () => {
  // The figures this changes are already in the sheet. Leaving PARSE_VERSION
  // alone would fix new invoices and leave Anthropic's September at 5,026.36
  // for good.
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  const version = Number(src.match(/const PARSE_VERSION = (\d+);/)[1]);
  assert.ok(version >= 5, `PARSE_VERSION is ${version}; a balance-applied invoice cached under 4 keeps its Total`);
});
