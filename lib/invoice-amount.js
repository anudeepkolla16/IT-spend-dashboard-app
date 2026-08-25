// Pulls the payable total and its currency out of an invoice PDF.
//
// pdf-parse is pinned to 1.x and required by its inner path. Version 2 wraps
// modern pdf.js, which needs browser globals (DOMMatrix, Path2D) and
// @napi-rs/canvas; on Vercel's Node runtime it throws ReferenceError at REQUIRE
// time, which killed the whole function before any handler could catch it and
// took invoice filing down with it. The inner path also dodges 1.1.1's
// debug-mode branch, which reads a bundled sample file that is not deployed.
//
// The require is lazy and guarded for the same reason: reading a total is a
// nice-to-have, and it must never be able to break the filing that works.
//
// The currency matters more than the number. The spend sheet is entirely in USD,
// but Indian vendors bill in INR — Tata Tele's June invoice reads
// "Net Payable (INR) 150591.60" where the sheet correctly holds 1574.40, the
// converted figure. Writing the face value would be a 95x error. So a total is
// only ever offered for the sheet when it is unambiguously USD; anything else is
// returned with its currency and left for a human.

// Ordered most- to least-specific. An invoice usually states several totals —
// Adobe prints "NET AMOUNT (USD) 34.97" (pre-tax) before
// "GRAND TOTAL (USD) 37.16" — so the payable total has to win regardless of
// which appears first in the text.
const TOTAL_PATTERNS = [
  { re: /\bnet\s+payable\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'net payable' },
  { re: /\bgrand\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'grand total' },
  { re: /\btotal\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total paid' },
  { re: /\bamount\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount paid' },
  { re: /\bamount\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount due' },
  { re: /\btotal\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total due' },
  { re: /\binvoice\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'invoice total' },
  { re: /\btotal\s*\(\s*[A-Z]{3}\s*\)[^0-9\-]{0,20}([\d.,]+)/i, label: 'total (CUR)' },
  { re: /\btotal\b[^0-9\-]{0,20}([\d.,]+)\s*$/im, label: 'total' },
];

const CURRENCY_WORDS = [
  { re: /\b(INR|₹|Rs\.?)\b/i, code: 'INR' },
  { re: /\b(USD|US\$)\b/i, code: 'USD' },
  { re: /\b(EUR|€)\b/i, code: 'EUR' },
  { re: /\b(GBP|£)\b/i, code: 'GBP' },
  { re: /\b(AUD|CAD|SGD|AED|JPY|CHF)\b/i, code: null }, // captured below
];

function parseNumber(raw) {
  const s = String(raw || '').trim().replace(/,/g, '');
  if (!/^\d+(\.\d+)?$/.test(s)) return null;
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

// Which currency the invoice is denominated in. Explicit statements
// ("Currency USD", "Price (INR)") beat a bare symbol, and a document that names
// more than one currency is reported as ambiguous rather than guessed at.
function detectCurrency(text) {
  const t = String(text || '');
  const declared = t.match(/\bcurrency\b\s*[:\-]?\s*([A-Z]{3})\b/i)
    || t.match(/\b([A-Z]{3})\s+currency\b/i)
    || t.match(/\bnet\s+payable\s*\(\s*([A-Z]{3})\s*\)/i)
    || t.match(/\bgrand\s+total\s*\(\s*([A-Z]{3})\s*\)/i)
    || t.match(/\btotal\s*\(\s*([A-Z]{3})\s*\)/i)
    || t.match(/\bamounts?\s+are\s+in\s+([a-z]{3})\b/i);
  if (declared) return { code: declared[1].toUpperCase(), via: 'declared' };

  const found = new Set();
  for (const { re, code } of CURRENCY_WORDS) {
    const m = t.match(re);
    if (m) found.add(code || m[1].toUpperCase());
  }
  if (t.includes('$') && !found.has('INR')) found.add('USD');
  if (found.size === 1) return { code: [...found][0], via: 'symbol' };
  if (found.size > 1) return { code: null, via: 'ambiguous', candidates: [...found] };
  return { code: null, via: 'unknown' };
}

// Returns { amount, currency, via, usable }. `usable` means: safe to offer as a
// figure for the USD spend sheet without a human deciding anything.
function extractInvoiceTotal(text) {
  const t = String(text || '').replace(/ /g, ' ');
  if (!t.trim()) return { amount: null, currency: null, via: 'empty', usable: false, note: 'No text could be read from the PDF (it may be a scan).' };

  const currency = detectCurrency(t);

  let amount = null;
  let via = null;
  for (const { re, label } of TOTAL_PATTERNS) {
    const m = t.match(re);
    if (!m) continue;
    const n = parseNumber(m[1]);
    if (n === null || n <= 0) continue;
    amount = n;
    via = label;
    break;
  }

  if (amount === null) {
    return { amount: null, currency: currency.code, via: 'no-total', usable: false, note: 'No payable total found in the invoice text.' };
  }
  if (currency.code !== 'USD') {
    return {
      amount, currency: currency.code, via, usable: false,
      note: currency.code
        ? `Invoice is in ${currency.code}; the sheet is in USD, so this needs converting before it can be used.`
        : `Could not tell which currency this invoice is in${currency.candidates ? ` (saw ${currency.candidates.join(', ')})` : ''}.`,
    };
  }
  return { amount, currency: 'USD', via, usable: true, note: '' };
}

// Returns { text, error }. Never throws — a PDF that cannot be read yields no
// amount, and the invoice is still filed.
let pdfParse;
let pdfParseError;
async function readPdfText(buffer) {
  if (pdfParse === undefined && pdfParseError === undefined) {
    try {
      pdfParse = require('pdf-parse/lib/pdf-parse.js');
    } catch (e) {
      pdfParseError = e;
      pdfParse = null;
    }
  }
  if (!pdfParse) return { text: '', error: `PDF reader unavailable (${pdfParseError && pdfParseError.message}).` };
  try {
    const parsed = await pdfParse(buffer);
    return { text: (parsed && parsed.text) || '', error: null };
  } catch (e) {
    return { text: '', error: `Could not read the PDF (${e.message}).` };
  }
}

module.exports = { extractInvoiceTotal, detectCurrency, readPdfText, TOTAL_PATTERNS };
