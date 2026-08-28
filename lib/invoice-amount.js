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
// An invoice settled from a prepaid balance still prints a Total, but nothing
// was actually charged. Anthropic's July Claude receipt reads "$0.00 paid" with
// "Applied balance $4,037.39" and a "Total $4,037.39" — and the sheet correctly
// records 0.00. Taking the Total there would be wrong by the whole amount, so
// what was actually paid wins whenever the document says so.
const PAID_PATTERNS = [
  { re: /\$\s*([\d.,]+)\s*(?:USD\s*)?paid\s+on\b/i, label: 'paid on' },
  { re: /\bamount\s+paid\b[^0-9\-\n]{0,40}\$?\s*([\d.,]+)/i, label: 'amount paid' },
  { re: /\btotal\s+paid\b[^0-9\-\n]{0,40}\$?\s*([\d.,]+)/i, label: 'total paid' },
];

// A tax word between "total" and the figure changes what the figure IS. Luzmo's
// invoice prints three of them in a column:
//   Total excl. Sales Tax:      524.50   <- pre-tax, not what is owed
//   Total Sales Tax:             32.78   <- the tax alone, a 17x error if taken
//   TOTAL INCL. SALES TAX:      557.28   <- the payable total
// A bare /total/ pattern takes whichever of those it reaches first, so the
// tax-inclusive line is matched explicitly and ahead of it, and the bare pattern
// refuses any line whose label says tax at all.
const TAX_WORDS = String.raw`(?:sales\s+tax|tax|vat|gst|excl\.?(?:uding)?)`;

const TOTAL_PATTERNS = [
  { re: /\bnet\s+payable\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'net payable' },
  { re: /\bgrand\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'grand total' },
  { re: new RegExp(String.raw`\btotal\s+incl\.?(?:uding)?\s*(?:of\s+)?(?:sales\s+tax|tax|vat|gst)\b[^0-9\-]{0,40}([\d.,]+)`, 'i'), label: 'total incl. tax' },
  { re: /\btotal\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total paid' },
  { re: /\bamount\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount paid' },
  { re: /\bamount\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount due' },
  { re: /\btotal\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total due' },
  { re: /\binvoice\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'invoice total' },
  { re: /\btotal\s*\(\s*[A-Z]{3}\s*\)[^0-9\-]{0,20}([\d.,]+)/i, label: 'total (CUR)' },
  { re: new RegExp(String.raw`\btotal\b(?![^0-9\-]{0,20}?\b${TAX_WORDS}\b)[^0-9\-]{0,20}([\d.,]+)\s*$`, 'im'), label: 'total' },
  // Last resort: an invoice that only ever states a pre-tax total. Better the
  // figure it does print than none — but never the tax line, which is not a
  // total of anything payable.
  { re: /\btotal\s+excl\.?(?:uding)?\s*(?:of\s+)?(?:sales\s+)?(?:tax|vat|gst)\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total excl. tax' },
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
// An invoice and its payment receipt are one charge, sent as two documents, and
// a folder holding both would be totalled twice. Apollo's August folder is
// exactly that: Invoice-A0589F17-0017.pdf and Receipt-2601-5895.pdf, both
// $85.00, and the receipt is the payment for that invoice. Only one PDF parsed,
// so the folder happened to total 85.00 — the moment both are readable it
// totals 170.00.
//
// The two are told apart by the INVOICE number, which both documents carry
// ("Invoice number A0589F17 0017" appears on the receipt too). The receipt's own
// "Receipt number" is a different number and must never be the key, so the label
// is matched specifically and the capture stops at the next field.
const INVOICE_REF = /invoice\s*(?:number|no\.?|num\.?|#)\s*[:#-]?\s*([A-Za-z0-9][A-Za-z0-9 \-\/]{0,40})/i;
const REF_STOPWORDS = /\b(receipt|date|dated|bill|billed|amount|page|due|paid|order|customer|account|subtotal|total)\b/i;

function extractInvoiceRef(text) {
  const m = String(text || '').match(INVOICE_REF);
  if (!m) return null;
  // Stop at whatever field comes next: the capture is deliberately greedy so a
  // number split across runs of spaces ("A0589F17   0017") survives intact.
  const cut = m[1].split(REF_STOPWORDS)[0];
  const ref = cut.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  // Two or three characters is a page number or a stray digit, not a reference.
  return ref.length >= 4 ? ref : null;
}

function extractInvoiceTotal(text) {
  const t = String(text || '').replace(/ /g, ' ');
  if (!t.trim()) return { amount: null, currency: null, via: 'empty', usable: false, note: 'No text could be read from the PDF (it may be a scan).' };

  const currency = detectCurrency(t);

  let amount = null;
  let via = null;

  // What was actually paid beats what was billed, but only when the document
  // shows a balance was applied — otherwise "paid" and "total" agree anyway and
  // an ordinary receipt should not be read as zero on a stray phrase.
  const balanceApplied = /\bapplied\s+balance\b/i.test(t) || /\bcredit\s+applied\b/i.test(t);
  if (balanceApplied) {
    for (const { re, label } of PAID_PATTERNS) {
      const m = t.match(re);
      if (!m) continue;
      const n = parseNumber(m[1]);
      if (n === null) continue;
      amount = n;           // may legitimately be 0 — nothing was charged
      via = `${label} (balance applied)`;
      break;
    }
  }

  if (amount === null) for (const { re, label } of TOTAL_PATTERNS) {
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

module.exports = { extractInvoiceTotal, extractInvoiceRef, detectCurrency, readPdfText, TOTAL_PATTERNS };
