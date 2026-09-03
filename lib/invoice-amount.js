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
  // Google's invoices and Ads statements: "Total in USD $13.87". The Ads
  // statement is "not a bill" but its Total in USD is the month's spend, which
  // is what the sheet has always recorded for it.
  { re: /\btotal\s+in\s+[A-Z]{3}\b[^0-9\-]{0,20}([\d.,]+)/i, label: 'total in CUR' },
  { re: /\bgrand\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'grand total' },
  { re: new RegExp(String.raw`\btotal\s+incl\.?(?:uding)?\s*(?:of\s+)?(?:sales\s+tax|tax|vat|gst)\b[^0-9\-]{0,40}([\d.,]+)`, 'i'), label: 'total incl. tax' },
  { re: /\btotal\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total paid' },
  { re: /\bamount\s+paid\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount paid' },
  { re: /\bamount\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'amount due' },
  { re: /\btotal\s+due\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total due' },
  { re: /\binvoice\s+total\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'invoice total' },
  { re: /\btotal\s*\(\s*[A-Z]{3}\s*\)[^0-9\-]{0,20}([\d.,]+)/i, label: 'total (CUR)' },
  // "Total $82.31 USD" (Sentry), "Total USD 816.00" (Webflow): the figure is
  // followed by its currency rather than ending the line.
  { re: new RegExp(String.raw`\btotal\b(?![^0-9\-]{0,20}?\b${TAX_WORDS}\b)[^0-9\-]{0,20}([\d.,]+)\s*(?:USD|INR|EUR|GBP)\b`, 'i'), label: 'total (trailing currency)' },
  { re: new RegExp(String.raw`\btotal\b(?![^0-9\-]{0,20}?\b${TAX_WORDS}\b)[^0-9\-]{0,20}([\d.,]+)\s*$`, 'im'), label: 'total' },
  // Last resort: an invoice that only ever states a pre-tax total. Better the
  // figure it does print than none — but never the tax line, which is not a
  // total of anything payable.
  { re: /\btotal\s+excl\.?(?:uding)?\s*(?:of\s+)?(?:sales\s+)?(?:tax|vat|gst)\b[^0-9\-]{0,40}([\d.,]+)/i, label: 'total excl. tax' },
];

// Currency codes are matched CASE-SENSITIVELY. With the `i` flag "Rs" fired
// on the word "rs", and the declared-currency patterns read "any currency" as
// the code ANY and "currency and" as AND — two of the archive's invoices were
// reported as being in a currency that does not exist.
const CURRENCY_WORDS = [
  { re: /\bINR\b|₹|\bRs\.?\s*\d/, code: 'INR' },
  { re: /\bUSD\b|US\$/, code: 'USD' },
  { re: /\bEUR\b|€/, code: 'EUR' },
  { re: /\bGBP\b|£/, code: 'GBP' },
  { re: /\b(AUD|CAD|SGD|AED|JPY|CHF)\b/, code: null }, // captured below
];
const KNOWN_CODES = new Set(['INR', 'USD', 'EUR', 'GBP', 'AUD', 'CAD', 'SGD', 'AED', 'JPY', 'CHF']);

function parseNumber(raw) {
  const s = String(raw || '').trim().replace(/,/g, '');
  if (!/^\d+(\.\d+)?$/.test(s)) return null;
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

// Which currency the invoice is denominated in. Explicit statements
// ("Currency USD", "Price (INR)") beat a bare symbol, and a document that names
// more than one currency is reported as ambiguous rather than guessed at.
function detectCurrency(text, preferred) {
  const t = String(text || '');
  const declared = t.match(/\b[Cc]urrency\b\s*[:\-]?\s*([A-Z]{3})\b/)
    || t.match(/\b([A-Z]{3})\s+[Cc]urrency\b/)
    || t.match(/\b[Nn]et\s+[Pp]ayable\s*\(\s*([A-Z]{3})\s*\)/)
    || t.match(/\b[Gg]rand\s+[Tt]otal\s*\(\s*([A-Z]{3})\s*\)/)
    || t.match(/\b[Tt]otal\s*\(\s*([A-Z]{3})\s*\)/)
    || t.match(/\b[Tt]otal\s+in\s+([A-Z]{3})\b/)
    || t.match(/\b[Aa]mounts?\s+are\s+in\s+([A-Za-z]{3})\b/);
  if (declared && KNOWN_CODES.has(declared[1].toUpperCase())) return { code: declared[1].toUpperCase(), via: 'declared' };

  const found = new Set();
  for (const { re, code } of CURRENCY_WORDS) {
    const m = t.match(re);
    if (m) found.add(code || m[1].toUpperCase());
  }
  if (t.includes('$') && !found.has('INR')) found.add('USD');
  if (found.size === 1) return { code: [...found][0], via: 'symbol' };
  if (found.size > 1) {
    // A vendor whose invoice always states two currencies — PhantomBuster
    // prints EUR beside USD — can say in its rule which one the sheet takes.
    const want = String(preferred || '').toUpperCase();
    if (want && found.has(want)) return { code: want, via: 'rule', candidates: [...found] };
    return { code: null, via: 'ambiguous', candidates: [...found] };
  }
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
const INVOICE_REF = /invoice\s*(?:number|no\.?|num\.?|#)\s*[:#-]?\s*([A-Za-z0-9][A-Za-z0-9 \-\/\n]{0,40})/i;
const REF_STOPWORDS = /\b(receipt|date|dated|bill|billed|amount|page|due|paid|order|customer|account|subtotal|total)\b/i;

// A Stripe invoice number is an eight-character account prefix and a sequence
// ("Q8MUNTUC-0180"), and the two are often split by a line break in the text.
// The prefix alone is NOT an invoice number: every invoice on that account
// carries it, and treating it as one made five Cursor invoices in a month
// read as one invoice and four duplicates of it. So the sequence is taken
// across the break when it is there, and a bare prefix yields nothing.
const STRIPE_PREFIX = /^[A-Z0-9]{8}$/;

function extractInvoiceRef(text) {
  const m = String(text || '').match(INVOICE_REF);
  if (!m) return null;
  // Stop at whatever field comes next: the capture is deliberately greedy so a
  // number split across runs of spaces ("A0589F17   0017") survives intact.
  const cut = m[1].split(REF_STOPWORDS)[0];
  const ref = cut.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  // Two or three characters is a page number or a stray digit, not a reference.
  if (ref.length < 4) return null;
  // Only the prefix was captured but the sequence is on the page (past the
  // capture window, or after a word the cut stops at): put the two together.
  if (STRIPE_PREFIX.test(ref) && /[A-Z]/.test(ref)) {
    const seq = String(text).match(new RegExp(`\\b${ref}[-\\s]*(\\d{4,})\\b`, 'i'));
    if (seq) return ref + seq[1];
  }
  return ref;
}

// Every distinct invoice number in the text, for a file that holds more than
// one invoice — see compoundTotal.
function extractInvoiceRefs(text) {
  const out = [];
  const re = new RegExp(INVOICE_REF.source, 'gi');
  let m;
  while ((m = re.exec(String(text || ''))) !== null) {
    const ref = extractInvoiceRef(m[0]);
    if (ref && !out.includes(ref)) out.push(ref);
  }
  return out;
}

// A single figure the sheet could never hold. A ten-digit invoice number has
// been read as an amount before ("Invoice number: 5685830900" became a
// 5,685,830,900 charge for Google Ads), so anything past this is refused.
const MAX_PLAUSIBLE = 5000000;

// One PDF, several invoices. The owner staples a month's invoices into one
// file — seven Cursor invoices as one PDF — and Anthropic's "July 26.pdf" is
// four. Read as one document, such a file yields the FIRST total (138.82 of
// 13,895.83), so it is split into its invoices and each is totalled.
//
// The split is by page. "Page 1 of N" opens an invoice and the N pages are
// its own; without page numbers, a page carrying an invoice header (an
// invoice or receipt number, a date of issue, a "Bill to") opens one. A group
// that reads no total is folded into the next when they share an invoice
// number or one of them has none — a running header on page 2 must not
// split an invoice in two. Without page texts (pdf-parse) the file is split
// on its invoice-number lines instead.
//
// An invoice and its receipt stapled together carry the same number and count
// once. Any invoice in the file that will not read fails the whole file,
// which is then held: a merged file totalled short would look complete.
const PAGE_OF = /\bpage\s+(\d+)\s+of\s+(\d+)\b/i;
const INVOICE_HEADER = /\b(?:invoice\s*(?:number|no\.?|num\.?|#|id)|receipt\s*(?:number|no\.?|#)|date\s+of\s+issue|invoice\s+date|tax\s+invoice|bill\s+to|billed\s+to)\b/i;

function groupPages(pages) {
  const groups = [];
  let cur = null;
  for (const page of pages) {
    const m = String(page).match(PAGE_OF);
    const opens = m ? Number(m[1]) === 1 : (INVOICE_HEADER.test(page) && (!cur || cur.total !== null));
    if (opens || !cur) { cur = { text: '', total: null }; groups.push(cur); }
    cur.text += (cur.text ? '\n\n' : '') + page;
    cur.total = null;
    const r = extractInvoiceTotal(cur.text, { single: true });
    cur.total = r.amount !== null && r.usable ? r.amount : null;
    cur.ref = extractInvoiceRef(cur.text);
  }
  // A headerless continuation or a running header: fold a group with no
  // total into its neighbour when the numbers agree or one is missing.
  for (let k = 0; k < groups.length - 1;) {
    const g = groups[k], n = groups[k + 1];
    if (g.total === null && (!g.ref || !n.ref || g.ref === n.ref)) {
      n.text = g.text + '\n\n' + n.text;
      const r = extractInvoiceTotal(n.text, { single: true });
      n.total = r.amount !== null && r.usable ? r.amount : null;
      n.ref = extractInvoiceRef(n.text);
      groups.splice(k, 1);
    } else k++;
  }
  return groups;
}

function compoundTotal(t, opts) {
  const pages = opts && Array.isArray(opts.pages) && opts.pages.length > 1 ? opts.pages : null;
  let sections;
  if (pages) {
    const groups = groupPages(pages);
    if (groups.length < 2) return null;
    sections = groups.map(g => g.text);
  } else {
    const refs = extractInvoiceRefs(t);
    if (refs.length < 2) return null;
    sections = t.split(/(?=\binvoice\s*(?:number|no\.?|num\.?|#)\b)/i).filter(x => /\binvoice\s*(?:number|no\.?|num\.?|#)/i.test(x));
    if (sections.length < 2) return null;
  }
  const parts = [];
  const seen = new Set();
  let duplicates = 0;
  for (const section of sections) {
    const r = extractInvoiceTotal(section, { ...(opts || {}), single: true, pages: undefined });
    if (r.amount === null || !r.usable) return { failed: true, note: `one of the ${sections.length} invoices in this file could not be read: ${r.note}` };
    const ref = extractInvoiceRef(section);
    if (ref && seen.has(ref)) { duplicates++; continue; } // its own receipt
    if (ref) seen.add(ref);
    parts.push({ ref, amount: r.amount });
  }
  if (parts.length < 2 && !duplicates) return null;
  const total = Math.round(parts.reduce((a, b) => a + b.amount, 0) * 100) / 100;
  return { amount: total, parts, refs: parts.map(p => p.ref).filter(Boolean), duplicates };
}

// Google's PDFs lay the summary out as a column of figures beside a column of
// labels, and the text comes out as all the figures then all the labels:
//   "$10.00 $0.73 $0.19 $1.45 $1.50 $13.87 Subtotal in USD State sales tax …
//    Federal Universal Service Fund State 911 Tax Total in USD"
// Six figures, six labels, in the same order — so the last figure is the
// Total. Taken only when the run of figures is immediately followed by a label
// block that starts with a subtotal and ends with the total, which is what
// makes the pairing safe to rely on.
const FLAT_COLUMNS = /((?:-?\$\s?[\d,]+\.\d{2}\s+){2,10})(Subtotal\s+in\s+[A-Z]{3}\b(?:(?!\$)[\s\S]){0,240}?\bTotal\s+in\s+[A-Z]{3}\b)/i;

function flattenedTotal(t) {
  const m = t.match(FLAT_COLUMNS);
  if (!m) return null;
  const figures = m[1].trim().split(/\s+(?=-?\$)/).map(f => f.replace(/[$\s]/g, ''));
  // Labels are separated by two or more spaces, or by a closing bracket or a
  // currency code followed by a capital — "Tax (0%)  Total in USD".
  const labels = m[2].split(/\s{2,}|(?<=\)|\bUSD|\bINR|\bEUR|\bGBP)\s+(?=[A-Z])/).map(l => l.trim()).filter(Boolean);
  // The pairing only holds when there are as many figures as labels.
  if (figures.length !== labels.length) return null;
  return figures[figures.length - 1];
}

function extractInvoiceTotal(text, options) {
  const opts = options || {};
  // pdf-parse runs a label straight into its figure ("TotalUSD 816.00",
  // "Total$82.31") when the two sit in different table cells; give the
  // patterns the word boundary they expect.
  const t = String(text || '').replace(/\u00a0/g, ' ').replace(/([a-z])(USD|INR|EUR|GBP|\$)/g, '$1 $2');
  if (!t.trim()) return { amount: null, currency: null, via: 'empty', usable: false, note: 'No text could be read from the PDF (it may be a scan).' };

  const currency = detectCurrency(t, opts.currency);

  if (!opts.single) {
    const compound = compoundTotal(t, opts);
    if (compound && compound.failed) return { amount: null, currency: currency.code, via: 'compound', usable: false, note: compound.note };
    if (compound) {
      return {
        amount: compound.amount, currency: 'USD', via: `${compound.parts.length} invoices in one file${compound.duplicates ? ` (${compound.duplicates} receipt${compound.duplicates === 1 ? '' : 's'} counted once)` : ''}`, usable: true, note: '',
        refs: compound.refs, parts: compound.parts,
      };
    }
  }

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

  if (amount === null) {
    const flat = flattenedTotal(t);
    const n = flat === null ? null : parseNumber(flat.replace(/^-/, ''));
    if (n !== null && n > 0) { amount = n; via = 'total in CUR (columns)'; }
  }

  let zero = false;
  if (amount === null) for (const { re, label } of TOTAL_PATTERNS) {
    const m = t.match(re);
    if (!m) continue;
    const n = parseNumber(m[1]);
    if (n === null) continue;
    // A nought is a real total when the document says so plainly — PostHog's
    // "$0.00 USD due" invoices are on file every month and cost nothing. It is
    // remembered but the search goes on, so a later, non-zero total wins.
    if (n === 0) { zero = true; continue; }
    amount = n;
    via = label;
    break;
  }
  if (amount === null && zero && /\$\s*0\.00\s*(?:USD\s*)?due\b|amount\s+due\s*\$?\s*0\.00\b/i.test(t)) {
    amount = 0;
    via = 'amount due 0.00';
  }

  if (amount === null) {
    return { amount: null, currency: currency.code, via: 'no-total', usable: false, note: 'No payable total found in the invoice text.' };
  }
  if (amount > MAX_PLAUSIBLE) {
    return { amount, currency: currency.code, via, usable: false, note: `${amount.toLocaleString('en-US')} is not a plausible charge — most likely an invoice or account number read as the total.` };
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

// Returns { text, error, reader }. Never throws — a PDF that cannot be read
// yields no amount, and the invoice is still held or filed.
//
// Two readers, because neither reads everything:
//   · pdfjs-dist 3.x (the legacy build, which runs on plain Node with no canvas
//     and no browser globals). It rebuilds a broken cross-reference table, which
//     pdf-parse 1.x — a 2018 pdf.js — refuses with "bad XRef entry"; twenty-one
//     of the archive's invoices failed that way while reading fine elsewhere.
//     Its text is put back into lines by position, so a label and its figure
//     in two table cells come out on one line with a space between them.
//   · pdf-parse 1.x, kept for anything pdfjs declines.
// Both are required lazily and guarded: reading a total is never allowed to
// break the filing that works. pdf-parse v2 is avoided on purpose — it needs
// browser globals at require time and took the whole function down on Vercel.
let pdfjs;
let pdfjsError;
let pdfParse;
let pdfParseError;

// The worker is required HERE, statically, and its path handed to pdfjs.
// pdfjs itself loads the worker through `eval("require")(workerSrc)`, which no
// bundler can follow: on Vercel the function shipped without pdf.worker.js
// and every read failed with "Setting up fake worker failed" (the very files
// this reader exists for, while pdf-parse failed on them too). A plain
// require is traced and bundled; vercel.json lists the file as well.
function loadPdfjs() {
  if (pdfjs === undefined && pdfjsError === undefined) {
    try {
      pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
      require('pdfjs-dist/legacy/build/pdf.worker.js');
      pdfjs.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');
    } catch (e) {
      pdfjsError = e;
      pdfjs = null;
    }
  }
  return pdfjs;
}

function loadPdfParse() {
  if (pdfParse === undefined && pdfParseError === undefined) {
    try {
      pdfParse = require('pdf-parse/lib/pdf-parse.js');
    } catch (e) {
      pdfParseError = e;
      pdfParse = null;
    }
  }
  return pdfParse;
}

// Items that share a baseline (within two points) are one line, left to right.
// Between two items, a gap wider than a third of the font size is a space and
// a wide one a column break; a narrower gap is within a word, so a word split
// across two items does not grow a space. "Total" and "USD 816.00" from two
// table cells therefore read as "Total   USD 816.00".
function linesFromTextContent(content) {
  const rows = [];
  for (const it of content.items || []) {
    if (!it.str) continue;
    const y = it.transform[5];
    const x = it.transform[4];
    const size = Math.abs(it.transform[0]) || Math.abs(it.transform[3]) || 10;
    let row = rows.find(r => Math.abs(r.y - y) <= 2);
    if (!row) { row = { y, items: [] }; rows.push(row); }
    row.items.push({ x, end: x + (it.width || 0), size, str: it.str });
  }
  rows.sort((a, b) => b.y - a.y);
  return rows.map(r => {
    r.items.sort((a, b) => a.x - b.x);
    let line = '';
    let prevEnd = null;
    for (const it of r.items) {
      if (prevEnd !== null) {
        const gap = it.x - prevEnd;
        if (gap > it.size * 1.5) line += '   ';
        else if (gap > it.size * 0.15) line += ' ';
      }
      line += it.str;
      prevEnd = it.end;
    }
    return line;
  });
}

async function readWithPdfjs(buffer) {
  const lib = loadPdfjs();
  if (!lib) throw new Error(`pdfjs unavailable (${pdfjsError && pdfjsError.message})`);
  const task = lib.getDocument({
    data: new Uint8Array(buffer), disableFontFace: true, isEvalSupported: false, verbosity: 0,
  });
  const doc = await task.promise;
  try {
    const pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const content = await page.getTextContent();
      pages.push(linesFromTextContent(content).join('\n'));
    }
    return { text: pages.join('\n\n'), pages };
  } finally {
    try { await doc.destroy(); } catch (_) { /* nothing to do */ }
  }
}

async function readWithPdfParse(buffer) {
  const lib = loadPdfParse();
  if (!lib) throw new Error(`pdf-parse unavailable (${pdfParseError && pdfParseError.message})`);
  const parsed = await lib(buffer);
  return { text: (parsed && parsed.text) || '', pages: null };
}

async function readPdfText(buffer) {
  const errors = [];
  for (const [reader, fn] of [['pdfjs', readWithPdfjs], ['pdf-parse', readWithPdfParse]]) {
    try {
      const { text, pages } = await fn(buffer);
      // `pages` is the text page by page (pdfjs only) — what tells a file
      // holding several invoices apart, see compoundTotal.
      if (String(text).trim()) return { text, pages, error: null, reader };
      errors.push(`${reader}: no text (a scanned image?)`);
    } catch (e) {
      errors.push(`${reader}: ${e.message}`);
    }
  }
  return { text: '', error: `Could not read the PDF (${errors.join('; ')}).`, reader: null };
}

module.exports = { extractInvoiceTotal, extractInvoiceRef, extractInvoiceRefs, detectCurrency, readPdfText, linesFromTextContent, TOTAL_PATTERNS, MAX_PLAUSIBLE };
