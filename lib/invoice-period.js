// Which month an invoice actually belongs to.
//
// The month used to come only from when the mail arrived, which is wrong for any
// vendor whose cycle straddles two months. Luzmo's August invoice reads
// "Monthly Elite license fee (period from 2026-08-26 until 2026-09-26)": six
// days of that fall in August and twenty-five in September, and the charge itself
// falls due on Sep 09 — so it belongs in Sep-26, not the Aug-26 the mail date
// gave it.
//
// The rule is: file by the month holding most of the billing period. A period
// that sits inside one calendar month (Jul 1 - Jul 31) resolves to that month,
// exactly as before; only a straddling cycle moves.
//
// Only an EXPLICITLY LABELLED period moves an invoice — "period", "billing
// period", "service term" and the like. Two things on an ordinary invoice look
// just like a date range and are not one:
//   · the invoice date beside the due date ("Date: Aug 26, 2026 ... Due date:
//     Sep 09, 2026"), which would read as a period that is nothing of the sort;
//   · a line item's own dates ("Starter Web Plan 8/12/26 - 9/12/26"), where a
//     vendor like Bubble bills a dozen applications on separate cycles inside
//     one month. Those totals were reconciled against the month they arrived in,
//     and are left there.

const MONTH_WORDS = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
};

const DAY_MS = 24 * 60 * 60 * 1000;
const pad = n => String(n).padStart(2, '0');
const ym = d => `${d.year}-${pad(d.month)}`;

// Anything that reads as a single date. Ordered most- to least-specific, since
// the alternation takes the first that fits.
const DATE = [
  String.raw`\d{4}[-/.]\d{1,2}[-/.]\d{1,2}`,                                  // 2026-08-26
  String.raw`\d{1,2}[-/.\s][A-Za-z]{3,9}[-/.\s]\d{2,4}`,                      // 01-JUN-2026, 1 Aug 2026
  String.raw`\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}`,                                // 8/12/26, 26-08-2026
  String.raw`[A-Za-z]{3,9}\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s*\d{4}`,           // Aug 26, 2026
  String.raw`\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]{3,9}`,                        // 26 August
  String.raw`[A-Za-z]{3,9}\.?\s+\d{1,2}(?:st|nd|rd|th)?`,                     // Aug 26
].join('|');

// A labelled period, and only a labelled one. The label may lead ("Billing
// period: X to Y") or the word "period" may stand alone ("period from X until Y").
const PERIOD_RE = new RegExp(
  String.raw`\b(?:billing|service|subscription|licen[cs]e|coverage|usage|invoice|contract|plan)?\s*`
  + String.raw`(?:period|term)s?\b\s*[:\-–—]?\s*(?:is\s+|of\s+|for\s+|covering\s+)?(?:from\s+|starting\s+)?`
  + `(${DATE})`
  + String.raw`\s*(?:-|–|—|to|until|untill|through|thru|till|and)\s*`
  + `(${DATE})`,
  'gi'
);

function monthFromWord(word) {
  const w = String(word || '').toLowerCase().replace(/\./g, '').slice(0, 3);
  return MONTH_WORDS[w] || null;
}

const fullYear = (n) => (n < 100 ? 2000 + n : n);

// Returns { year|null, month, day }. The year is null for a date that names none
// ("Aug 26"); the caller fills it in from the other end of the range.
// `dayFirst` settles the all-numeric form, where 08/09 could be either order.
function parseDate(raw, dayFirst) {
  const s = String(raw || '').trim().replace(/(\d)(st|nd|rd|th)\b/gi, '$1').replace(/,/g, ' ').replace(/\s+/g, ' ');

  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return valid({ year: +m[1], month: +m[2], day: +m[3] });

  m = s.match(/^(\d{1,2})[-/.\s]([A-Za-z]{3,9})[-/.\s](\d{2,4})$/);
  if (m) return valid({ year: fullYear(+m[3]), month: monthFromWord(m[2]), day: +m[1] });

  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (m) {
    const a = +m[1], b = +m[2];
    const useDayFirst = a > 12 ? true : (b > 12 ? false : !!dayFirst);
    return valid({ year: fullYear(+m[3]), month: useDayFirst ? b : a, day: useDayFirst ? a : b });
  }

  m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2})\s+(\d{4})$/);
  if (m) return valid({ year: +m[3], month: monthFromWord(m[1]), day: +m[2] });

  m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})$/);
  if (m) return valid({ year: null, month: monthFromWord(m[2]), day: +m[1] });

  m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2})$/);
  if (m) return valid({ year: null, month: monthFromWord(m[1]), day: +m[2] });

  return null;
}

function valid(d) {
  if (!d || !d.month || d.month < 1 || d.month > 12) return null;
  if (!d.day || d.day < 1 || d.day > 31) return null;
  if (d.year !== null && (d.year < 2000 || d.year > 2100)) return null;
  return d;
}

// "Billing period Aug 26 - Sep 26, 2026" states the year once, at the end. Carry
// it back to the start, rolling the year over when the range crosses December.
function fillYears(start, end) {
  if (start.year === null && end.year === null) return null;
  if (start.year === null) start = { ...start, year: end.month < start.month ? end.year - 1 : end.year };
  if (end.year === null) end = { ...end, year: end.month < start.month ? start.year + 1 : start.year };
  return { start, end };
}

const toUtc = d => Date.UTC(d.year, d.month - 1, d.day);
const iso = d => `${d.year}-${pad(d.month)}-${pad(d.day)}`;

// The first labelled period the invoice states, or null. A dozen line items can
// each carry the same period; the first one is as good as any, and taking the
// union of all of them would only blur a straddling cycle back into a long one.
function extractBillingPeriod(text) {
  const flat = String(text || '').replace(/\s+/g, ' ');
  PERIOD_RE.lastIndex = 0;
  let m;
  while ((m = PERIOD_RE.exec(flat)) !== null) {
    // Numeric dates are read US-style (month first) unless EITHER end proves
    // otherwise — "08/09/2026 to 26/09/2026" can only be day-first, and both
    // ends of one range are written the same way. Vendors billing this sheet in
    // USD write month-first, so that is the default.
    const dayFirst = [m[1], m[2]].some(d => /^\d{1,2}[-/.]\d{1,2}/.test(d) && +d.split(/[-/.]/)[0] > 12);
    const rawStart = parseDate(m[1], dayFirst);
    const rawEnd = parseDate(m[2], dayFirst);
    if (!rawStart || !rawEnd) continue;
    const filled = fillYears(rawStart, rawEnd);
    if (!filled) continue;
    const { start, end } = filled;
    if (toUtc(end) < toUtc(start)) continue;
    return { start: iso(start), end: iso(end), days: Math.round((toUtc(end) - toUtc(start)) / DAY_MS), text: m[0].trim() };
  }
  return null;
}

// How many days of the period fall in each calendar month. The end date is the
// day the next cycle starts ("until 2026-09-26"), so it is not counted; a
// same-day period still counts as one day.
function monthOfPeriod(period) {
  if (!period) return null;
  const s = Date.parse(`${period.start}T00:00:00Z`);
  let e = Date.parse(`${period.end}T00:00:00Z`);
  if (Number.isNaN(s) || Number.isNaN(e)) return null;
  if (e <= s) e = s + DAY_MS;

  const startMonth = period.start.slice(0, 7);
  // Quarterly and annual cycles have no meaningful majority month — a year is
  // spread over twelve of them — so they stay in the month they start in.
  if (e - s > 62 * DAY_MS) return { month: startMonth, rule: 'period-start' };

  const days = new Map();
  for (let t = s; t < e; t += DAY_MS) {
    const d = new Date(t);
    const key = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}`;
    days.set(key, (days.get(key) || 0) + 1);
  }
  const [month] = [...days.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0];
  return { month, rule: month === startMonth ? 'period-start' : 'period-majority', days: Object.fromEntries(days) };
}

const monthsApart = (a, b) => {
  const pa = String(a || '').split('-').map(Number);
  const pb = String(b || '').split('-').map(Number);
  if (pa.length !== 2 || pb.length !== 2 || pa.some(Number.isNaN) || pb.some(Number.isNaN)) return null;
  return Math.abs((pa[0] * 12 + pa[1]) - (pb[0] * 12 + pb[1]));
};

// The month a period puts an invoice in, measured against the month it would
// otherwise be filed under (the mail's month, or the folder it already sits in).
// Split out from invoiceMonth so a period already read and cached can be
// resolved again without the PDF.
function monthForPeriod(period, referenceMonth) {
  const resolved = period ? monthOfPeriod(period) : null;
  if (!resolved) return { month: referenceMonth, via: 'received' };

  // A period more than two months from the reference is far more likely to be a
  // misread than a real cycle — a contract term, or a renewal date pair. The
  // reference is the safer answer there.
  const gap = referenceMonth ? monthsApart(resolved.month, referenceMonth) : 0;
  if (referenceMonth && (gap === null || gap > 2)) {
    return { month: referenceMonth, via: 'received', ignoredPeriod: true };
  }
  return { month: resolved.month, via: resolved.rule };
}

// The month to file an invoice under, given its text and the month the mail
// arrived in. Returns { month, via, period } — `via` says which rule decided it,
// so a run can report the invoices it moved.
function invoiceMonth(text, receivedMonth) {
  const period = extractBillingPeriod(text);
  if (!period) return { month: receivedMonth, via: 'received', period: null };
  return { ...monthForPeriod(period, receivedMonth), period };
}

module.exports = { extractBillingPeriod, monthOfPeriod, monthForPeriod, invoiceMonth, parseDate, PERIOD_RE };
