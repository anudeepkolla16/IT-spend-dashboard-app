const XLSX = require('xlsx');

// Parses the monthly statement workbook Finance hands over. Two layouts appear
// in the same file and both are supported by locating columns from the header
// text rather than by position:
//   Sheet1: BOA | Date | Month | Description | Amount | Comments
//   Sheet2: Transaction Code | Statment Period | Date | Description | Amount | ...
// "Comments" is the hand-typed vendor label; Sheet2 has none, so the vendor is
// inferred from the raw bank descriptor instead.

const MONTH_MAP = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

function parseAmount(raw) {
  if (typeof raw === 'number') return raw;
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  // Accounting style: (1,234.56) means negative.
  const paren = /^\((.*)\)$/.test(s);
  const cleaned = s.replace(/[()]/g, '').replace(/[^0-9.\-]/g, '');
  if (!cleaned || cleaned === '-' || cleaned === '.') return null;
  const n = parseFloat(cleaned);
  if (Number.isNaN(n)) return null;
  return paren ? -Math.abs(n) : n;
}

// Accepts "6/2/2026", "2026-06-02", a JS Date (cellDates), "26-Jun" and "Jun-26".
function parseWhen(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return { year: raw.getUTCFullYear(), month: raw.getUTCMonth() + 1, day: raw.getUTCDate() };
  }
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return { year: Number(m[3]), month: Number(m[1]), day: Number(m[2]) };
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return { year: Number(m[1]), month: Number(m[2]), day: Number(m[3]) };
  // Statement-period style: "26-Jun" (yy-Mon) or "Jun-26" (Mon-yy).
  m = s.match(/^(\d{2})-([A-Za-z]{3})$/);
  if (m && MONTH_MAP[m[2].toLowerCase()]) return { year: 2000 + Number(m[1]), month: MONTH_MAP[m[2].toLowerCase()], day: null };
  m = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (m && MONTH_MAP[m[1].toLowerCase()]) return { year: 2000 + Number(m[2]), month: MONTH_MAP[m[1].toLowerCase()], day: null };
  return null;
}

const ymOf = (w) => (w ? `${w.year}-${String(w.month).padStart(2, '0')}` : null);

// Bank descriptors embed the date the charge was actually made, e.g.
// "ANTHROPIC 05/31 PURCHASE ...". That can fall in the month before the one the
// statement covers, which is exactly where the two attribution rules diverge.
function purchaseMonthFromDescription(desc, fallback) {
  if (!fallback) return null;
  const m = String(desc || '').match(/\b(\d{1,2})\/(\d{1,2})\b(?=\s|$)/);
  if (!m) return null;
  const month = Number(m[1]);
  const day = Number(m[2]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  // Only the month/day is present, so infer the year from the statement month,
  // rolling back a year when a December purchase appears on a January statement.
  let year = fallback.year;
  if (month === 12 && fallback.month === 1) year -= 1;
  return { year, month, day };
}

function findHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row = (rows[i] || []).map(c => String(c == null ? '' : c).trim());
    const hasAmount = row.some(c => /^amount$/i.test(c));
    const hasDesc = row.some(c => /descriptio?n/i.test(c));
    if (hasAmount && hasDesc) return i;
  }
  return -1;
}

const findCol = (headers, re) => headers.findIndex(h => re.test(String(h == null ? '' : h).trim()));

// Sheet1 records charges as negatives, Sheet2 as positives. Rather than taking
// absolute values (which would silently turn a refund into a charge), each sheet
// gets a polarity from its majority sign; rows against that sign stay negative
// and net the app's total down, and get surfaced as refunds.
function sheetPolarity(amounts) {
  let neg = 0, pos = 0;
  for (const a of amounts) { if (a < 0) neg++; else if (a > 0) pos++; }
  return neg > pos ? -1 : 1;
}

function parseStatement(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const txns = [];
  const sheetReports = [];

  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', raw: true });
    const headerIdx = findHeader(rows);
    if (headerIdx === -1) { sheetReports.push({ sheet: sheetName, skipped: 'no Amount/Description header row' }); continue; }

    const headers = rows[headerIdx] || [];
    const amountCol = findCol(headers, /^amount$/i);
    const descCol = findCol(headers, /descriptio?n/i);
    const dateCol = findCol(headers, /^date$/i);
    const vendorCol = findCol(headers, /^comments?$/i);
    const periodCol = findCol(headers, /^(month|statment period|statement period)$/i);

    const body = rows.slice(headerIdx + 1);
    const rawAmounts = body.map(r => parseAmount((r || [])[amountCol])).filter(a => a !== null && a !== 0);
    if (!rawAmounts.length) { sheetReports.push({ sheet: sheetName, skipped: 'no numeric amounts' }); continue; }
    const polarity = sheetPolarity(rawAmounts);

    let count = 0;
    body.forEach((row, i) => {
      const raw = parseAmount((row || [])[amountCol]);
      if (raw === null || raw === 0) return;
      const description = String((row || [])[descCol] == null ? '' : (row || [])[descCol]).trim();
      const vendorLabel = vendorCol >= 0 ? String((row || [])[vendorCol] == null ? '' : (row || [])[vendorCol]).trim() : '';
      if (!description && !vendorLabel) return;

      const txnWhen = dateCol >= 0 ? parseWhen((row || [])[dateCol]) : null;
      const periodWhen = periodCol >= 0 ? parseWhen((row || [])[periodCol]) : null;
      // Statement month: the period column if present, else the posting date.
      const statementMonth = ymOf(periodWhen) || ymOf(txnWhen);
      // Transaction month: the date embedded in the descriptor if there is one,
      // else the posting date.
      const purchaseWhen = purchaseMonthFromDescription(description, txnWhen || periodWhen);
      const transactionMonth = ymOf(purchaseWhen) || ymOf(txnWhen) || statementMonth;
      if (!statementMonth && !transactionMonth) return;

      const amount = raw * polarity; // charges positive, refunds negative
      txns.push({
        sheet: sheetName,
        row: headerIdx + 2 + i, // 1-based row in the source workbook, for traceability
        date: txnWhen ? `${txnWhen.year}-${String(txnWhen.month).padStart(2, '0')}-${String(txnWhen.day || 1).padStart(2, '0')}` : '',
        description,
        vendorLabel,
        amount,
        isRefund: amount < 0,
        statementMonth: statementMonth || transactionMonth,
        transactionMonth: transactionMonth || statementMonth,
      });
      count++;
    });
    sheetReports.push({ sheet: sheetName, transactions: count, polarity });
  }

  return { txns, sheetReports };
}

// Which month a transaction counts toward. Default follows the statement period,
// which is what the largest lines in the existing sheet (Bubble Starter, Google
// cloud) were reconciled against.
function monthFor(txn, attribution) {
  return attribution === 'transaction' ? (txn.transactionMonth || txn.statementMonth) : (txn.statementMonth || txn.transactionMonth);
}

// Statement files are named inconsistently ("apps Apr.xlsx", "Saras appas &
// subscription June 2026.xlsx"), so the covered month is taken from the data,
// not the filename. Returns months ordered by how many transactions each holds.
function monthsCovered(txns, attribution) {
  const counts = {};
  for (const t of txns) {
    const ym = monthFor(t, attribution);
    if (ym) counts[ym] = (counts[ym] || 0) + 1;
  }
  return Object.entries(counts)
    .map(([month, transactions]) => ({ month, transactions }))
    .sort((a, b) => b.transactions - a.transactions || a.month.localeCompare(b.month));
}

module.exports = { parseStatement, monthFor, monthsCovered, parseAmount, parseWhen, purchaseMonthFromDescription };
