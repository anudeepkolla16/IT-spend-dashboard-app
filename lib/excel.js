const { encodeGraphPath, graphFetch } = require('./graph');

// Writing amounts back into the spend workbook goes through the Graph Excel API
// (not SheetJS) on purpose: the sheet carries =SUM() totals, number formats and
// two sibling sheets. Downloading, rewriting with SheetJS and re-uploading would
// flatten formulas and drop formatting. The Excel API edits cells in place.

const MONTH_RE = /^([A-Za-z]{3})-(\d{2})$/;
const MONTH_MAP = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Excel stores a date as a serial day count from 1899-12-30. The month headers
// in the spend sheet are real dates formatted as "mmm-yy", so the Graph Excel
// API reports them as bare numbers in `values` and only as "Jan-26" in `text`.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const SERIAL_MIN = 20000; // ~1954, below any plausible month header
const SERIAL_MAX = 80000; // ~2119
function monthFromSerial(n) {
  if (!Number.isFinite(n) || n < SERIAL_MIN || n > SERIAL_MAX) return null;
  const d = new Date(EXCEL_EPOCH_MS + Math.floor(n) * 86400000);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

// "Jan-26" -> "2026-01". Also accepts the other two ways the same header reaches
// us: Excel's date rendering ("01/01/2026") and, when the cell is a real date
// read through the Graph Excel API, the raw serial number (46023).
function normMonthHeader(h) {
  if (typeof h === 'number') return monthFromSerial(h);
  const s = String(h == null ? '' : h).trim();
  const m = s.match(MONTH_RE);
  if (m) {
    const mon = MONTH_MAP[m[1].toLowerCase()];
    if (mon) return `20${m[2]}-${String(mon).padStart(2, '0')}`;
  }
  const d = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (d) return `${d[3]}-${String(Number(d[1])).padStart(2, '0')}`;
  const iso = s.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (iso) return `${iso[1]}-${iso[2]}`;
  // A serial that arrived as a string, e.g. "46023".
  if (/^\d+(\.\d+)?$/.test(s)) return monthFromSerial(Number(s));
  return null;
}

function monthLabel(ym) {
  const [y, m] = String(ym).split('-');
  return `${MONTH_NAMES[Number(m) - 1]}-${String(y).slice(2)}`;
}

// 0-based column index -> spreadsheet column letters (0 -> A, 26 -> AA).
function colLetter(idx) {
  let n = idx + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function wbBase(driveId, itemId) {
  return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;
}

async function graphCall(token, url, { method = 'GET', body, sessionId } = {}) {
  const headers = { Authorization: `Bearer ${token}` };
  if (body) headers['Content-Type'] = 'application/json';
  if (sessionId) headers['workbook-session-id'] = sessionId;
  const res = await graphFetch(url, { method, headers, body: body ? JSON.stringify(body) : undefined });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excel API ${method} failed (${res.status}): ${text.slice(0, 300)}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function resolveItemId(token, driveId, path) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id,name`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Could not resolve workbook "${path}" (${res.status}): ${text.slice(0, 200)}`);
  }
  return (await res.json()).id;
}

// A persisted session batches our cell writes into one workbook transaction and
// keeps Graph from reloading the file on every PATCH.
async function createSession(token, driveId, itemId) {
  const json = await graphCall(token, `${wbBase(driveId, itemId)}/createSession`, {
    method: 'POST',
    body: { persistChanges: true },
  });
  return json.id;
}

async function closeSession(token, driveId, itemId, sessionId) {
  if (!sessionId) return;
  try {
    await graphCall(token, `${wbBase(driveId, itemId)}/closeSession`, { method: 'POST', sessionId });
  } catch (_) {
    // A session that failed to close expires on its own; never fail the request for it.
  }
}

async function listWorksheets(token, driveId, itemId, sessionId) {
  const json = await graphCall(token, `${wbBase(driveId, itemId)}/worksheets?$select=id,name`, { sessionId });
  return json.value || [];
}

// usedRange addresses come back as "Spendings!A1:S72" — we need the top-left
// offsets to turn array indices into real cell addresses.
function parseRangeStart(address) {
  const m = String(address || '').match(/!\$?([A-Z]+)\$?(\d+)/);
  if (!m) return { col: 0, row: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) - 1 };
}

async function readUsedRange(token, driveId, itemId, sheetName, sessionId) {
  const url = `${wbBase(driveId, itemId)}/worksheets('${encodeURIComponent(sheetName)}')/usedRange(valuesOnly=false)?$select=address,values,formulas,text`;
  const json = await graphCall(token, url, { sessionId });
  const start = parseRangeStart(json.address);
  return { values: json.values || [], formulas: json.formulas || [], text: json.text || [], start, address: json.address };
}

// Find the sheet that actually holds the monthly amounts. The workbook also has
// an "Invoices tracker" sheet with the same column shape but TRUE/FALSE cells,
// so name preference comes first and tracker-ish sheets are excluded.
function pickSpendSheet(sheetNames) {
  const preferred = (process.env.SPEND_SHEET_NAME || 'Spendings').trim().toLowerCase();
  const exact = sheetNames.find(n => n.trim().toLowerCase() === preferred);
  if (exact) return exact;
  return sheetNames.find(n => !/tracker|mail/i.test(n)) || sheetNames[0];
}

// Locate the header row, the app-name column, and the month columns, the same
// way api/spend-data.js does when reading — so writes land where reads look.
// `text` is the display-string grid Graph returns alongside `values`. It is what
// makes date-formatted headers legible ("Jan-26" rather than 46023), so header
// detection reads it first and falls back to the raw value.
function locateGrid(values, text) {
  const at = (i, j) => {
    const t = (text && text[i] || [])[j];
    if (t !== undefined && t !== null && String(t).trim() !== '') return t;
    return (values[i] || [])[j];
  };
  const rowCells = (i) => {
    const width = Math.max((values[i] || []).length, ((text && text[i]) || []).length);
    return Array.from({ length: width }, (_, j) => at(i, j));
  };
  const isNameCell = (c) => /application|sw\s*\/\s*license/i.test(String(c == null ? '' : c));

  let headerRowIdx = -1;
  let headers = [];
  for (let i = 0; i < values.length; i++) {
    const row = rowCells(i);
    if (row.some(c => normMonthHeader(c)) && row.some(isNameCell)) {
      headerRowIdx = i;
      headers = row;
      break;
    }
  }
  if (headerRowIdx === -1) throw new Error('Could not find the header row (a row with an "APPLICATION / SW / LICENSE" column and month columns) in the spend sheet.');

  const nameCol = headers.findIndex(isNameCell);
  const monthCols = {};
  headers.forEach((h, idx) => {
    const ym = normMonthHeader(h);
    if (ym && monthCols[ym] === undefined) monthCols[ym] = idx;
  });

  const apps = [];
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const name = String(at(i, nameCol) == null ? '' : at(i, nameCol)).trim();
    if (!name || /^total$/i.test(name)) continue;
    apps.push({ name, rowIdx: i });
  }
  return { headerRowIdx, nameCol, monthCols, apps };
}

function cellAddress(start, rowIdx, colIdx) {
  return `${colLetter(start.col + colIdx)}${start.row + rowIdx + 1}`;
}

async function writeCell(token, driveId, itemId, sheetName, address, value, sessionId) {
  const url = `${wbBase(driveId, itemId)}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${address}')`;
  // Only `values` is sent, so each cell keeps its existing number format.
  await graphCall(token, url, { method: 'PATCH', body: { values: [[value]] }, sessionId });
}

// Sets a cell's formula rather than its value. Used to repair the Total row,
// whose SUM ranges go stale as apps are added to the sheet.
async function writeFormula(token, driveId, itemId, sheetName, address, formula, sessionId) {
  const url = `${wbBase(driveId, itemId)}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${address}')`;
  await graphCall(token, url, { method: 'PATCH', body: { formulas: [[formula]] }, sessionId });
}

// Splits "H3" into the row and column a range address is built from. Anything
// that is not a plain single-cell address (a range, a sheet-qualified name, an
// absolute $H$3) is left un-mergeable rather than guessed at.
function parseAddress(address) {
  const m = /^([A-Z]+)(\d+)$/.exec(String(address || ''));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) - 1 };
}

// Groups planned cells into the longest runs of neighbours on one row, so a row
// ticked across eight months costs one call instead of eight. Only cells that
// were going to be written anyway are ever merged — a gap in the columns ends
// the run — so this still never rewrites a cell we did not intend to touch.
function groupIntoRuns(cells) {
  const parsed = [];
  const loners = [];
  for (const cell of cells || []) {
    const at = parseAddress(cell.address);
    if (at) parsed.push({ cell, ...at }); else loners.push({ cells: [cell], address: cell.address });
  }
  parsed.sort((a, b) => a.row - b.row || a.col - b.col);

  const runs = [];
  for (const item of parsed) {
    const last = runs[runs.length - 1];
    if (last && last.row === item.row && item.col === last.endCol + 1) {
      last.cells.push(item.cell);
      last.endCol = item.col;
      last.address = `${colLetter(last.startCol)}${last.row + 1}:${colLetter(last.endCol)}${last.row + 1}`;
      continue;
    }
    runs.push({ row: item.row, startCol: item.col, endCol: item.col, address: item.cell.address, cells: [item.cell] });
  }
  return [...runs, ...loners];
}

// Writes run through a small pool: never whole rows or columns, so a cell we did
// not intend to touch is never rewritten, and bounded so a large month does not
// stack up dozens of simultaneous Graph calls. Neighbouring cells on one row go
// out together — a backfill ticking hundreds of cells one PATCH at a time takes
// longer than the request is allowed to live.
async function writeCells(token, driveId, itemId, sheetName, cells, sessionId, poolSize) {
  const runs = groupIntoRuns(cells);
  const results = [];
  let cursor = 0;
  const POOL = poolSize || 4;
  await Promise.all(Array.from({ length: Math.min(POOL, runs.length) }, async () => {
    while (cursor < runs.length) {
      const run = runs[cursor++];
      const url = `${wbBase(driveId, itemId)}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${run.address}')`;
      try {
        // Only `values` is sent, so each cell keeps its existing number format.
        await graphCall(token, url, { method: 'PATCH', body: { values: [run.cells.map(c => c.value)] }, sessionId });
        for (const cell of run.cells) results.push({ ...cell, ok: true });
      } catch (e) {
        // A run fails as a unit: none of its cells are known to have landed.
        for (const cell of run.cells) results.push({ ...cell, ok: false, error: e.message });
      }
    }
  }));
  return results;
}

module.exports = {
  normMonthHeader, monthLabel, colLetter, resolveItemId, createSession, closeSession,
  listWorksheets, readUsedRange, pickSpendSheet, locateGrid, cellAddress, writeCell, writeFormula, writeCells, groupIntoRuns,
};
