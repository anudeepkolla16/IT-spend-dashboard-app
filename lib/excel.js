const { encodeGraphPath } = require('./graph');

// Writing amounts back into the spend workbook goes through the Graph Excel API
// (not SheetJS) on purpose: the sheet carries =SUM() totals, number formats and
// two sibling sheets. Downloading, rewriting with SheetJS and re-uploading would
// flatten formulas and drop formatting. The Excel API edits cells in place.

const MONTH_RE = /^([A-Za-z]{3})-(\d{2})$/;
const MONTH_MAP = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// "Jan-26" -> "2026-01". Also accepts Excel's date rendering of the same header
// ("01/01/2026"), which is how the column headers come back when the cell is a
// real date rather than text.
function normMonthHeader(h) {
  const s = String(h == null ? '' : h).trim();
  const m = s.match(MONTH_RE);
  if (m) {
    const mon = MONTH_MAP[m[1].toLowerCase()];
    if (mon) return `20${m[2]}-${String(mon).padStart(2, '0')}`;
  }
  const d = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (d) return `${d[3]}-${String(Number(d[1])).padStart(2, '0')}`;
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
  const res = await fetch(url, { method, headers, body: body ? JSON.stringify(body) : undefined });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excel API ${method} failed (${res.status}): ${text.slice(0, 300)}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function resolveItemId(token, driveId, path) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id,name`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
function locateGrid(values) {
  let headerRowIdx = -1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const hasMonth = row.some(c => normMonthHeader(c));
    const hasName = row.some(c => /application|sw\s*\/\s*license/i.test(String(c == null ? '' : c)));
    if (hasMonth && hasName) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) throw new Error('Could not find the header row (a row with an "APPLICATION / SW / LICENSE" column and month columns) in the spend sheet.');

  const headers = values[headerRowIdx] || [];
  const nameCol = headers.findIndex(h => /application|sw\s*\/\s*license/i.test(String(h == null ? '' : h)));
  const monthCols = {};
  headers.forEach((h, idx) => {
    const ym = normMonthHeader(h);
    if (ym && monthCols[ym] === undefined) monthCols[ym] = idx;
  });

  const apps = [];
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const name = String((values[i] || [])[nameCol] == null ? '' : (values[i] || [])[nameCol]).trim();
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

// Writes run through a small pool: cell-by-cell (never whole rows/columns) so a
// cell we did not intend to touch is never rewritten, and bounded so a large
// month does not stack up dozens of simultaneous Graph calls.
async function writeCells(token, driveId, itemId, sheetName, cells, sessionId, poolSize) {
  const results = [];
  let cursor = 0;
  const POOL = poolSize || 4;
  await Promise.all(Array.from({ length: Math.min(POOL, cells.length) }, async () => {
    while (cursor < cells.length) {
      const cell = cells[cursor++];
      try {
        await writeCell(token, driveId, itemId, sheetName, cell.address, cell.value, sessionId);
        results.push({ ...cell, ok: true });
      } catch (e) {
        results.push({ ...cell, ok: false, error: e.message });
      }
    }
  }));
  return results;
}

module.exports = {
  normMonthHeader, monthLabel, colLetter, resolveItemId, createSession, closeSession,
  listWorksheets, readUsedRange, pickSpendSheet, locateGrid, cellAddress, writeCell, writeCells,
};
