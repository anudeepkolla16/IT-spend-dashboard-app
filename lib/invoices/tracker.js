// Ticking the "Invoices tracker" worksheet.
//
// Two callers want the same write. The mail sync ticks what it has just filed;
// the checklist's backfill ticks everything the archive already holds — the
// tracker was never filled in for the hundreds of invoices archived by hand
// before any of this existed, so every one of them read FALSE. The rules are
// identical either way and are worth stating once:
//
//   · never untick — a TRUE somebody set by hand is a fact this code does not
//     have, and the archive not holding a copy is not evidence against it
//   · one write per app-month, however many invoices back it
//   · never invent a row or a column: an app or a month the tracker does not
//     carry is reported, not created
const excel = require('../excel');
const { spendFilePath } = require('../spend-sheet');
const { resolveDriveId } = require('../graph');
const { norm } = require('../vendor-map');

// A backfill covers every app in the sheet across every month it tracks, so the
// honest ceiling is roughly rows × columns. Well past that is a malformed
// request rather than a real archive, and each mark costs a cell write.
const MAX_MARKS = 3000;

// How many marks one apply request handles. A backfill of a year's archive is
// several hundred cells, and writing them in one request outran the function's
// 60 seconds — the client loops instead, advancing an offset over its own list
// of marks. Slicing the marks rather than the plan keeps the offsets stable:
// cells ticked by an earlier chunk drop out of the next plan, which would shift
// every offset after them.
const APPLY_CHUNK = 120;

// The tracker path needs the workbook, not the spend grid. Opening the spend
// sheet to get at its ids costs a used-range read of the largest worksheet in
// the file, for two strings.
async function locateWorkbook(token) {
  const upn = (process.env.TARGET_USER_UPN || '').trim();
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
  const driveId = await resolveDriveId(token, upn);
  const itemId = await excel.resolveItemId(token, driveId, spendFilePath());
  return { driveId, itemId };
}

// The tracker is a second worksheet maintained by hand beside the spend grid,
// and the two disagree on how a few rows are spelled: "Tmobile" against
// "TMobile", "Render " against "Render". An exact match skips those silently —
// the app looks like it ticked everything while two rows stay FALSE forever —
// so fall back to the same normalized form every other name here is matched by.
// A normalized key shared by two tracker rows resolves to neither: ticking the
// wrong row is worse than reporting the pair as having no cell.
function trackerRowIndex(grid) {
  const exact = new Map();
  const byNorm = new Map();
  for (const a of grid.apps) {
    exact.set(a.name, a.rowIdx);
    const key = norm(a.name);
    if (!key) continue;
    byNorm.set(key, byNorm.has(key) ? null : a.rowIdx);
  }
  return (app) => {
    const hit = exact.get(app);
    if (hit !== undefined) return hit;
    const loose = byNorm.get(norm(app));
    return loose == null ? undefined : loose;
  };
}

// Works out which tracker cells actually need writing. Kept separate from the
// Graph calls so the rules are testable: one cell per app-month no matter how
// many invoices arrived, nothing already ticked, nothing off-grid.
function planTrackerCells(marks, grid, used) {
  const rowFor = trackerRowIndex(grid);
  const cells = [];
  const seenAddresses = new Set();
  for (const { app, month } of marks || []) {
    const rowIdx = rowFor(app);
    const colIdx = grid.monthCols[month];
    if (rowIdx === undefined || colIdx === undefined) continue;
    const addr = excel.cellAddress(used.start, rowIdx, colIdx);
    if (seenAddresses.has(addr)) continue; // never write one cell twice
    seenAddresses.add(addr);
    const raw = (used.values[rowIdx] || [])[colIdx];
    const existing = String(raw == null ? '' : raw).trim();
    if (/^true$/i.test(existing)) continue; // already ticked
    cells.push({ app, month, address: addr, value: true });
  }
  return cells;
}

// The tracker is its own worksheet beside the spend grid, carrying the same app
// rows and month columns. Returns null when the workbook has no such sheet,
// which is not an error: a workbook may simply not track invoices.
async function openTracker(token, driveId, itemId, sessionId) {
  const sheets = await excel.listWorksheets(token, driveId, itemId, sessionId);
  const name = sheets.map(s => s.name).find(n => /tracker/i.test(n));
  if (!name) return null;
  const used = await excel.readUsedRange(token, driveId, itemId, name, sessionId);
  return { name, used, grid: excel.locateGrid(used.values, used.text) };
}

async function markTracker(token, driveId, itemId, marks, sessionId) {
  const tracker = await openTracker(token, driveId, itemId, sessionId);
  if (!tracker) return { sheet: null, marked: 0 };

  const cells = planTrackerCells(marks, tracker.grid, tracker.used);
  if (!cells.length) return { sheet: tracker.name, marked: 0 };

  const results = await excel.writeCells(token, driveId, itemId, tracker.name, cells, sessionId, 4);
  return { sheet: tracker.name, marked: results.filter(r => r.ok).length };
}

// Marks arrive from the page, which is the only place the folder->app join is
// worked out, so they are treated as a request rather than as truth: anything
// malformed is dropped here and anything off-grid is reported by the plan.
function normalizeMarks(marks) {
  const out = [];
  const seen = new Set();
  for (const m of Array.isArray(marks) ? marks : []) {
    if (!m || typeof m !== 'object') continue;
    const app = typeof m.app === 'string' ? m.app.trim() : '';
    const month = typeof m.month === 'string' ? m.month.trim() : '';
    if (!app || !/^\d{4}-\d{2}$/.test(month)) continue;
    const key = `${app} ${month}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ app, month });
    if (out.length >= MAX_MARKS) break;
  }
  return out;
}

// What a backfill would do, without doing it. Splits the marks three ways so
// the confirmation says something the user can check: what gets ticked, what
// was already ticked, and what the tracker has no cell for — that last group is
// how an archive folder with no row in the sheet becomes visible.
async function planBackfill(token, marks, sessionId) {
  const wanted = normalizeMarks(marks);
  const { driveId, itemId } = await locateWorkbook(token);
  const tracker = await openTracker(token, driveId, itemId, sessionId);
  if (!tracker) {
    return { sheet: null, requested: wanted.length, cells: [], alreadyTicked: [], offGrid: wanted, chunk: APPLY_CHUNK };
  }

  const rowFor = trackerRowIndex(tracker.grid);
  const offGrid = wanted.filter(m =>
    rowFor(m.app) === undefined || tracker.grid.monthCols[m.month] === undefined);
  const cells = planTrackerCells(wanted, tracker.grid, tracker.used);
  const toWrite = new Set(cells.map(c => `${c.app} ${c.month}`));
  const offGridKeys = new Set(offGrid.map(m => `${m.app} ${m.month}`));
  const alreadyTicked = wanted.filter(m => {
    const key = `${m.app} ${m.month}`;
    return !toWrite.has(key) && !offGridKeys.has(key);
  });

  return { sheet: tracker.name, requested: wanted.length, cells, alreadyTicked, offGrid, chunk: APPLY_CHUNK };
}

// Re-plans against the live sheet rather than trusting the addresses the plan
// handed out: the workbook can be edited between the two calls, and a cell
// address from the client is the one thing here that must never be believed.
//
// Handles one chunk of the marks and reports where the next one starts, so the
// client can loop. Re-ticking a chunk is harmless — an already-ticked cell is
// skipped — so a retry after a failed chunk costs nothing.
async function applyBackfill(token, marks, options) {
  const opts = options || {};
  const wanted = normalizeMarks(marks);
  const from = Number.isInteger(opts.offset) && opts.offset > 0 ? opts.offset : 0;
  const slice = wanted.slice(from, from + APPLY_CHUNK);
  const nextOffset = from + slice.length;

  const { driveId, itemId } = await locateWorkbook(token);
  const result = await markTracker(token, driveId, itemId, slice, opts.sessionId);
  return {
    ...result,
    requested: wanted.length,
    offset: from,
    nextOffset: nextOffset < wanted.length ? nextOffset : null,
    done: nextOffset >= wanted.length,
  };
}

module.exports = { planTrackerCells, trackerRowIndex, openTracker, markTracker, planBackfill, applyBackfill, normalizeMarks, MAX_MARKS, APPLY_CHUNK };
