const { getGraphToken } = require('../../lib/graph');
const { verify, parseCookies } = require('../../lib/session');
const excel = require('../../lib/excel');
const { openSpendSheet, cellValue, readAliasMap, writeAliasMap, appendLog } = require('../../lib/spend-sheet');

// Writes approved amounts into the spend sheet, one cell at a time through the
// Graph Excel API so the =SUM() totals, number formats and the other two sheets
// are left exactly as they were.

const MAX_CELLS = 200;

// The audit trail records who approved each write. Middleware has already
// rejected anyone without a valid session by the time this runs.
function whoami(req) {
  try {
    const session = verify(parseCookies(req.headers.cookie).session);
    return (session && (session.email || session.name)) || 'unknown';
  } catch (_) {
    return 'unknown';
  }
}

module.exports = async (req, res) => {
  let sessionId = null;
  let ctx = null;
  let token = null;
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const { cells, aliasUpdates, statementName, attribution } = req.body || {};
    if (!Array.isArray(cells) || !cells.length) {
      res.status(400).json({ error: 'Nothing to write — send at least one cell' });
      return;
    }
    if (cells.length > MAX_CELLS) {
      res.status(400).json({ error: `Too many cells in one request (max ${MAX_CELLS})` });
      return;
    }

    token = await getGraphToken();

    // A persisted workbook session keeps every write in one transaction against
    // one loaded copy of the file.
    const pre = await openSpendSheet(token);
    sessionId = await excel.createSession(token, pre.driveId, pre.itemId);
    // Re-read inside the session so the addresses we compute match the copy we
    // are about to edit.
    ctx = await openSpendSheet(token, sessionId);

    const rowByApp = new Map(ctx.grid.apps.map(a => [a.name, a.rowIdx]));

    const planned = [];
    const rejected = [];
    for (const cell of cells) {
      const app = String(cell && cell.app || '').trim();
      const month = String(cell && cell.month || '').trim();
      const amount = Number(cell && cell.amount);

      if (!app || !month) { rejected.push({ app, month, error: 'missing app or month' }); continue; }
      if (!Number.isFinite(amount)) { rejected.push({ app, month, error: 'amount is not a number' }); continue; }

      const rowIdx = rowByApp.has(app) ? rowByApp.get(app) : undefined;
      if (rowIdx === undefined) { rejected.push({ app, month, error: 'no row for that app in the sheet' }); continue; }
      const colIdx = ctx.grid.monthCols[month];
      if (colIdx === undefined) { rejected.push({ app, month, error: 'no column for that month in the sheet' }); continue; }

      const before = cellValue(ctx.values, rowIdx, colIdx);
      const address = excel.cellAddress(ctx.start, rowIdx, colIdx);
      const value = Math.round(amount * 100) / 100;
      planned.push({ app, month, address, value, before });
    }

    if (!planned.length) {
      await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
      res.status(400).json({ error: 'None of the requested cells could be matched to the sheet', rejected });
      return;
    }

    const results = await excel.writeCells(token, ctx.driveId, ctx.itemId, ctx.sheetName, planned, sessionId, 4);
    await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
    sessionId = null;

    const written = results.filter(r => r.ok);
    const failed = results.filter(r => !r.ok);

    // Remember any vendor label the user mapped by hand, so the same label
    // resolves itself next month.
    let aliasesSaved = 0;
    if (aliasUpdates && typeof aliasUpdates === 'object' && Object.keys(aliasUpdates).length) {
      const existing = await readAliasMap(token, ctx.driveId);
      const merged = { ...existing };
      for (const [key, app] of Object.entries(aliasUpdates)) {
        if (!key || !app) continue;
        if (!rowByApp.has(String(app))) continue; // never save a mapping to a row that isn't there
        merged[key] = String(app);
        aliasesSaved++;
      }
      if (aliasesSaved) await writeAliasMap(token, ctx.driveId, merged);
    }

    const by = whoami(req);
    await appendLog(token, ctx.driveId, {
      at: new Date().toISOString(),
      by,
      statement: statementName || null,
      attribution: attribution || 'statement',
      sheet: ctx.sheetName,
      cells: written.map(w => ({ app: w.app, month: w.month, address: w.address, before: w.before, after: w.value })),
      failed: failed.map(f => ({ app: f.app, month: f.month, error: f.error })),
    });

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({
      ok: true,
      written: written.length,
      failed: failed.length,
      by,
      cells: written.map(w => ({ app: w.app, month: w.month, address: w.address, before: w.before, after: w.value })),
      errors: failed.map(f => ({ app: f.app, month: f.month, error: f.error })),
      rejected,
      aliasesSaved,
    });
  } catch (err) {
    if (sessionId && ctx && token) await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
