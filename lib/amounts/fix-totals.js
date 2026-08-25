const { getGraphToken } = require('../graph');
const { verify, parseCookies } = require('../session');
const excel = require('../excel');
const { openSpendSheet, appendLog } = require('../spend-sheet');

// Repairs the Total row's SUM ranges so they span every app row.
//
// The ranges were written once and have gone stale as apps were added: the month
// columns do not all sum the same range, so some months quietly leave rows out.
// Rather than hardcode row numbers — which shift every time a row is inserted —
// the correct range is derived from the live sheet each time this runs, so it is
// safe to re-run whenever the sheet grows.

function plan(grid, used) {
  if (!grid.apps.length) throw new Error('No app rows found in the sheet.');
  const firstRow = used.start.row + grid.apps[0].rowIdx + 1;              // 1-based
  const lastRow = used.start.row + grid.apps[grid.apps.length - 1].rowIdx + 1;

  // The Total row is the one whose name cell reads "Total" — it is excluded from
  // grid.apps, so find it directly.
  let totalRowIdx = -1;
  for (let i = grid.headerRowIdx + 1; i < used.values.length; i++) {
    const name = String((used.values[i] || [])[grid.nameCol] ?? '').trim();
    if (/^total$/i.test(name)) { totalRowIdx = i; break; }
  }
  if (totalRowIdx === -1) throw new Error('Could not find the Total row in the sheet.');
  const totalRow = used.start.row + totalRowIdx + 1;

  const changes = [];
  for (const [month, colIdx] of Object.entries(grid.monthCols)) {
    const col = excel.colLetter(used.start.col + colIdx);
    const address = `${col}${totalRow}`;
    const want = `=SUM(${col}${firstRow}:${col}${lastRow})`;
    const have = String(((used.formulas || [])[totalRowIdx] || [])[colIdx] ?? '').trim();
    changes.push({ month, address, from: have || '(empty)', to: want, needsFix: have !== want });
  }
  changes.sort((a, b) => a.month.localeCompare(b.month));
  return { firstRow, lastRow, totalRow, changes };
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
    const apply = !!(req.body && req.body.apply);

    token = await getGraphToken();
    const pre = await openSpendSheet(token);
    sessionId = await excel.createSession(token, pre.driveId, pre.itemId);
    ctx = await openSpendSheet(token, sessionId);

    const used = { values: ctx.values, formulas: ctx.formulas, start: ctx.start };
    const planned = plan(ctx.grid, used);
    const toFix = planned.changes.filter(c => c.needsFix);

    if (!apply || !toFix.length) {
      await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ ok: true, applied: false, sheet: ctx.sheetName, ...planned, needingFix: toFix.length });
      return;
    }

    const done = [];
    for (const c of toFix) {
      try {
        await excel.writeFormula(token, ctx.driveId, ctx.itemId, ctx.sheetName, c.address, c.to, sessionId);
        done.push(c);
      } catch (e) {
        c.error = e.message;
      }
    }
    await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
    sessionId = null;

    const by = (() => {
      try { const s = verify(parseCookies(req.headers.cookie).session); return (s && (s.email || s.name)) || 'unknown'; }
      catch (_) { return 'unknown'; }
    })();

    if (done.length) {
      await appendLog(token, ctx.driveId, {
        at: new Date().toISOString(), by, source: 'fix-totals', sheet: ctx.sheetName,
        cells: done.map(c => ({ app: 'Total', month: c.month, address: c.address, before: c.from, after: c.to })),
        failed: toFix.filter(c => c.error).map(c => ({ month: c.month, error: c.error })),
      });
    }

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({
      ok: true, applied: true, by, sheet: ctx.sheetName,
      firstRow: planned.firstRow, lastRow: planned.lastRow, totalRow: planned.totalRow,
      fixed: done.map(c => ({ month: c.month, address: c.address, from: c.from, to: c.to })),
      failed: toFix.filter(c => c.error).map(c => ({ month: c.month, address: c.address, error: c.error })),
      unchanged: planned.changes.filter(c => !c.needsFix).map(c => c.month),
    });
  } catch (err) {
    if (sessionId && ctx && token) await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};

module.exports.plan = plan;
