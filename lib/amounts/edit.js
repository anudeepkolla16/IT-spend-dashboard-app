const { getGraphToken, resolveArchiveRoot, archiveFile, readJsonFile, writeJsonFile } = require('../graph');
const { verify, parseCookies, canEdit, editorList } = require('../session');
const excel = require('../excel');
const { openSpendSheet, cellValue, appendLog } = require('../spend-sheet');
const rulesLib = require('../invoices/rules');

// Editing the spend sheet from the dashboard, one cell at a time.
//
// The sheet is the system of record and several people can open the dashboard,
// so this is deliberately narrow: one cell per request, only the columns named
// below, only from an account on the editor list, and every write recorded in
// `_amount-log.json` with who made it and what was there before.
//
// The gate is here, in the handler. Hiding the controls in the page is a
// courtesy to the people who cannot use them, never the thing that stops
// anyone: the browser is not a place to enforce anything.

// The metadata columns a person may change, found by their header text so a
// reordered sheet still works. The application name itself is not editable —
// it is the key every other part of this app joins on, and renaming a row from
// here would silently orphan its invoices, its rules and its locks.
const FIELDS = {
  dept:          { label: 'Department',        match: /department/i,               max: 60 },
  poc:           { label: 'POC',               match: /poc/i,                      max: 60 },
  renewalDate:   { label: 'Renewal date',      match: /renewal/i,                  max: 40 },
  recurring:     { label: 'Recurring/Onetime', match: /recurring\s*\/\s*onetime/i, max: 20 },
  cycle:         { label: 'Frequency',         match: /frequency/i,                max: 30 },
  paymentMethod: { label: 'Payment method',    match: /payment\s*method/i,         max: 60 },
};
// Higher than any real charge in this sheet (the largest is ~41k) and low
// enough that a typo with an extra digit or two is caught rather than written.
const MAX_AMOUNT = 10000000;

// cellValue parses a cell into a number, which is right for an amount and
// wrong for everything else: it reads "1st of every month" as 1. A metadata
// column is read as the text a person would see, display text first, the way
// the grid itself reads its headers.
function rawCell(ctx, rowIdx, colIdx) {
  const t = ((ctx.text || [])[rowIdx] || [])[colIdx];
  if (t !== undefined && t !== null && String(t).trim() !== '') return String(t).trim();
  const v = ((ctx.values || [])[rowIdx] || [])[colIdx];
  return v === undefined || v === null ? '' : String(v).trim();
}

// The sheet's columns move about, so a column is found by its header text.
function colOf(ctx, match) {
  return ((ctx.grid && ctx.grid.headers) || []).findIndex(h => match.test(String(h == null ? '' : h)));
}

// The same reading spend-data uses: "Onetime", "One-time", "one time".
const ONE_TIME = /one\s*-?\s*time|onetime/i;

function sessionOf(req) {
  try { return verify(parseCookies(req.headers.cookie).session) || null; } catch (_) { return null; }
}

// An amount the owner typed is theirs, not the invoice sync's to revise. The
// same lock the five hand-set cells already use: without it the next run would
// total the month's invoices and ask to lower it straight back.
async function lockCell(token, driveId, app, month, value, by) {
  const root = await resolveArchiveRoot(token, driveId);
  if (!root || !root.resolved) return { locked: false, why: 'the invoice archive could not be found' };
  const path = archiveFile(root, rulesLib.RULES_FILE);
  const raw = await readJsonFile(token, driveId, path);
  // Never write a rules file we could not read: an empty or failed read here
  // would replace every vendor rule with nothing.
  if (!raw || !Array.isArray(raw.vendors)) return { locked: false, why: 'the vendor rules file could not be read' };
  const rules = rulesLib.normalizeRules(rulesLib.upgradeRules(raw).rules);
  rules.locks = (rules.locks || []).filter(l => !(l.app === app && l.month === month));
  if (value !== '') {
    rules.locks.push({ app, month, value, note: `Set in the dashboard by ${by} on ${new Date().toISOString().slice(0, 10)}.` });
  }
  await writeJsonFile(token, driveId, path, { ...rules, savedAt: new Date().toISOString() });
  return { locked: value !== '' };
}

module.exports = async (req, res) => {
  let sessionId = null, ctx = null, token = null;
  res.setHeader('Cache-Control', 'no-store');
  try {
    if (req.method !== 'POST') { res.status(405).json({ error: 'Method not allowed' }); return; }

    const session = sessionOf(req);
    const by = (session && session.email) || '';
    if (!canEdit(by)) {
      res.status(403).json({
        error: editorList().length
          ? 'Your account is not allowed to edit the sheet. Ask the sheet owner to add you to EDITOR_EMAILS.'
          : 'Editing is not configured: set EDITOR_EMAILS in Vercel to the accounts allowed to write to the sheet.',
      });
      return;
    }

    const body = req.body || {};
    const app = String(body.app || '').trim();
    const month = body.month == null ? '' : String(body.month).trim();
    const field = body.field == null ? '' : String(body.field).trim();
    if (!app) { res.status(400).json({ error: 'Which application?' }); return; }
    if (!month && !field) { res.status(400).json({ error: 'Send either a month (to edit an amount) or a field (to edit a detail)' }); return; }
    if (month && field) { res.status(400).json({ error: 'One thing at a time: a month or a field, not both' }); return; }
    if (field && !FIELDS[field]) { res.status(400).json({ error: `"${field}" is not an editable column. Editable: ${Object.keys(FIELDS).join(', ')}` }); return; }

    // The value, validated before anything is opened.
    let value, shown;
    if (month) {
      if (!/^\d{4}-\d{2}$/.test(month)) { res.status(400).json({ error: 'A month reads as 2026-08' }); return; }
      const raw = body.amount;
      if (raw === null || raw === undefined || String(raw).trim() === '') {
        value = ''; shown = 'cleared';
      } else {
        const n = Number(String(raw).replace(/[$,\s]/g, ''));
        if (!Number.isFinite(n)) { res.status(400).json({ error: 'That is not a number' }); return; }
        if (n < 0) { res.status(400).json({ error: 'An amount cannot be negative' }); return; }
        if (n > MAX_AMOUNT) { res.status(400).json({ error: `That is larger than any charge this sheet has ever held (max ${MAX_AMOUNT.toLocaleString('en-US')}) — check for an extra digit` }); return; }
        value = Math.round(n * 100) / 100; shown = String(value);
      }
    } else {
      const text = body.value == null ? '' : String(body.value).trim();
      if (text.length > FIELDS[field].max) { res.status(400).json({ error: `${FIELDS[field].label} is limited to ${FIELDS[field].max} characters` }); return; }
      value = text; shown = text || 'cleared';
    }

    token = await getGraphToken();
    const pre = await openSpendSheet(token);
    sessionId = await excel.createSession(token, pre.driveId, pre.itemId);
    ctx = await openSpendSheet(token, sessionId);

    const row = ctx.grid.apps.find(a => a.name === app);
    if (!row) {
      await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId); sessionId = null;
      res.status(400).json({ error: `"${app}" is not a row in the sheet` });
      return;
    }

    let colIdx = -1;
    if (month) {
      colIdx = ctx.grid.monthCols[month];
      if (colIdx === undefined) {
        await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId); sessionId = null;
        res.status(400).json({ error: `The sheet has no column for ${month}` });
        return;
      }
    }

    // What actually gets written. One cell for a month or a plain column; the
    // dashboard's "Cycle" is the one exception, because the sheet does not have
    // a Cycle column — spend-data derives it from two, where Recurring/Onetime
    // wins if it says "one" and Frequency decides otherwise. Writing only
    // Frequency would leave a row reading One-time no matter what was typed,
    // and typing "One-time" into Frequency would come back as Monthly.
    let writes;
    if (month) {
      writes = [{ colIdx, value }];
    } else if (field === 'cycle') {
      const freqCol = colOf(ctx, FIELDS.cycle.match), roCol = colOf(ctx, FIELDS.recurring.match);
      if (freqCol < 0 && roCol < 0) {
        await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId); sessionId = null;
        res.status(400).json({ error: 'The sheet has no Frequency or Recurring/Onetime column' });
        return;
      }
      const roNow = roCol < 0 ? '' : rawCell(ctx, row.rowIdx, roCol);
      if (ONE_TIME.test(value)) {
        if (roCol < 0) {
          await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId); sessionId = null;
          res.status(400).json({ error: 'The sheet has no Recurring/Onetime column, so a row cannot be marked one-time' });
          return;
        }
        writes = [{ colIdx: roCol, value: 'Onetime' }];
      } else {
        writes = [{ colIdx: freqCol, value }];
        // A row already marked one-time would ignore the frequency just typed.
        if (roCol >= 0 && ONE_TIME.test(roNow)) writes.push({ colIdx: roCol, value: 'Recurring' });
      }
    } else {
      colIdx = colOf(ctx, FIELDS[field].match);
      if (colIdx < 0) {
        await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId); sessionId = null;
        res.status(400).json({ error: `The sheet has no ${FIELDS[field].label} column` });
        return;
      }
      writes = [{ colIdx, value }];
    }

    const done = [];
    for (const w of writes) {
      const address = excel.cellAddress(ctx.start, row.rowIdx, w.colIdx);
      done.push({
        address,
        before: month ? cellValue(ctx.values, row.rowIdx, w.colIdx) : rawCell(ctx, row.rowIdx, w.colIdx),
        after: w.value,
      });
      await excel.writeCell(token, ctx.driveId, ctx.itemId, ctx.sheetName, address, w.value, sessionId);
    }
    await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId);
    sessionId = null;
    const { address, before } = done[0];

    // An amount stands on the owner's say-so, so the sync is told to leave it
    // alone. Best effort: the edit is already saved, and failing to lock it is
    // worth reporting, not worth undoing a correct write for.
    let lock = { locked: false };
    if (month) {
      try { lock = await lockCell(token, ctx.driveId, app, month, value, by); }
      catch (e) { lock = { locked: false, why: e.message || String(e) }; }
    }

    await appendLog(token, ctx.driveId, {
      at: new Date().toISOString(), by, attribution: 'dashboard-edit', sheet: ctx.sheetName,
      cells: done.map(d => ({ app, month: month || null, field: field || null, address: d.address, before: d.before, after: d.after })),
      failed: [],
    });

    res.status(200).json({ ok: true, app, month: month || null, field: field || null, address, before, after: value, shown, by, lock });
  } catch (err) {
    if (sessionId && ctx && token) await excel.closeSession(token, ctx.driveId, ctx.itemId, sessionId).catch(() => {});
    res.status(502).json({ error: err.message || String(err) });
  }
};

module.exports.FIELDS = FIELDS;
module.exports.MAX_AMOUNT = MAX_AMOUNT;
