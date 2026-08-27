const { getGraphToken, resolveShare, graphFetch } = require('../graph');
const { parseStatement, monthFor, monthsCovered } = require('../statement');
const { buildResolver, norm } = require('../vendor-map');
const { openSpendSheet, cellValue, trailingAverage, readAliasMap } = require('../spend-sheet');

// Reads a statement and works out what each app's month cell should say, then
// sorts every proposal into "safe to write without asking" and "needs a human".
// Nothing is written here — this endpoint only ever reads.

const MAX_BYTES = 3 * 1024 * 1024; // base64 inflates ~33%; Vercel caps bodies near 4.5 MB
const EPSILON = 0.005;             // cents-level equality
const SWING_HIGH = 2.5;            // 2.5x the trailing average is worth a look
const SWING_LOW = 0.4;
const SWING_MIN_DELTA = 200;       // ignore swings on rows too small to matter
const MAX_TXNS_PER_PROPOSAL = 60;  // cap the drill-down payload

async function loadStatementBuffer(body, token) {
  if (body.contentBase64) {
    const buffer = Buffer.from(body.contentBase64, 'base64');
    if (!buffer.length) throw new Error('The uploaded file is empty');
    if (buffer.length > MAX_BYTES) throw new Error('Statement file is too large (max 3 MB)');
    return buffer;
  }
  if (body.sourceUrl) {
    const share = await resolveShare(token, body.sourceUrl);
    if (share.isFolder) throw new Error('That link points to a folder — link the statement file itself');
    const res = await graphFetch(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(share.driveId)}/items/${encodeURIComponent(share.itemId)}/content`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Could not download the linked statement (${res.status})`);
    return Buffer.from(await res.arrayBuffer());
  }
  throw new Error('Provide either an uploaded file or a link to the statement');
}

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const body = req.body || {};
    const attribution = body.attribution === 'transaction' ? 'transaction' : 'statement';

    const token = await getGraphToken();
    const buffer = await loadStatementBuffer(body, token);

    const { txns, sheetReports } = parseStatement(buffer);
    if (!txns.length) {
      res.status(400).json({ error: 'No transactions found in that file. Expected a sheet with "Description" and "Amount" columns.', sheetReports });
      return;
    }

    const sheet = await openSpendSheet(token);
    const appNames = sheet.grid.apps.map(a => a.name);
    const rowByApp = new Map(sheet.grid.apps.map(a => [a.name, a.rowIdx]));

    const savedAliases = await readAliasMap(token, sheet.driveId);
    // Mappings the user has just made in the review UI but not yet applied, so
    // re-previewing folds them in without saving anything.
    const overrides = (body.aliasOverrides && typeof body.aliasOverrides === 'object') ? body.aliasOverrides : {};
    const resolve = buildResolver({ ...savedAliases, ...overrides }, appNames);

    const covered = monthsCovered(txns, attribution);
    // Default to the month the statement is mostly about; the UI can narrow it.
    const monthFilter = body.month || null;

    const groups = new Map();   // "app||month" -> aggregate
    const unmapped = new Map(); // normalized label -> aggregate

    for (const txn of txns) {
      const month = monthFor(txn, attribution);
      if (!month) continue;
      if (monthFilter && month !== monthFilter) continue;

      const match = resolve(txn.vendorLabel, txn.description);
      const brief = { date: txn.date, description: txn.description, amount: txn.amount, sheet: txn.sheet, row: txn.row, label: txn.vendorLabel };

      if (!match.app) {
        const key = norm(txn.vendorLabel) || norm(txn.description).slice(0, 40);
        if (!unmapped.has(key)) {
          unmapped.set(key, { key, label: txn.vendorLabel || txn.description.slice(0, 40), suggestion: match.suggestion || null, amount: 0, months: new Set(), txns: [] });
        }
        const u = unmapped.get(key);
        u.amount += txn.amount;
        u.months.add(month);
        if (u.txns.length < MAX_TXNS_PER_PROPOSAL) u.txns.push(brief);
        continue;
      }

      const key = `${match.app}||${month}`;
      if (!groups.has(key)) groups.set(key, { app: match.app, month, amount: 0, via: match.via, labels: new Set(), txns: [], hasRefund: false });
      const g = groups.get(key);
      g.amount += txn.amount;
      if (txn.vendorLabel) g.labels.add(txn.vendorLabel);
      if (txn.isRefund) g.hasRefund = true;
      if (g.txns.length < MAX_TXNS_PER_PROPOSAL) g.txns.push(brief);
    }

    const proposals = [];
    for (const g of groups.values()) {
      const rowIdx = rowByApp.get(g.app);
      const colIdx = sheet.grid.monthCols[g.month];
      const amount = Math.round(g.amount * 100) / 100;

      if (colIdx === undefined) {
        proposals.push({ ...g, amount, labels: [...g.labels], current: null, status: 'review', reason: 'no-column', note: `The sheet has no ${g.month} column.` });
        continue;
      }

      const current = cellValue(sheet.values, rowIdx, colIdx);
      const avg = trailingAverage(sheet.values, sheet.grid, rowIdx, g.month, 3);
      const base = { ...g, labels: [...g.labels], amount, current, trailingAvg: avg == null ? null : Math.round(avg * 100) / 100, address: null };

      let status = 'auto';
      let reason = 'new';
      let note = '';

      if (amount < 0) {
        status = 'review'; reason = 'net-refund';
        note = 'Refunds exceed charges, so this month is negative.';
      } else if (current !== null && Math.abs(current - amount) < EPSILON) {
        status = 'unchanged'; reason = 'unchanged';
        note = 'The sheet already holds this figure.';
      } else if (current !== null && current !== 0) {
        status = 'review'; reason = 'overwrite';
        note = `Cell already holds ${current.toLocaleString('en-US', { minimumFractionDigits: 2 })}.`;
      } else if (avg && amount > 0 && Math.abs(amount - avg) >= SWING_MIN_DELTA && (amount > avg * SWING_HIGH || amount < avg * SWING_LOW)) {
        status = 'review'; reason = 'swing';
        note = `${amount > avg ? 'Well above' : 'Well below'} the ${Math.round(avg).toLocaleString('en-US')} trailing average.`;
      } else if (g.hasRefund) {
        status = 'review'; reason = 'has-refund';
        note = 'Includes a refund netted against the charges.';
      }

      proposals.push({ ...base, status, reason, note });
    }

    proposals.sort((a, b) => (b.amount || 0) - (a.amount || 0));

    const unmappedList = [...unmapped.values()]
      .map(u => ({ ...u, months: [...u.months], amount: Math.round(u.amount * 100) / 100 }))
      .sort((a, b) => b.amount - a.amount);

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({
      ok: true,
      attribution,
      monthsCovered: covered,
      month: monthFilter,
      sheetName: sheet.sheetName,
      sheetReports,
      appNames,
      proposals,
      unmapped: unmappedList,
      summary: {
        auto: proposals.filter(p => p.status === 'auto').length,
        review: proposals.filter(p => p.status === 'review').length,
        unchanged: proposals.filter(p => p.status === 'unchanged').length,
        unmapped: unmappedList.length,
        transactions: txns.length,
      },
    });
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
