const { getGraphToken, resolveDriveId } = require('../graph');
const { readLog } = require('../spend-sheet');
const { readActivity } = require('../activity');

// The admin's view of who has used the dashboard and what they did.
//
// It reads two files and merges them, rather than keeping a third that
// restates what they hold:
//
//   _activity-log.json  sign-ins, mailbox syncs, archive tidies, imports —
//                       the things that leave no cell behind
//   _amount-log.json    every cell this app has written, already kept by the
//                       sheet writer with who did it and what was there before
//
// A write in the second file becomes one line per cell, because "changed AWS
// September from 11.00 to 2,401.67" is the sentence someone is looking for;
// "a write touching 14 cells" is not.
//
// The route this hangs off is editor-only, so nothing here checks again —
// api/amounts.js gates before it dispatches.

// A write's own description of itself. `by` is an email for a person, or the
// name of the job for the sync's own writes.
const WHO = {
  'invoice-sync': 'the invoice sync',
  'period-backfill': 'the billing-period backfill',
  'fix-totals': 'the totals repair',
};

const money = (n) => n == null
  ? 'nothing'
  : Number(n).toLocaleString('en-US', { style: 'currency', currency: 'USD' });

function cellLines(logEntry) {
  const by = String(logEntry.by || '').trim();
  const who = WHO[by] || by || 'unknown';
  const isPerson = who.includes('@');
  const source = logEntry.attribution || logEntry.source || '';
  return (logEntry.cells || []).map((c) => ({
    at: logEntry.at,
    kind: isPerson ? 'edit' : 'sync-write',
    by: who,
    what: `${c.app} ${c.month}: ${money(c.before)} → ${money(c.after)}`,
    detail: [
      c.address ? `cell ${c.address}` : '',
      logEntry.sheet ? `on ${logEntry.sheet}` : '',
      source ? `from ${source}` : '',
    ].filter(Boolean).join(' · '),
  }));
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  try {
    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);

    // Either file may be missing on a fresh archive; neither is a failure.
    const [events, writeLog] = await Promise.all([
      readActivity(token, driveId),
      readLog(token, driveId).catch(() => null),
    ]);

    const fromWrites = (writeLog && Array.isArray(writeLog.entries) ? writeLog.entries : [])
      .flatMap(cellLines);

    const merged = events.concat(fromWrites)
      .filter((e) => e && e.at)
      .sort((a, b) => String(b.at).localeCompare(String(a.at)));

    // Who has been in here, newest first, so the page can show it without
    // walking the feed itself.
    const seen = new Map();
    for (const e of merged) {
      const by = String(e.by || '').trim();
      if (!by || !by.includes('@')) continue;       // a job is not a person
      if (!seen.has(by)) seen.set(by, { by, lastAt: e.at, events: 0, signIns: 0 });
      const row = seen.get(by);
      row.events++;
      if (e.kind === 'sign-in') row.signIns++;
    }

    res.status(200).json({
      entries: merged.slice(0, 400),
      people: [...seen.values()],
      counts: { total: merged.length, events: events.length, writes: fromWrites.length },
    });
  } catch (err) {
    res.status(500).json({ error: (err && err.message) || String(err) });
  }
};
