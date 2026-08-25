const { getGraphToken, resolveDriveId, readJsonFile } = require('../graph');
const { LOG_PATH } = require('../spend-sheet');

// Recent amount writes, newest first — the visible half of the audit trail so a
// figure in the sheet can always be traced back to a statement and a person.
module.exports = async (req, res) => {
  try {
    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
    const limit = Math.min(Number(req.query && req.query.limit) || 20, 100);

    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);
    const log = await readJsonFile(token, driveId, LOG_PATH);
    const entries = (log && Array.isArray(log.entries) ? log.entries : []).slice(-limit).reverse();

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({ ok: true, entries });
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
