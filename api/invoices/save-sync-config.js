const { getGraphToken, resolveDriveId, writeJsonFile, resolveArchiveRoot, archiveFile } = require('../../lib/graph');

// Persists the folder->app mapping (confirmed by the user in the import review UI)
// to _sync-config.json in the invoice archive, so the daily cron can mirror new invoices
// automatically without a human re-confirming the mapping each time.
module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const { sourceUrl, mapping } = req.body || {};
    if (!sourceUrl || !mapping || typeof mapping !== 'object' || !Object.keys(mapping).length) {
      res.status(400).json({ error: 'Missing sourceUrl or mapping' });
      return;
    }

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);
    const root = await resolveArchiveRoot(token, driveId);

    await writeJsonFile(token, driveId, archiveFile(root, '_sync-config.json'), {
      sourceUrl,
      mapping, // { [sourceFolderName]: appName }
      savedAt: new Date().toISOString(),
    });

    res.status(200).json({ ok: true, folders: Object.keys(mapping).length });
  } catch (err) {
    res.status(502).json({ error: err.message || String(err) });
  }
};
