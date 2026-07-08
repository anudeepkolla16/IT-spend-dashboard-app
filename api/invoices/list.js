const { getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment, listFilesRecursive } = require('../../lib/graph');

module.exports = async (req, res) => {
  try {
    const appName = req.query && req.query.app;
    if (!appName) {
      res.status(400).json({ error: 'Missing app query param' });
      return;
    }
    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const folder = `Invoices/${sanitizeSegment(appName)}`;
    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);
    const folderUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(folder)}?$select=id`;

    const folderRes = await fetch(folderUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (folderRes.status === 404) {
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ files: [] });
      return;
    }
    if (!folderRes.ok) {
      const text = await folderRes.text();
      throw new Error(`Graph folder lookup failed (${folderRes.status}): ${text.slice(0, 300)}`);
    }
    const folderItem = await folderRes.json();

    const rawFiles = await listFilesRecursive(token, driveId, folderItem.id);
    const files = rawFiles
      .map(f => ({ name: f.name, size: f.size, uploadedAt: f.createdDateTime, webUrl: f.webUrl, subfolder: f.relPath || '' }))
      .sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({ files });
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
