const { getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment, listFilesRecursive, resolveArchiveRoot, graphFetch } = require('../../lib/graph');
const { buildInventory } = require('../../lib/invoices/inventory');
const { buildDocumentIndex } = require('../../lib/documents');

// Three shapes behind one route:
//   ?app=Bubble        -> the files filed under that one app (the app modal)
//   ?mode=checklist    -> the whole archive, for the Invoices tab
//   ?mode=documents    -> the agreements/KYC folder, for the Agreements page
// They share a function because each file under api/ counts against the Hobby
// plan's 12-Serverless-Function limit and the deployment is already at it.

// A full-archive crawl is ~one Graph listing per app folder, so it is cached
// briefly. Serverless instances are ephemeral, but this still absorbs the
// repeat loads from opening the tab and from several people on the dashboard.
const INVENTORY_TTL_MS = 5 * 60 * 1000;
let inventoryCache = { data: null, expiresAt: 0 };

async function checklist(req, res) {
  const forceRefresh = req.query && (req.query.refresh === '1' || req.query.refresh === 'true');
  const now = Date.now();
  if (!forceRefresh && inventoryCache.data && now < inventoryCache.expiresAt) {
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Cache', 'HIT');
    res.status(200).json(inventoryCache.data);
    return;
  }
  const upn = (process.env.TARGET_USER_UPN || '').trim();
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
  const token = await getGraphToken();
  const driveId = await resolveDriveId(token, upn);
  const data = await buildInventory(token, driveId, { deadline: now + 40 * 1000, fresh: forceRefresh });
  // A crawl cut short by the deadline is a partial picture; caching it would
  // make the gaps look real for the next five minutes.
  if (!data.truncated) inventoryCache = { data, expiresAt: now + INVENTORY_TTL_MS };
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Cache', 'MISS');
  res.status(200).json(data);
}

// The agreements/KYC drawer. Same caching shape as the checklist: one Graph
// listing per vendor folder, cheap but not free, and several people opening the
// page should not each pay for it.
const DOCUMENTS_TTL_MS = 5 * 60 * 1000;
let documentsCache = { data: null, expiresAt: 0 };

async function documents(req, res) {
  const forceRefresh = req.query && (req.query.refresh === '1' || req.query.refresh === 'true');
  const now = Date.now();
  if (!forceRefresh && documentsCache.data && now < documentsCache.expiresAt) {
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Cache', 'HIT');
    res.status(200).json(documentsCache.data);
    return;
  }
  const upn = (process.env.TARGET_USER_UPN || '').trim();
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
  const token = await getGraphToken();
  const driveId = await resolveDriveId(token, upn);
  const data = await buildDocumentIndex(token, driveId, { deadline: now + 40 * 1000, fresh: forceRefresh });
  // A listing cut short by the deadline is a partial drawer; caching it would
  // make the missing vendors look like vendors with nothing on file.
  if (!data.truncated) documentsCache = { data, expiresAt: now + DOCUMENTS_TTL_MS };
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Cache', 'MISS');
  res.status(200).json(data);
}

async function perApp(req, res, appName) {
  const upn = (process.env.TARGET_USER_UPN || '').trim();
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

  const token = await getGraphToken();
  const driveId = await resolveDriveId(token, upn);
  const root = await resolveArchiveRoot(token, driveId);
  const folder = `${root.path}/${sanitizeSegment(appName)}`;
  const folderUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(folder)}?$select=id`;

  const folderRes = await graphFetch(folderUrl, { headers: { Authorization: `Bearer ${token}` } });
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
}

module.exports = async (req, res) => {
  try {
    const mode = String((req.query && req.query.mode) || '').trim();
    if (mode === 'checklist') return await checklist(req, res);
    if (mode === 'documents') return await documents(req, res);

    const appName = req.query && req.query.app;
    if (!appName) {
      res.status(400).json({ error: 'Missing app query param' });
      return;
    }
    return await perApp(req, res, appName);
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
