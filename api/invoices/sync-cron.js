const {
  getGraphToken, resolveDriveId, resolveShare, listFilesRecursive,
  encodeGraphPath, sanitizeSegment, readJsonFile, uploadFileContent,
} = require('../../lib/graph');

function sanitizeFileName(name) {
  const dot = name.lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot).replace(/[<>:"/\\|?*\x00-\x1F]/g, '') : '';
  return sanitizeSegment(base) + ext;
}
function sanitizeRelPath(relPath) {
  return (relPath || '').split('/').filter(Boolean).map(sanitizeSegment).join('/');
}
const keyOf = (relPath, name) => `${sanitizeRelPath(relPath)}/${sanitizeFileName(name)}`;

module.exports = async (req, res) => {
  try {
    // Auth: Vercel automatically sends `Authorization: Bearer <CRON_SECRET>` for
    // scheduled invocations when CRON_SECRET is set. Reject anything else.
    const secret = process.env.CRON_SECRET;
    const auth = req.headers['authorization'] || '';
    if (!secret || auth !== `Bearer ${secret}`) {
      res.status(401).json({ error: 'Unauthorized' });
      return;
    }

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const token = await getGraphToken();
    const targetDriveId = await resolveDriveId(token, upn);

    const config = await readJsonFile(token, targetDriveId, 'Invoices/_sync-config.json');
    if (!config || !config.sourceUrl || !config.mapping) {
      res.status(200).json({ ok: true, note: 'No sync config saved yet — run an import from the dashboard once to set it up.' });
      return;
    }

    const share = await resolveShare(token, config.sourceUrl);
    if (!share.isFolder) throw new Error('Saved sourceUrl no longer points to a folder');

    // List source subfolders (one per app/vendor).
    const subRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(share.driveId)}/items/${share.itemId}/children?$select=id,name,folder&$top=200`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!subRes.ok) throw new Error(`Source listing failed (${subRes.status})`);
    const subfolders = ((await subRes.json()).value || []).filter(f => f.folder);

    const summary = { copiedNew: 0, alreadyPresent: 0, errors: [], perApp: {}, timedOut: false };

    // Time-box the run so we always return before the function's hard limit. Because
    // only NEW files are copied (existing ones skipped), a run that stops early is
    // simply finished by the next daily run — the mirror self-converges.
    const DEADLINE = Date.now() + 45 * 1000;

    const mapped = subfolders.filter(f => config.mapping[f.name]);

    // Process one source folder: scan destination + source, copy anything new.
    async function processFolder(folder) {
      if (Date.now() > DEADLINE) { summary.timedOut = true; return; }
      const targetAppRaw = config.mapping[folder.name];
      const targetApp = sanitizeSegment(targetAppRaw);

      let existing = new Set();
      try {
        const destFolder = await readDestFolderId(token, targetDriveId, `Invoices/${targetApp}`);
        if (destFolder) {
          const destFiles = await listFilesRecursive(token, targetDriveId, destFolder);
          existing = new Set(destFiles.map(f => keyOf(f.relPath, f.name)));
        }
      } catch (e) {
        summary.errors.push(`${folder.name}: dest scan failed (${e.message})`);
      }

      let srcFiles;
      try {
        srcFiles = await listFilesRecursive(token, share.driveId, folder.id);
      } catch (e) {
        summary.errors.push(`${folder.name}: source scan failed (${e.message})`);
        return;
      }

      for (const file of srcFiles) {
        const k = keyOf(file.relPath, file.name);
        if (existing.has(k)) { summary.alreadyPresent++; continue; }
        if (Date.now() > DEADLINE) { summary.timedOut = true; break; }
        try {
          const contentRes = await fetch(
            `https://graph.microsoft.com/v1.0/drives/${share.driveId}/items/${file.id}/content`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (!contentRes.ok) throw new Error(`download ${contentRes.status}`);
          const buf = Buffer.from(await contentRes.arrayBuffer());

          const relFolder = sanitizeRelPath(file.relPath);
          const destPath = `Invoices/${targetApp}${relFolder ? '/' + relFolder : ''}/${sanitizeFileName(file.name)}`;
          await uploadFileContent(token, targetDriveId, destPath, buf); // handles >4MB via upload session
          summary.copiedNew++;
          summary.perApp[targetAppRaw] = (summary.perApp[targetAppRaw] || 0) + 1;
        } catch (e) {
          summary.errors.push(`${folder.name}/${file.name}: ${e.message}`);
        }
      }
    }

    // Run folders through a small concurrency pool so the many Graph listings happen
    // in parallel (cuts wall-clock ~4x) without hammering Graph's rate limits.
    let cursor = 0;
    const POOL = 4;
    await Promise.all(Array.from({ length: Math.min(POOL, mapped.length) }, async () => {
      while (cursor < mapped.length && Date.now() <= DEADLINE) {
        await processFolder(mapped[cursor++]);
      }
    }));

    res.status(200).json({ ok: true, ranAt: new Date().toISOString(), ...summary });
  } catch (err) {
    res.status(502).json({ error: err.message || String(err) });
  }
};

// Resolve a folder path to its item id (null if it doesn't exist yet).
async function readDestFolderId(token, driveId, path) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`lookup ${res.status}`);
  return (await res.json()).id;
}
