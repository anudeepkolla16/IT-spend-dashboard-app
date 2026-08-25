const {
  getGraphToken, resolveDriveId, resolveShare, listFilesRecursive,
  encodeGraphPath, sanitizeSegment, readJsonFile, uploadFileContent,
} = require('../../lib/graph');
const { verify, parseCookies } = require('../../lib/session');
const { runMailSync } = require('../../lib/mail-sync');

// The daily invoice job. Two sources, one function:
//   1. mirrors new invoice PDFs from the mapped SharePoint source folders
//   2. files invoices out of the shared invoices mailbox
// They share one route (and one timeout) because each file under api/ counts
// against the Hobby plan's 12-Serverless-Function limit for a deployment.

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

// Mirror new files from the mapped source folders into Invoices/{App}/.
async function mirrorSourceFolders(token, targetDriveId, deadline) {
  const config = await readJsonFile(token, targetDriveId, 'Invoices/_sync-config.json');
  if (!config || !config.sourceUrl || !config.mapping) {
    return { note: 'No sync config saved yet — run an import from the dashboard once to set it up.' };
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
  const mapped = subfolders.filter(f => config.mapping[f.name]);

  // Process one source folder: scan destination + source, copy anything new.
  async function processFolder(folder) {
    if (Date.now() > deadline) { summary.timedOut = true; return; }
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
      if (Date.now() > deadline) { summary.timedOut = true; break; }
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
    while (cursor < mapped.length && Date.now() <= deadline) {
      await processFolder(mapped[cursor++]);
    }
  }));

  return summary;
}

module.exports = async (req, res) => {
  try {
    // Vercel automatically sends `Authorization: Bearer <CRON_SECRET>` for
    // scheduled invocations. The dashboard's "Fetch Invoices" button calls the
    // same job with a signed-in session instead, so accept either.
    const secret = process.env.CRON_SECRET;
    const auth = req.headers['authorization'] || '';
    const isCron = !!secret && auth === `Bearer ${secret}`;
    const isUser = !!verify(parseCookies(req.headers.cookie).session);
    if (!isCron && !isUser) {
      res.status(401).json({ error: 'Unauthorized' });
      return;
    }

    // ?mode=mail or ?mode=folders runs just one half; the cron runs both.
    const mode = String((req.query && req.query.mode) || 'all').trim();

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const token = await getGraphToken();
    const targetDriveId = await resolveDriveId(token, upn);

    // Time-box the run so we always return before the function's hard limit.
    // Because only NEW files are copied (existing ones skipped), a run that
    // stops early is simply finished by the next one — the mirror self-converges.
    // Both halves share the budget, so neither can starve the other completely.
    const started = Date.now();
    const HARD_DEADLINE = started + 45 * 1000;
    const out = { ok: true, ranAt: new Date().toISOString(), mode };

    if (mode !== 'mail') {
      // Leave at least a third of the budget for the mailbox pass.
      const folderDeadline = mode === 'folders' ? HARD_DEADLINE : started + 28 * 1000;
      try {
        out.folders = await mirrorSourceFolders(token, targetDriveId, folderDeadline);
      } catch (e) {
        out.folders = { error: e.message };
      }
    }

    if (mode !== 'folders') {
      // A mailbox failure (most likely a missing Mail.Read grant) must not take
      // the folder mirror down with it.
      try {
        out.mail = await runMailSync(token, targetDriveId, { deadline: HARD_DEADLINE });
      } catch (e) {
        out.mail = { error: e.message };
      }
    }

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json(out);
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
