const { getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment, listFilesRecursive, uploadFileContent, resolveArchiveRoot, graphFetch } = require('../../lib/graph');

// Graph's "resolve a sharing URL" trick: base64url-encode the URL, prefix with "u!".
// https://learn.microsoft.com/en-us/graph/api/shares-get
function encodeShareUrl(url) {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  const base64url = base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
  return `u!${base64url}`;
}

async function graphGet(token, url) {
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${res.status}: ${text.slice(0, 200)}`);
  }
  return res.json();
}

function sanitizeFileName(name) {
  const dot = name.lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot).replace(/[<>:"/\\|?*\x00-\x1F]/g, '') : '';
  return sanitizeSegment(base) + ext;
}
function sanitizeRelPath(relPath) {
  return (relPath || '').split('/').filter(Boolean).map(sanitizeSegment).join('/');
}

function normalize(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}
function tokens(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
}
// Best-effort fuzzy match, used only to pre-select a suggestion in the review UI —
// the user confirms or overrides every mapping before anything is copied.
function bestMatch(sourceFolderName, appNames) {
  const na = normalize(sourceFolderName);
  const ta = new Set(tokens(sourceFolderName));
  let best = null, bestScore = 0;
  for (const app of appNames) {
    const nb = normalize(app);
    if (!na || !nb) continue;
    let score = 0;
    if (na === nb) score = 1;
    else if (na.includes(nb) || nb.includes(na)) score = 0.8;
    else {
      const tb = new Set(tokens(app));
      const inter = [...ta].filter(x => tb.has(x)).length;
      const union = new Set([...ta, ...tb]).size;
      score = union ? inter / union : 0;
      for (const t of tb) if (t.length >= 4 && na.includes(t)) score = Math.max(score, 0.6);
      for (const t of ta) if (t.length >= 4 && nb.includes(t)) score = Math.max(score, 0.6);
    }
    if (score > bestScore) { bestScore = score; best = app; }
  }
  return { app: bestScore >= 0.4 ? best : null, score: bestScore };
}

const COMMIT_BATCH = 12; // files copied per commit request, to stay under the function timeout

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const { sourceUrl, mode, mapping, appNames, folderId, targetApp: commitApp, offset } = req.body || {};
    if (!sourceUrl) {
      res.status(400).json({ error: 'Missing sourceUrl' });
      return;
    }

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const token = await getGraphToken();

    const shareToken = encodeShareUrl(sourceUrl);
    const rootItem = await graphGet(
      token,
      `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem?$select=id,name,parentReference,folder`
    );
    if (!rootItem.folder) {
      res.status(400).json({ error: 'That link does not point to a folder' });
      return;
    }
    const sourceDriveId = rootItem.parentReference.driveId;

    // The source and the archive used to be two different folders, with this
    // copying between them. They are now one and the same, so copying a source
    // subfolder into {archive}/{app}/ would duplicate it under a second name —
    // 85 Anthropic PDFs landing in both "Claude Api" and "Anthropic(Api
    // Console)", with no easy way back. The mapping is still worth saving (the
    // mailbox sync files by it), so the review flow runs; only the copy stops.
    const targetDriveId = await resolveDriveId(token, upn);
    const archiveRoot = await resolveArchiveRoot(token, targetDriveId);
    const sourceIsArchive = sourceDriveId === targetDriveId && rootItem.id === archiveRoot.itemId;

    const subfoldersRes = await graphGet(
      token,
      `https://graph.microsoft.com/v1.0/drives/${sourceDriveId}/items/${rootItem.id}/children?$select=id,name,folder,folder&$top=200`
    );
    const appFolders = (subfoldersRes.value || []).filter(f => f.folder);

    if (mode === 'preview') {
      const proposals = [];
      for (const folder of appFolders) {
        let fileCount = 0;
        try {
          const files = await listFilesRecursive(token, sourceDriveId, folder.id);
          fileCount = files.length;
        } catch { /* leave at 0, still list the folder */ }
        const { app, score } = bestMatch(folder.name, appNames || []);
        proposals.push({ folderId: folder.id, folderName: folder.name, fileCount, suggestedApp: app, score });
      }
      res.status(200).json({ sourceDriveId, folders: proposals, sourceIsArchive, archive: archiveRoot.path });
      return;
    }

    if (mode !== 'commit') {
      res.status(400).json({ error: `Unknown mode "${mode}" — expected "preview" or "commit". If you're not sure why you're seeing this, hard-refresh the dashboard page and try again (this can happen if your browser is running an older cached version of the page).` });
      return;
    }
    // commit mode processes ONE source folder per request, in a bounded batch of
    // files starting at `offset`. The client loops (advancing offset, then moving
    // to the next folder) so no single request runs long enough to time out.
    if (!folderId || !commitApp) {
      res.status(400).json({ error: 'commit mode needs folderId and targetApp' });
      return;
    }
    const folder = appFolders.find(f => f.id === folderId);
    if (!folder) {
      res.status(400).json({ error: 'Source folder not found (it may have moved). Re-scan and try again.' });
      return;
    }

    const targetApp = sanitizeSegment(commitApp);
    const startOffset = Number(offset) || 0;

    // Already where it belongs: record it as present rather than copying it
    // beside itself under another name.
    if (sourceIsArchive) {
      let already = 0;
      try { already = (await listFilesRecursive(token, sourceDriveId, folder.id)).length; } catch { /* count is cosmetic */ }
      res.status(200).json({
        folderName: folder.name, total: already, done: already, nextOffset: null,
        copied: [], alreadyPresent: already, skippedTooLarge: [], errors: [],
        note: `"${folder.name}" is already inside the archive (${archiveRoot.path}) — nothing to copy; the mapping is saved.`,
      });
      return;
    }

    let files;
    try {
      files = await listFilesRecursive(token, sourceDriveId, folder.id);
    } catch (e) {
      res.status(200).json({ folderName: folder.name, total: 0, done: 0, nextOffset: null, copied: [], alreadyPresent: 0, skippedTooLarge: [], errors: [`${folder.name}: couldn't list files (${e.message})`] });
      return;
    }

    // Build the set of files already at the destination so we skip re-copying them —
    // makes re-runs near-instant and far less prone to transient network failures.
    let existing = new Set();
    try {
      const destId = await lookupFolderId(token, targetDriveId, `${archiveRoot.path}/${targetApp}`);
      if (destId) {
        const destFiles = await listFilesRecursive(token, targetDriveId, destId);
        existing = new Set(destFiles.map(f => `${sanitizeRelPath(f.relPath)}/${sanitizeFileName(f.name)}`));
      }
    } catch { /* if dest scan fails, fall back to copying (safe) */ }

    const batch = files.slice(startOffset, startOffset + COMMIT_BATCH);
    const summary = { folderName: folder.name, total: files.length, copied: [], alreadyPresent: 0, skippedTooLarge: [], errors: [] };

    for (const file of batch) {
      const label = file.relPath ? `${folder.name}/${file.relPath}/${file.name}` : `${folder.name}/${file.name}`;
      try {
        const destKey = `${sanitizeRelPath(file.relPath)}/${sanitizeFileName(file.name)}`;
        if (existing.has(destKey)) { summary.alreadyPresent++; continue; }
        const contentRes = await graphFetch(
          `https://graph.microsoft.com/v1.0/drives/${sourceDriveId}/items/${file.id}/content`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        if (!contentRes.ok) throw new Error(`download failed (${contentRes.status})`);
        const buf = Buffer.from(await contentRes.arrayBuffer());

        const relFolder = sanitizeRelPath(file.relPath);
        const destPath = `${archiveRoot.path}/${targetApp}${relFolder ? '/' + relFolder : ''}/${sanitizeFileName(file.name)}`;
        await uploadFileContent(token, targetDriveId, destPath, buf); // handles >4MB via upload session
        summary.copied.push(label);
      } catch (e) {
        summary.errors.push(`${label}: ${e.message}`);
      }
    }

    const done = startOffset + batch.length;
    summary.done = done;
    summary.nextOffset = done < files.length ? done : null;
    res.status(200).json(summary);
  } catch (err) {
    res.status(502).json({ error: err.message || String(err) });
  }
};

// Resolve a folder path to its item id (null if it doesn't exist yet).
async function lookupFolderId(token, driveId, path) {
  const res = await graphFetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`lookup ${res.status}`);
  return (await res.json()).id;
}
