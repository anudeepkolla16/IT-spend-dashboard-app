const { getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment } = require('../../lib/graph');

// Graph's "resolve a sharing URL" trick: base64url-encode the URL, prefix with "u!".
// https://learn.microsoft.com/en-us/graph/api/shares-get
function encodeShareUrl(url) {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  const base64url = base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
  return `u!${base64url}`;
}

async function graphGet(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
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

const MAX_IMPORT_BYTES = 4 * 1024 * 1024; // Graph simple-upload cap

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const { sourceUrl, mode, mapping, appNames } = req.body || {};
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
          const filesRes = await graphGet(
            token,
            `https://graph.microsoft.com/v1.0/drives/${sourceDriveId}/items/${folder.id}/children?$select=id,file&$top=200`
          );
          fileCount = (filesRes.value || []).filter(f => f.file).length;
        } catch { /* leave at 0, still list the folder */ }
        const { app, score } = bestMatch(folder.name, appNames || []);
        proposals.push({ folderId: folder.id, folderName: folder.name, fileCount, suggestedApp: app, score });
      }
      res.status(200).json({ sourceDriveId, folders: proposals });
      return;
    }

    // commit mode: mapping is { [folderId]: targetAppName } — only folders present
    // in the mapping (and with a non-empty target) get copied.
    if (!mapping || typeof mapping !== 'object') {
      res.status(400).json({ error: 'Missing mapping for commit mode' });
      return;
    }
    const targetDriveId = await resolveDriveId(token, upn);
    const summary = { copied: [], skippedTooLarge: [], errors: [] };

    for (const folder of appFolders) {
      const targetAppRaw = mapping[folder.id];
      if (!targetAppRaw) continue; // user chose to skip this folder
      const targetApp = sanitizeSegment(targetAppRaw);

      let filesRes;
      try {
        filesRes = await graphGet(
          token,
          `https://graph.microsoft.com/v1.0/drives/${sourceDriveId}/items/${folder.id}/children?$select=id,name,file,size&$top=200`
        );
      } catch (e) {
        summary.errors.push(`${folder.name}: couldn't list files (${e.message})`);
        continue;
      }

      for (const file of (filesRes.value || []).filter(f => f.file)) {
        try {
          if (file.size > MAX_IMPORT_BYTES) {
            summary.skippedTooLarge.push(`${folder.name}/${file.name}`);
            continue;
          }
          const contentRes = await fetch(
            `https://graph.microsoft.com/v1.0/drives/${sourceDriveId}/items/${file.id}/content`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (!contentRes.ok) throw new Error(`download failed (${contentRes.status})`);
          const buf = Buffer.from(await contentRes.arrayBuffer());

          const destPath = `Invoices/${targetApp}/${sanitizeFileName(file.name)}`;
          const putUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(targetDriveId)}/root:/${encodeGraphPath(destPath)}:/content`;
          const putRes = await fetch(putUrl, {
            method: 'PUT',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/octet-stream' },
            body: buf,
          });
          if (!putRes.ok) {
            const text = await putRes.text();
            throw new Error(`upload failed (${putRes.status}): ${text.slice(0, 150)}`);
          }
          summary.copied.push(`${folder.name} → ${targetAppRaw} / ${file.name}`);
        } catch (e) {
          summary.errors.push(`${folder.name}/${file.name}: ${e.message}`);
        }
      }
    }

    res.status(200).json(summary);
  } catch (err) {
    res.status(502).json({ error: err.message || String(err) });
  }
};
