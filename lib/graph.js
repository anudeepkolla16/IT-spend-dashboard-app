async function getGraphToken() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET env vars');
  }
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
  });
  const res = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph token request failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const json = await res.json();
  return json.access_token;
}

async function resolveDriveId(token, upn) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/drive`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph drive lookup failed (${res.status}) for upn "${upn}": ${text.slice(0, 300)}`);
  }
  const json = await res.json();
  return json.id;
}

function encodeGraphPath(path) {
  return path.split('/').map(encodeURIComponent).join('/');
}

// Keep folder/file names safe for SharePoint and confined to their own scope
// (no path traversal via "..", no characters SharePoint disallows).
function sanitizeSegment(name) {
  return String(name || '')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '-')
    .replace(/\.\.+/g, '-')
    .trim()
    .slice(0, 150) || 'Unnamed';
}

// Recursively collect files under a folder (Graph item id), following subfolders
// (e.g. invoices organized as AppFolder/Month/file.pdf). relPath tracks the
// subfolder chain relative to the starting folder, for preserving structure.
async function listFilesRecursive(token, driveId, folderId, relPath, depth) {
  depth = depth || 0;
  if (depth > 4) return [];
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}/children?$select=id,name,file,folder,size,createdDateTime,webUrl&$top=200`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph folder listing failed (${res.status}): ${text.slice(0, 200)}`);
  }
  const json = await res.json();
  let files = [];
  for (const child of json.value || []) {
    if (child.file) {
      files.push({ id: child.id, name: child.name, size: child.size, createdDateTime: child.createdDateTime, webUrl: child.webUrl, relPath: relPath || '' });
    } else if (child.folder) {
      const nested = await listFilesRecursive(token, driveId, child.id, relPath ? `${relPath}/${child.name}` : child.name, depth + 1);
      files = files.concat(nested);
    }
  }
  return files;
}

// Graph "resolve a sharing URL" encoding: base64url of the URL, prefixed with "u!".
function encodeShareUrl(url) {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  return 'u!' + base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
}

// Resolve a share link to { driveId, itemId, name, isFolder }.
async function resolveShare(token, sourceUrl) {
  const url = `https://graph.microsoft.com/v1.0/shares/${encodeShareUrl(sourceUrl)}/driveItem?$select=id,name,parentReference,folder`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph share resolve failed (${res.status}): ${text.slice(0, 200)}`);
  }
  const json = await res.json();
  return { driveId: json.parentReference && json.parentReference.driveId, itemId: json.id, name: json.name, isFolder: !!json.folder };
}

// Read a small JSON file by path from a drive; returns null if it doesn't exist.
async function readJsonFile(token, driveId, path) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/content`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph read "${path}" failed (${res.status}): ${text.slice(0, 200)}`);
  }
  return res.json();
}

// Write (overwrite) a small JSON file by path.
async function writeJsonFile(token, driveId, path, obj) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/content`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(obj, null, 2),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph write "${path}" failed (${res.status}): ${text.slice(0, 200)}`);
  }
  return res.json();
}

module.exports = {
  getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment, listFilesRecursive,
  encodeShareUrl, resolveShare, readJsonFile, writeJsonFile,
};
