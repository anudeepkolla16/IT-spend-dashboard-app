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
  const res = await graphFetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
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

// Graph throttles, and a throttled call is not a failure — it is a call to make
// again. 429 and the 5xx family are the two things a run of a few hundred Graph
// calls sees routinely, and every one of them used to land in the run summary as
// an error beside the invoices it lost. Retried here instead, honouring the
// Retry-After Graph sends, and giving up rather than blowing the function's
// 45-second budget: the caller then reports a real failure, not a hiccup.
const RETRY_STATUS = new Set([429, 502, 503, 504]);
const MAX_ATTEMPTS = 4;
const MAX_BACKOFF_MS = 8000;

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Graph's Retry-After is seconds (occasionally an HTTP date). It can say 300,
// which is longer than the whole run — capped, and the deadline check below
// decides whether even the capped wait is affordable.
function retryAfterMs(res, attempt) {
  const raw = res && res.headers && typeof res.headers.get === 'function' ? res.headers.get('retry-after') : null;
  if (raw) {
    const secs = Number(raw);
    if (Number.isFinite(secs) && secs >= 0) return Math.min(secs * 1000, MAX_BACKOFF_MS);
    const when = Date.parse(raw);
    if (!Number.isNaN(when)) return Math.min(Math.max(when - Date.now(), 0), MAX_BACKOFF_MS);
  }
  // 400ms, 800ms, 1600ms … plus jitter, so parallel callers do not all come
  // back at the same instant and throttle each other again.
  return Math.min(400 * Math.pow(2, attempt - 1), MAX_BACKOFF_MS) + Math.floor(Math.random() * 250);
}

// fetch, with retries for throttling and transient network faults. Returns the
// Response for the caller to check — a 404 or a 403 is an answer, not a fault,
// and only the caller knows which of those it can live with.
async function graphFetch(url, init, opts) {
  const options = opts || {};
  const attempts = options.attempts || MAX_ATTEMPTS;
  let lastRes = null;
  let networkError = null;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      lastRes = await fetch(url, init);
      networkError = null;
    } catch (e) {
      lastRes = null;
      networkError = e; // socket reset, DNS blip — worth one more go
    }
    if (lastRes && !RETRY_STATUS.has(lastRes.status)) return lastRes;
    if (attempt === attempts) break;

    const wait = retryAfterMs(lastRes, attempt);
    // Never spend the run's remaining time sleeping: hand back what we have and
    // let the caller report it, so the next run picks the work up.
    if (options.deadline && Date.now() + wait > options.deadline) break;
    await sleep(wait);
  }

  // Out of attempts. A response that is still a throttle or a 5xx is handed
  // back for the caller to report; a network fault has no response to hand back.
  if (lastRes) return lastRes;
  throw networkError;
}

// Follows @odata.nextLink so a folder with more children than one page is read
// whole. A truncated listing is worse than a failed one: it reads as "these are
// all the invoices there are", and the totals built on it are quietly short.
async function graphListAll(token, url, describe) {
  const out = [];
  let next = url;
  let pages = 0;
  while (next && pages < 20) {
    const res = await graphFetch(next, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`${describe || 'Graph listing'} failed (${res.status}): ${text.slice(0, 200)}`);
    }
    const json = await res.json();
    for (const item of json.value || []) out.push(item);
    next = json['@odata.nextLink'] || null;
    pages++;
  }
  return out;
}

async function resolveDriveId(token, upn) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/drive`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
  const children = await graphListAll(token, url, 'Graph folder listing');
  let files = [];
  for (const child of children) {
    if (child.file) {
      files.push({ id: child.id, name: child.name, size: child.size, createdDateTime: child.createdDateTime, webUrl: child.webUrl, relPath: relPath || '' });
    } else if (child.folder) {
      const nested = await listFilesRecursive(token, driveId, child.id, relPath ? `${relPath}/${child.name}` : child.name, depth + 1);
      files = files.concat(nested);
    }
  }
  return files;
}

// Resolve a path to its item id, or null when nothing is there.
async function itemIdByPath(token, driveId, path) {
  const res = await graphFetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Graph lookup of "${path}" failed (${res.status})`);
  return (await res.json()).id;
}

// The id of a folder, creating it (and any missing parent) if it is not there.
// Uploading to a path creates the folders on the way implicitly; moving an item
// does not, so a move into a month folder that has never existed needs this.
async function ensureFolder(token, driveId, path) {
  const clean = String(path || '').replace(/^\/+|\/+$/g, '');
  if (!clean) throw new Error('ensureFolder needs a path');

  const existing = await itemIdByPath(token, driveId, clean);
  if (existing) return existing;

  const cut = clean.lastIndexOf('/');
  const parentPath = cut === -1 ? '' : clean.slice(0, cut);
  const name = cut === -1 ? clean : clean.slice(cut + 1);
  const parentId = parentPath ? await ensureFolder(token, driveId, parentPath) : null;
  const url = parentId
    ? `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`
    : `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root/children`;

  const res = await graphFetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
  });
  // 409: something created it between the lookup and now — take what is there.
  if (res.status === 409) {
    const raced = await itemIdByPath(token, driveId, clean);
    if (raced) return raced;
  }
  if (!res.ok) throw new Error(`could not create folder "${clean}" (${res.status})`);
  return (await res.json()).id;
}

// Move an item into another folder. A move keeps the file's identity, its
// version history and any link anyone has to it — which a download, re-upload
// and delete would not.
async function moveItem(token, driveId, itemId, newParentId, newName) {
  const body = { parentReference: { id: newParentId } };
  if (newName) body.name = newName;
  const res = await graphFetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
    {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    }
  );
  if (res.status === 409) throw new Error('a file of that name is already in the destination');
  if (!res.ok) throw new Error(`move failed (${res.status}): ${(await res.text()).slice(0, 150)}`);
  return res.json();
}

// Graph "resolve a sharing URL" encoding: base64url of the URL, prefixed with "u!".
function encodeShareUrl(url) {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  return 'u!' + base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
}

// Resolve a share link to { driveId, itemId, name, isFolder }.
async function resolveShare(token, sourceUrl) {
  const url = `https://graph.microsoft.com/v1.0/shares/${encodeShareUrl(sourceUrl)}/driveItem?$select=id,name,parentReference,folder`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
  const res = await graphFetch(url, {
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

// Upload file bytes to a drive path. Small files use a simple PUT; files over
// ~4 MB use a resumable upload session (Graph rejects simple PUT past that),
// chunked at 5 MiB (a multiple of 320 KiB, as Graph requires).
const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;
async function uploadFileContent(token, driveId, path, buf, contentType) {
  const base = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}`;
  if (buf.length <= SIMPLE_UPLOAD_MAX) {
    const res = await graphFetch(`${base}:/content`, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType || 'application/octet-stream' },
      body: buf,
    });
    if (!res.ok) throw new Error(`upload failed (${res.status}): ${(await res.text()).slice(0, 150)}`);
    return;
  }

  const sessRes = await graphFetch(`${base}:/createUploadSession`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
  });
  if (!sessRes.ok) throw new Error(`createUploadSession failed (${sessRes.status}): ${(await sessRes.text()).slice(0, 150)}`);
  const { uploadUrl } = await sessRes.json();

  const CHUNK = 5 * 1024 * 1024; // 5 MiB = 320 KiB * 16
  for (let start = 0; start < buf.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, buf.length);
    const chunk = buf.subarray(start, end);
    // The uploadUrl is pre-authenticated — do NOT send the Authorization header.
    const put = await graphFetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(chunk.length),
        'Content-Range': `bytes ${start}-${end - 1}/${buf.length}`,
      },
      body: chunk,
    });
    if (!put.ok && put.status !== 202) {
      throw new Error(`chunk upload failed (${put.status}): ${(await put.text()).slice(0, 150)}`);
    }
  }
}

module.exports = {
  getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment, listFilesRecursive,
  encodeShareUrl, resolveShare, readJsonFile, writeJsonFile, uploadFileContent,
  graphFetch, graphListAll, itemIdByPath, ensureFolder, moveItem,
};
