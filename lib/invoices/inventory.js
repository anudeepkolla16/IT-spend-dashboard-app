const { listFilesRecursive, readJsonFile, resolveArchiveRoot, archiveFile, graphListAll } = require('../graph');

// Builds the invoice checklist: what is actually sitting in the archive, per app
// and per month. The dashboard joins this against the spend rows it already has,
// so a month with a charge but no PDF shows up as a gap.
//
// The archive read here is {archive}/{App}/{Month}/, where {archive} is located
// by resolveArchiveRoot rather than hardcoded — see the long note in lib/graph.js
// for why guessing that path is how this tab came to report every month as
// missing an invoice.

const MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'];

const pad = n => String(n).padStart(2, '0');

// "aug" / "August" -> 7. Requires at least three letters, so the abbreviations
// stay unambiguous ("ma" could be March or May; "mar" and "may" cannot).
function monthIndexFromWord(word) {
  const w = String(word || '').toLowerCase();
  if (w.length < 3) return -1;
  return MONTH_NAMES.findIndex(m => m === w || m.startsWith(w));
}

// A month-only folder ("Aug", "July") has no year in it. Take the year from when
// the file landed, but roll back when the folder names a month well ahead of
// that date — a December invoice filed in January belongs to the year before.
function yearForBareMonth(monthIdx, refIso) {
  const ref = refIso ? new Date(refIso) : new Date();
  const y = Number.isNaN(ref.getTime()) ? new Date().getUTCFullYear() : ref.getUTCFullYear();
  const refMonth = Number.isNaN(ref.getTime()) ? new Date().getUTCMonth() : ref.getUTCMonth();
  return monthIdx > refMonth + 1 ? y - 1 : y;
}

// The month subfolders in the archive were named by hand over time and are not
// consistent — "Aug" under one vendor, "July" under another, "Aug-26" elsewhere.
// Accept every shape rather than only the one the sync writes.
function parseMonthFolder(name, refIso) {
  const s = String(name || '').trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})[-_. ](\d{1,2})$/);
  if (m && +m[2] >= 1 && +m[2] <= 12) return `${m[1]}-${pad(+m[2])}`;

  m = s.match(/^(\d{1,2})[-_. ](\d{4})$/);
  if (m && +m[1] >= 1 && +m[1] <= 12) return `${m[2]}-${pad(+m[1])}`;

  m = s.match(/^([A-Za-z]{3,9})[-_. ]?(\d{2}|\d{4})$/);
  if (m) {
    const idx = monthIndexFromWord(m[1]);
    if (idx >= 0) {
      const y = m[2].length === 2 ? 2000 + Number(m[2]) : Number(m[2]);
      return `${y}-${pad(idx + 1)}`;
    }
  }

  m = s.match(/^([A-Za-z]{3,9})$/);
  if (m) {
    const idx = monthIndexFromWord(m[1]);
    if (idx >= 0) return `${yearForBareMonth(idx, refIso)}-${pad(idx + 1)}`;
  }
  return null;
}

// Invoices are filed as {App}/{Month}/file.pdf, but some sit a level deeper.
// Walk the relative path outwards-in and take the first segment that reads as a
// month, so "Bubble/2026/Aug/x.pdf" resolves as well as "Bubble/Aug/x.pdf".
function monthFromRelPath(relPath, refIso) {
  for (const seg of String(relPath || '').split('/').filter(Boolean)) {
    const ym = parseMonthFolder(seg, refIso);
    if (ym) return ym;
  }
  return null;
}

// Amounts were parsed against the procurement folder's paths, so they cannot be
// joined to the mirror by path. Join by file name instead — the mirror copies the
// name through unchanged. A name that carries two different amounts is dropped
// rather than guessed at, because attributing the wrong total is worse than
// showing none.
function amountsByFileName(index) {
  const seen = new Map();
  const add = (path, amount, currency, usable) => {
    const name = String(path || '').split('/').pop();
    if (!name || amount === null || amount === undefined) return;
    const prior = seen.get(name.toLowerCase());
    if (prior === null) return;                       // already known ambiguous
    if (prior && (prior.amount !== amount || prior.currency !== currency)) {
      seen.set(name.toLowerCase(), null);
      return;
    }
    seen.set(name.toLowerCase(), { amount, currency: currency || null, usable: !!usable });
  };

  for (const e of (Array.isArray(index && index.amounts) ? index.amounts : [])) {
    if (e && e.path) add(e.path, e.amount, e.currency, e.usable);
  }
  for (const e of (Array.isArray(index && index.entries) ? index.entries : [])) {
    if (e && e.file) add(e.file, e.amount, e.currency, e.amountUsable);
  }

  const out = new Map();
  for (const [k, v] of seen) if (v) out.set(k, v);
  return out;
}

async function appFolders(token, driveId, rootId, rootPath) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(rootId)}/children?$select=id,name,folder&$top=200`;
  // Paged: the archive has one folder per application and keeps growing, and a
  // listing cut off at 200 would report every app past it as having no invoices.
  const children = await graphListAll(token, url, `Graph listing of "${rootPath}"`);
  // The sync's own bookkeeping files (_invoice-index.json, _sync-config.json)
  // and its unmatched bucket are not applications.
  return children.filter(c => c.folder && !String(c.name).startsWith('_'));
}

// Every invoice in the archive, flattened, with the month it belongs to.
async function buildInventory(token, driveId, opts) {
  const options = opts || {};
  const deadline = options.deadline || (Date.now() + 40 * 1000);
  const pool = options.pool || 6;

  const root = await resolveArchiveRoot(token, driveId, { fresh: options.fresh });
  // An unresolved archive is reported, never rendered as "no invoices on file" —
  // an empty-looking checklist and a checklist of a folder that isn't there look
  // identical on screen, and only one of them is the truth.
  if (!root.resolved) {
    return {
      scannedAt: new Date().toISOString(), files: [], apps: [], months: [],
      truncated: false, empty: true, root: null, triedPaths: root.candidates,
      fileCount: 0, undatedCount: 0, errors: [],
    };
  }

  const folders = await appFolders(token, driveId, root.itemId, root.path);
  const index = await readJsonFile(token, driveId, archiveFile(root, '_invoice-index.json')).catch(() => null);
  const amounts = amountsByFileName(index);

  // The archive's folders are vendor names; the spend sheet's rows are app names,
  // and the two only sometimes agree ("Bubble" vs "Bubble Starter", "Claude Api"
  // vs "Anthropic(Api Console)"). The import saves the folder->app mapping that
  // settles the rest, so pass it on for the checklist to join through.
  const config = await readJsonFile(token, driveId, archiveFile(root, '_sync-config.json')).catch(() => null);
  const mapping = (config && config.mapping && typeof config.mapping === 'object') ? config.mapping : {};

  const files = [];
  const errors = [];
  let truncated = false;
  let cursor = 0;

  await Promise.all(Array.from({ length: Math.min(pool, folders.length) }, async () => {
    while (cursor < folders.length) {
      if (Date.now() > deadline) { truncated = true; return; }
      const folder = folders[cursor++];
      try {
        for (const f of await listFilesRecursive(token, driveId, folder.id)) {
          const known = amounts.get(String(f.name).toLowerCase());
          files.push({
            app: folder.name,
            month: monthFromRelPath(f.relPath, f.createdDateTime),
            name: f.name,
            subfolder: f.relPath || '',
            size: f.size || 0,
            uploadedAt: f.createdDateTime || null,
            webUrl: f.webUrl || null,
            amount: known ? known.amount : null,
            currency: known ? known.currency : null,
            amountUsable: known ? known.usable : false,
          });
        }
      } catch (e) {
        errors.push(`${folder.name}: ${e.message}`);
      }
    }
  }));

  files.sort((a, b) => a.app.localeCompare(b.app) || String(b.month).localeCompare(String(a.month)) || a.name.localeCompare(b.name));

  const months = [...new Set(files.map(f => f.month).filter(Boolean))].sort();
  const apps = [...new Set(files.map(f => f.app))].sort((a, b) => a.localeCompare(b));

  return {
    scannedAt: new Date().toISOString(),
    root: root.path,
    mapping,
    files, apps, months, truncated,
    errors: errors.slice(0, 20),
    fileCount: files.length,
    undatedCount: files.filter(f => !f.month).length,
  };
}

module.exports = { buildInventory, parseMonthFolder, monthFromRelPath, amountsByFileName, monthIndexFromWord };
