const { listFilesRecursive, graphListAll, graphFetch, encodeGraphPath } = require('./graph');

// The contracts drawer: the SharePoint folder holding each vendor's signed
// agreement, order form and KYC paperwork — "Aggrements and Kyc", one subfolder
// per vendor.
//
// Deliberately NOT the invoice archive, and deliberately not mirrored into it.
// An order form is not a bill: the invoice checklist counts a file under
// {archive}/{App}/ as "that month is invoiced", so copying a 12-month order form
// in there would tick months nothing was billed for. These live where they are
// and are read in place.

const DOCUMENTS_TTL_MS = 10 * 60 * 1000;
const rootCache = new Map(); // driveId -> { path, itemId, candidates, resolved, expiresAt }

// Same reasoning as the invoice archive's candidate list in lib/graph.js: the
// folder has been renamed before and a hardcoded path fails by returning
// nothing rather than by erroring. Probe, take the first that exists.
function documentCandidates() {
  const out = [];
  for (const raw of [
    process.env.DOCUMENTS_PATH,
    'Desktop/Anudeep files/Aggrements and Kyc',
    'Desktop/Anudeep files/Agreements and KYC',
    'Agreements and KYC',
  ]) {
    const clean = String(raw || '').trim().replace(/^\/+|\/+$/g, '');
    if (clean && !out.includes(clean)) out.push(clean);
  }
  return out;
}

async function folderAtPath(token, driveId, path) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id,folder,webUrl`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph lookup of "${path}" failed (${res.status}): ${text.slice(0, 200)}`);
  }
  const json = await res.json();
  return json.folder ? { itemId: json.id, webUrl: json.webUrl } : null;
}

// Resolves to { path, itemId, webUrl, candidates, resolved }. `resolved` is
// false when no candidate exists — the caller reports the miss rather than
// showing an empty drawer as though there were no contracts.
async function resolveDocumentsRoot(token, driveId, opts) {
  const now = Date.now();
  const cached = rootCache.get(driveId);
  if (!(opts && opts.fresh) && cached && now < cached.expiresAt) return cached;

  const candidates = documentCandidates();
  for (const path of candidates) {
    const hit = await folderAtPath(token, driveId, path);
    if (hit) {
      const rec = { path, itemId: hit.itemId, webUrl: hit.webUrl, candidates, resolved: true, expiresAt: now + DOCUMENTS_TTL_MS };
      rootCache.set(driveId, rec);
      return rec;
    }
  }
  return { path: candidates[0], itemId: null, webUrl: null, candidates, resolved: false, expiresAt: 0 };
}

// ---------------------------------------------------------------------------
// What kind of document a file is
//
// Read off the file name, because that is all there is: the folder has no
// metadata and reading 14 PDFs to classify them would cost more than the page
// is worth. The four kinds are the ones actually in the folder — an order form
// or purchase order, a signed agreement, a KYC/compliance certificate, and
// everything else (escalation matrices, plan comparisons, forwarded zips).

const TYPES = {
  order:     { label: 'Order form', icon: '📝' },
  agreement: { label: 'Agreement',  icon: '✍️' },
  kyc:       { label: 'KYC',        icon: '🏛' },
  other:     { label: 'Other',      icon: '📄' },
};
const TYPE_KEYS = ['order', 'agreement', 'kyc', 'other'];

// KYC first: a "GST Certificate" is KYC even though a certificate is a kind of
// agreement-ish document, and "Cancelled cheque" must never read as an order.
const KYC_RE = /\bkyc\b|\bgst\b|\budyam\b|\bmsme\b|\bpan\s*card\b|\bpan\b|cancell?ed\s*cheque|\bcheque\b|incorporat|\btan\b|address\s*proof|\bcin\b|certificate/i;
// "signed" and "DealRoom" are here because that is how the signed order forms
// in this folder are named — a signed deal room document is the contract. The
// loose spelling of "agreement" is deliberate: the folder itself is called
// "Aggrements and Kyc" and its files are spelled to match.
const AGREEMENT_RE = /agg?r[e]{0,2}ments?|contract|\bmsa\b|\bnda\b|\bdpa\b|\bsow\b|dealroom|deal\s*room|terms\b|addendum|amendment|\bsigned\b/i;
const ORDER_RE = /order\s*form|purchase\s*order|\bpo\b|quotation|\bquote\b|\bQ-\d/i;

function classify(name) {
  const s = String(name || '');
  if (KYC_RE.test(s)) return 'kyc';
  if (AGREEMENT_RE.test(s)) return 'agreement';
  if (ORDER_RE.test(s)) return 'order';
  return 'other';
}

// ---------------------------------------------------------------------------
// The term a document covers
//
// "Order Form may 2026 to may 2027.pdf" is the whole reason this page is worth
// having: it says when the commitment ends. Only the explicit two-ended form is
// read — a single date in a file name is as likely to be when it was signed as
// when it expires, and guessing an expiry wrong is worse than showing none.

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const pad = n => String(n).padStart(2, '0');

function monthIndex(word) {
  const w = String(word || '').toLowerCase().slice(0, 3);
  return MONTHS.indexOf(w);
}

const SEP = '(?:\\s*(?:to|until|till|through|[-–—])\\s*)';
const NAMED = '([A-Za-z]{3,9})[\\s.,-]*((?:19|20)\\d{2})';
const ISO = '((?:19|20)\\d{2})-(\\d{1,2})-(\\d{1,2})';

// Returns { start, end } as YYYY-MM strings (ISO ranges keep their day), or null.
function extractTerm(name) {
  const s = String(name || '').replace(/\.[A-Za-z0-9]{1,5}$/, ' ');

  let m = s.match(new RegExp(`${ISO}${SEP}${ISO}`));
  if (m) {
    const start = `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
    const end = `${m[4]}-${pad(+m[5])}-${pad(+m[6])}`;
    return end > start ? { start, end } : null;
  }

  m = s.match(new RegExp(`${NAMED}${SEP}${NAMED}`, 'i'));
  if (m) {
    const a = monthIndex(m[1]), b = monthIndex(m[3]);
    if (a >= 0 && b >= 0) {
      const start = `${m[2]}-${pad(a + 1)}`;
      const end = `${m[4]}-${pad(b + 1)}`;
      return end > start ? { start, end } : null;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------

// The folder holds a handful of files per vendor, so this is one Graph listing
// per vendor folder — cheap enough to do on demand, but still bounded by a
// deadline so a slow drive returns a partial page rather than a 504.
async function buildDocumentIndex(token, driveId, opts) {
  const options = opts || {};
  const deadline = options.deadline || (Date.now() + 40 * 1000);
  const root = await resolveDocumentsRoot(token, driveId, { fresh: options.fresh });

  const base = {
    root: root.path,
    rootUrl: root.webUrl || null,
    resolved: root.resolved,
    candidates: root.candidates,
    vendors: [],
    totals: { vendors: 0, files: 0, order: 0, agreement: 0, kyc: 0, other: 0 },
    syncedAt: new Date().toISOString(),
    truncated: false,
  };
  if (!root.resolved) return base;

  const children = await graphListAll(
    token,
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(root.itemId)}/children?$select=id,name,file,folder,size,createdDateTime,lastModifiedDateTime,webUrl&$top=200`,
    'Documents folder listing'
  );

  const vendors = [];

  // Files sitting loose at the top level belong to no vendor. They are shown
  // rather than dropped: a contract nobody filed is exactly the one to notice.
  const loose = children.filter(c => c.file);
  if (loose.length) vendors.push({ vendor: 'Unfiled', items: loose.map(f => ({ ...f, relPath: '' })) });

  for (const folder of children.filter(c => c.folder)) {
    if (Date.now() > deadline) { base.truncated = true; break; }
    try {
      const files = await listFilesRecursive(token, driveId, folder.id);
      vendors.push({ vendor: folder.name, webUrl: folder.webUrl, items: files });
    } catch (e) {
      vendors.push({ vendor: folder.name, webUrl: folder.webUrl, items: [], error: e.message });
    }
  }

  base.vendors = vendors
    .map(v => {
      const counts = { order: 0, agreement: 0, kyc: 0, other: 0 };
      const files = v.items.map(f => {
        const type = classify(f.name);
        counts[type]++;
        return {
          name: f.name,
          type,
          typeLabel: TYPES[type].label,
          size: f.size || 0,
          uploadedAt: f.lastModifiedDateTime || f.createdDateTime || null,
          webUrl: f.webUrl || null,
          subfolder: f.relPath || '',
          term: extractTerm(f.name),
        };
      }).sort((a, b) => TYPE_KEYS.indexOf(a.type) - TYPE_KEYS.indexOf(b.type) || a.name.localeCompare(b.name));
      return { vendor: v.vendor, webUrl: v.webUrl || null, error: v.error || null, counts, files };
    })
    .sort((a, b) => a.vendor.localeCompare(b.vendor));

  for (const v of base.vendors) {
    base.totals.vendors++;
    for (const f of v.files) { base.totals.files++; base.totals[f.type]++; }
  }
  return base;
}

module.exports = {
  documentCandidates, resolveDocumentsRoot, buildDocumentIndex,
  classify, extractTerm, TYPES, TYPE_KEYS,
};
