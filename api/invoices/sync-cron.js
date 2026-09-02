const {
  getGraphToken, resolveDriveId, resolveShare, listFilesRecursive,
  encodeGraphPath, sanitizeSegment, readJsonFile, uploadFileContent,
  graphFetch, graphListAll, resolveArchiveRoot, archiveFile,
} = require('../../lib/graph');
const { verify, parseCookies } = require('../../lib/session');
const { runMailSync, resolvePending, collectSlackAnswers, notify, readRules, writeRules } = require('../../lib/mail-sync');
const { scanPeriods, applyBackfill } = require('../../lib/invoices/period-backfill');
const { readPending } = require('../../lib/invoices/pending');
const { openSpendSheet } = require('../../lib/spend-sheet');
const { readPdfText, extractInvoiceTotal, extractInvoiceRef } = require('../../lib/invoice-amount');
const { extractBillingPeriod, extractInvoiceDate } = require('../../lib/invoice-period');
const { normalizeRules } = require('../../lib/invoices/rules');

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
  const res = await graphFetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}?$select=id`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`lookup ${res.status}`);
  return (await res.json()).id;
}

// Mirror new files from the mapped source folders into {archive}/{App}/.
async function mirrorSourceFolders(token, targetDriveId, deadline, config, archiveRoot) {
  if (!config || !config.sourceUrl || !config.mapping) {
    return { note: 'No sync config saved yet — run an import from the dashboard once to set it up.' };
  }

  const share = await resolveShare(token, config.sourceUrl);
  if (!share.isFolder) throw new Error('Saved sourceUrl no longer points to a folder');

  // The source and the archive used to be two folders, with this copying between
  // them. They are now one and the same, and mirroring a folder onto itself is at
  // best a wasted crawl of the whole archive on every run.
  if (share.driveId === targetDriveId && share.itemId === archiveRoot.itemId) {
    return { note: `Source and archive are the same folder ("${archiveRoot.path}") — nothing to mirror.`, copiedNew: 0, alreadyPresent: 0, errors: [], perApp: {}, timedOut: false };
  }

  // List source subfolders (one per app/vendor). Paged: there is one folder per
  // vendor and they only accumulate, so a listing cut off at 200 would stop
  // mirroring every vendor past it without ever saying so.
  const children = await graphListAll(
    token,
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(share.driveId)}/items/${share.itemId}/children?$select=id,name,folder&$top=200`,
    'Source listing'
  );
  const subfolders = children.filter(f => f.folder);

  const summary = { copiedNew: 0, alreadyPresent: 0, errors: [], perApp: {}, timedOut: false };
  const mapped = subfolders.filter(f => config.mapping[f.name]);

  // Process one source folder: scan destination + source, copy anything new.
  async function processFolder(folder) {
    if (Date.now() > deadline) { summary.timedOut = true; return; }
    const targetAppRaw = config.mapping[folder.name];
    const targetApp = sanitizeSegment(targetAppRaw);

    let existing = new Set();
    try {
      const destFolder = await readDestFolderId(token, targetDriveId, `${archiveRoot.path}/${targetApp}`);
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
        const contentRes = await graphFetch(
          `https://graph.microsoft.com/v1.0/drives/${share.driveId}/items/${file.id}/content`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        if (!contentRes.ok) throw new Error(`download ${contentRes.status}`);
        const buf = Buffer.from(await contentRes.arrayBuffer());

        const relFolder = sanitizeRelPath(file.relPath);
        const destPath = `${archiveRoot.path}/${targetApp}${relFolder ? '/' + relFolder : ''}/${sanitizeFileName(file.name)}`;
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

// Two run summaries (answers applied, then mail filed) as one, for one DM.
function mergeSummaries(a, b) {
  const out = { ...b };
  for (const key of Object.keys(a)) {
    if (Array.isArray(a[key])) out[key] = a[key].concat(Array.isArray(b[key]) ? b[key] : []);
    else if (typeof a[key] === 'number') out[key] = a[key] + (typeof b[key] === 'number' ? b[key] : 0);
    else if (key === 'perApp') {
      out.perApp = { ...(b.perApp || {}) };
      for (const [app, n] of Object.entries(a.perApp || {})) out.perApp[app] = (out.perApp[app] || 0) + n;
    } else if (key === 'tracker') {
      out.tracker = { sheet: (b.tracker && b.tracker.sheet) || (a.tracker && a.tracker.sheet) || null, marked: ((a.tracker && a.tracker.marked) || 0) + ((b.tracker && b.tracker.marked) || 0) };
    }
  }
  return out;
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
    // ?mode=periods and ?mode=periods-apply are the billing-period backfill,
    // which lives here for the same reason the mailbox sync does: the Hobby plan
    // allows 12 serverless functions and the project is at exactly 12.
    const mode = String((req.query && req.query.mode) || 'all').trim();

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const token = await getGraphToken();
    const targetDriveId = await resolveDriveId(token, upn);
    const archiveRoot = await resolveArchiveRoot(token, targetDriveId, { fresh: true });

    // Loaded once and shared: the mailbox pass needs the folder->app mapping to
    // file each invoice under the vendor folder it has always lived in.
    const config = await readJsonFile(token, targetDriveId, archiveFile(archiveRoot, '_sync-config.json'));

    // The questions the sync is waiting on, with what the page needs to answer
    // them: the sheet's rows and the rules file the owner edits.
    if (mode === 'pending') {
      const held = await readPending(token, targetDriveId, archiveRoot);
      const sheet = await openSpendSheet(token);
      const rules = await readRules(token, targetDriveId, archiveRoot);
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({
        ok: true, mode, items: held.items, apps: sheet.grid.apps.map(a => a.name),
        rules, slackConfigured: !!(process.env.SLACK_BOT_TOKEN && process.env.SLACK_DM_USER),
      });
      return;
    }

    // Answers from the dashboard: [{ id, app, month, amount, ignore }].
    if (mode === 'pending-resolve') {
      const body = req.body || {};
      const answers = Array.isArray(body.answers) ? body.answers : [];
      const result = await resolvePending(token, targetDriveId, answers, {
        deadline: Date.now() + 45 * 1000, root: archiveRoot, mapping: config && config.mapping,
      });
      result.slack = await notify(token, targetDriveId, result, null, archiveRoot);
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ ok: true, mode, result });
      return;
    }

    // The owner's rules, saved from the dashboard's editor.
    if (mode === 'rules-save') {
      const body = req.body || {};
      const rules = normalizeRules(body.rules);
      if (!rules.vendors.length) throw new Error('The rules have no vendors in them — refusing to save an empty file');
      const saved = await writeRules(token, targetDriveId, archiveRoot, rules);
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ ok: true, mode, rules: saved });
      return;
    }

    // Reads one archived PDF the way the sync does and says what it found —
    // the text, the total, the period, the invoice number — so an invoice the
    // sync got wrong can be looked at without guessing what it saw.
    if (mode === 'inspect') {
      const path = String((req.query && req.query.path) || (req.body && req.body.path) || '').trim();
      if (!path) throw new Error('inspect needs ?path=<archive-relative path>');
      const full = path.startsWith(archiveRoot.path) ? path : `${archiveRoot.path}/${path.replace(/^\/+/, '')}`;
      const dl = await graphFetch(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(targetDriveId)}/root:/${encodeGraphPath(full)}:/content`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!dl.ok) throw new Error(`could not download "${full}" (${dl.status})`);
      const bytes = Buffer.from(await dl.arrayBuffer());
      const read = await readPdfText(bytes);
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({
        ok: true, mode, path: full, reader: read.reader, error: read.error,
        total: read.error ? null : extractInvoiceTotal(read.text),
        period: read.error ? null : extractBillingPeriod(read.text),
        invoiceDate: read.error ? null : extractInvoiceDate(read.text),
        ref: read.error ? null : extractInvoiceRef(read.text),
        text: read.text.slice(0, 6000),
      });
      return;
    }

    // The billing-period backfill. The scan writes nothing; the apply moves only
    // the files, and writes only the cells, that the scan proposed and the user
    // ticked — the browser sends both lists back.
    if (mode === 'periods' || mode === 'periods-apply') {
      const deadline = Date.now() + 45 * 1000;
      const result = mode === 'periods'
        ? await scanPeriods(token, targetDriveId, { deadline, root: archiveRoot, mapping: config && config.mapping })
        : await applyBackfill(token, targetDriveId, req.body || {}, { deadline, root: archiveRoot });
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ ok: true, ranAt: new Date().toISOString(), mode, periods: result });
      return;
    }

    // Time-box the run so we always return before the function's hard limit.
    // Because only NEW files are copied (existing ones skipped), a run that
    // stops early is simply finished by the next one — the mirror self-converges.
    // Both halves share the budget, so neither can starve the other completely.
    const started = Date.now();
    const HARD_DEADLINE = started + 45 * 1000;
    const out = { ok: true, ranAt: new Date().toISOString(), mode, archive: archiveRoot.path, archiveFound: archiveRoot.resolved };

    // Mailbox first, then the mirror. The mailbox pass files invoices into the
    // procurement folders, and the mirror copies new files from there into
    // Invoices/{App}/ for the dashboard — so running it second means an invoice
    // that arrives by mail reaches the dashboard in the same run rather than
    // waiting a day.
    if (mode !== 'folders') {
      // Anything the owner answered in the Slack DM since the last run is
      // applied first, so a held invoice is filed before its month is totalled.
      let slackAnswers = null;
      try {
        slackAnswers = await collectSlackAnswers(token, targetDriveId, {
          deadline: started + 20 * 1000, root: archiveRoot, mapping: config && config.mapping,
        });
        if (slackAnswers) out.answers = slackAnswers;
      } catch (e) {
        out.answers = { error: e.message };
      }
      // A mailbox failure (most likely a missing Mail.Read grant) must not take
      // the folder mirror down with it.
      try {
        out.mail = await runMailSync(token, targetDriveId, {
          deadline: mode === 'mail' ? HARD_DEADLINE : started + 25 * 1000,
          mapping: config && config.mapping,
          root: archiveRoot,
          // ?rescan=1 re-reads mail already seen, so totals can be picked up
          // from invoices that were filed before amounts were being read.
          rescan: !!(req.query && (req.query.rescan === '1' || req.query.rescan === 'true')),
        });
        // One DM per run, covering the answers applied and the mail filed.
        const merged = slackAnswers && slackAnswers.result ? mergeSummaries(slackAnswers.result, out.mail) : out.mail;
        out.mail.slack = await notify(token, targetDriveId, merged, slackAnswers, archiveRoot);
      } catch (e) {
        out.mail = { error: e.message };
      }
    }

    if (mode !== 'mail') {
      try {
        out.folders = await mirrorSourceFolders(token, targetDriveId, HARD_DEADLINE, config, archiveRoot);
      } catch (e) {
        out.folders = { error: e.message };
      }
    }

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json(out);
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
