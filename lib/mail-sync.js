const { encodeGraphPath, sanitizeSegment, readJsonFile, writeJsonFile, uploadFileContent } = require('./graph');
const excel = require('./excel');
const { openSpendSheet, readAliasMap } = require('./spend-sheet');
const { buildResolver } = require('./vendor-map');
const mail = require('./mail');

// Files invoices from the shared mailbox. Called by the daily invoice cron
// rather than living at its own api/ path, because each file under api/ counts
// against the Hobby plan's 12-Serverless-Function limit for a deployment.

const STATE_PATH = 'Invoices/_mail-sync.json';
const INDEX_PATH = 'Invoices/_invoice-index.json';
const UNMATCHED_FOLDER = '_Unmatched';
const FIRST_RUN_LOOKBACK_DAYS = 60;
const DEADLINE_MS = 45 * 1000;

// Invoices are archived in the same place they have always been filed by hand —
// the procurement folder — not in a store of the app's own. The daily folder
// mirror then copies them on to Invoices/{App}/ for the dashboard, so there is
// one canonical archive rather than two competing ones.
// Note the folder really is spelled "Procurment".
const sourceBase = () =>
  (process.env.INVOICE_SOURCE_PATH || 'Desktop/Anudeep files/Procurment bills').trim().replace(/^\/+|\/+$/g, '');

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// The saved import mapping is sourceFolder -> app; invoices need the reverse.
function appToSourceFolder(mapping) {
  const out = {};
  for (const [folder, app] of Object.entries(mapping || {})) {
    if (app && !out[app]) out[app] = folder; // first mapping wins
  }
  return out;
}

// Existing month subfolders are inconsistent ("Aug" under Bubble, "July" under
// Cursor), so match whatever is already there for this month instead of adding
// a parallel folder beside it. Falls back to "Aug-26" style when none exists.
function monthFolderName(existingFolderNames, month) {
  const [year, mm] = String(month || '').split('-');
  const idx = Number(mm) - 1;
  if (!year || !(idx >= 0 && idx < 12)) return null;
  const full = MONTH_NAMES[idx];
  const abbr = full.slice(0, 3);
  const yy = year.slice(2);
  const candidates = [`${abbr}-${yy}`, `${abbr} ${yy}`, `${month}`, `${full} ${year}`, `${full}-${yy}`, full, abbr];
  const bySlug = new Map((existingFolderNames || []).map(n => [String(n).toLowerCase().trim(), n]));
  for (const c of candidates) {
    const hit = bySlug.get(c.toLowerCase());
    if (hit) return hit;
  }
  return `${abbr}-${yy}`;
}

function sanitizeFileName(name) {
  const dot = String(name || '').lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : (name || 'invoice');
  const ext = dot > 0 ? name.slice(dot).replace(/[<>:"/\\|?*\x00-\x1F]/g, '') : '.pdf';
  return sanitizeSegment(base) + ext;
}

// Children of a folder path, split into file names and subfolder names.
// A missing folder is not an error — it simply has no children yet.
async function folderChildren(token, driveId, path) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/children?$select=name,folder,file&$top=200`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) return { files: new Set(), folders: [] };
  const json = await res.json();
  const files = new Set();
  const folders = [];
  for (const c of json.value || []) {
    if (c.folder) folders.push(c.name);
    else files.add(c.name);
  }
  return { files, folders };
}

async function folderFileNames(token, driveId, path) {
  return (await folderChildren(token, driveId, path)).files;
}

// Ticks the app's month in the "Invoices tracker" sheet. That sheet has the same
// app-rows/month-columns shape as the spend sheet but holds TRUE/FALSE.
// Works out which tracker cells actually need writing. Kept separate from the
// Graph calls so the rules are testable: one cell per app-month no matter how
// many invoices arrived, nothing already ticked, nothing off-grid.
function planTrackerCells(marks, grid, used) {
  const rowByApp = new Map(grid.apps.map(a => [a.name, a.rowIdx]));
  const cells = [];
  const seenAddresses = new Set();
  for (const { app, month } of marks || []) {
    const rowIdx = rowByApp.get(app);
    const colIdx = grid.monthCols[month];
    if (rowIdx === undefined || colIdx === undefined) continue;
    const addr = excel.cellAddress(used.start, rowIdx, colIdx);
    if (seenAddresses.has(addr)) continue; // never write one cell twice
    seenAddresses.add(addr);
    const raw = (used.values[rowIdx] || [])[colIdx];
    const existing = String(raw == null ? '' : raw).trim();
    if (/^true$/i.test(existing)) continue; // already ticked
    cells.push({ app, month, address: addr, value: true });
  }
  return cells;
}

async function markTracker(token, driveId, itemId, marks, sessionId) {
  const sheets = await excel.listWorksheets(token, driveId, itemId, sessionId);
  const tracker = sheets.map(s => s.name).find(n => /tracker/i.test(n));
  if (!tracker) return { sheet: null, marked: 0 };

  const used = await excel.readUsedRange(token, driveId, itemId, tracker, sessionId);
  const grid = excel.locateGrid(used.values, used.text);
  const cells = planTrackerCells(marks, grid, used);
  if (!cells.length) return { sheet: tracker, marked: 0 };

  const results = await excel.writeCells(token, driveId, itemId, tracker, cells, sessionId, 4);
  return { sheet: tracker, marked: results.filter(r => r.ok).length };
}

// Returns a summary of what was filed. `deadline` is a timestamp the run stops
// at, so the caller can share one function timeout across several jobs.
async function runMailSync(token, driveId, options) {
  const opts = options || {};
  const mailbox = mail.mailboxAddress();
  const base = sourceBase();
  // Reuse the folder->app mapping the invoice import already saved, so invoices
  // land in the vendor folder they have always been filed under.
  const folderForApp = appToSourceFolder(opts.mapping);

  const state = (await readJsonFile(token, driveId, STATE_PATH)) || {};
  const seen = new Set(Array.isArray(state.seenMessageIds) ? state.seenMessageIds : []);
  const since = state.lastRunAt
    || new Date(Date.now() - FIRST_RUN_LOOKBACK_DAYS * 86400000).toISOString();

  const sheet = await openSpendSheet(token);
  const appNames = sheet.grid.apps.map(a => a.name);
  const resolve = buildResolver(await readAliasMap(token, driveId), appNames);

  const messages = await mail.listMessages(token, mailbox, since.replace(/\.\d+Z$/, 'Z'), 50);

  const summary = {
    mailbox, since, scanned: messages.length,
    filed: 0, alreadyPresent: 0, skippedNotInvoice: 0, unmatched: 0,
    perApp: {}, unmatchedItems: [], errors: [], timedOut: false,
    base, newFolders: [],
  };
  const indexEntries = [];
  const marks = [];
  const markKeys = new Set();
  const deadline = opts.deadline || (Date.now() + DEADLINE_MS);

  for (const message of messages) {
    if (Date.now() > deadline) { summary.timedOut = true; break; }
    if (seen.has(message.id)) continue;

    // The Graph query no longer filters on hasAttachments (see lib/mail.js), so
    // skip attachment-less mail here before spending a call listing nothing.
    if (!message.hasAttachments) {
      summary.skippedNotInvoice++;
      seen.add(message.id);
      continue;
    }

    let attachments;
    try {
      attachments = await mail.listAttachments(token, mailbox, message.id);
    } catch (e) {
      summary.errors.push(`${(message.subject || '').slice(0, 60)}: ${e.message}`);
      continue;
    }

    if (!mail.looksLikeInvoice(message, attachments)) {
      summary.skippedNotInvoice++;
      seen.add(message.id);
      continue;
    }

    const identity = mail.senderIdentity(message);
    const month = mail.messageMonth(message);

    // Try each signal in turn — subject, attachment name, sender domain — and
    // take the first confident match rather than a fuzzy guess.
    let app = null;
    for (const text of mail.matchText(message, attachments, identity)) {
      const hit = resolve('', text);
      if (hit.app) { app = hit.app; break; }
    }

    const pdfs = attachments.filter(mail.isPdf);

    // Vendor folder: the one this app is already mapped to; else the app's own
    // name, reported below because the folder mirror will not pick it up until
    // it is mapped; else the unmatched bucket.
    let vendorFolder;
    let newFolder = false;
    if (app) {
      if (folderForApp[app]) {
        vendorFolder = folderForApp[app];
      } else {
        vendorFolder = sanitizeSegment(app);
        newFolder = true;
      }
    } else {
      vendorFolder = UNMATCHED_FOLDER;
    }

    const { folders: monthFolders } = await folderChildren(token, driveId, `${base}/${vendorFolder}`);
    const monthFolder = monthFolderName(monthFolders, month) || month;
    const targetFolder = `${base}/${vendorFolder}/${monthFolder}`;
    const existing = await folderFileNames(token, driveId, targetFolder);

    if (newFolder && !summary.newFolders.includes(vendorFolder)) summary.newFolders.push(vendorFolder);

    for (const att of pdfs) {
      if (Date.now() > deadline) { summary.timedOut = true; break; }
      const fileName = sanitizeFileName(att.name);
      if (existing.has(fileName)) { summary.alreadyPresent++; continue; }
      try {
        const bytes = await mail.getAttachmentBytes(token, mailbox, message.id, att.id);
        await uploadFileContent(token, driveId, `${targetFolder}/${fileName}`, bytes, 'application/pdf');
        summary.filed++;
        if (app) summary.perApp[app] = (summary.perApp[app] || 0) + 1;
        indexEntries.push({
          app: app || null, month, file: fileName, folder: targetFolder,
          subject: message.subject || '', from: identity.originalAddress || identity.address,
          receivedAt: message.receivedDateTime, webLink: message.webLink || null,
        });
      } catch (e) {
        summary.errors.push(`${fileName}: ${e.message}`);
      }
    }

    if (app) {
      // Several invoices for the same app in one month are normal — they tick
      // one cell, not one per message.
      const markKey = `${app}||${month}`;
      if (month && !markKeys.has(markKey)) {
        markKeys.add(markKey);
        marks.push({ app, month });
      }
    } else {
      summary.unmatched++;
      summary.unmatchedItems.push({
        subject: (message.subject || '').slice(0, 90),
        from: identity.originalAddress || identity.address,
        month, folder: targetFolder,
      });
    }
    seen.add(message.id);
  }

  // Tick the invoice tracker for everything filed against a known app.
  let tracker = { sheet: null, marked: 0 };
  if (marks.length) {
    let sessionId = null;
    try {
      sessionId = await excel.createSession(token, sheet.driveId, sheet.itemId);
      tracker = await markTracker(token, sheet.driveId, sheet.itemId, marks, sessionId);
    } catch (e) {
      summary.errors.push(`tracker update: ${e.message}`);
    } finally {
      if (sessionId) await excel.closeSession(token, sheet.driveId, sheet.itemId, sessionId);
    }
  }

  if (indexEntries.length) {
    const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
    const merged = (Array.isArray(index.entries) ? index.entries : []).concat(indexEntries).slice(-500);
    await writeJsonFile(token, driveId, INDEX_PATH, { entries: merged, updatedAt: new Date().toISOString() });
  }

  await writeJsonFile(token, driveId, STATE_PATH, {
    mailbox,
    // On a timed-out run, don't advance the watermark past what was processed —
    // the next run picks the remainder up.
    lastRunAt: summary.timedOut ? state.lastRunAt || since : new Date().toISOString(),
    seenMessageIds: [...seen].slice(-500),
    updatedAt: new Date().toISOString(),
  });

  return { ok: true, ranAt: new Date().toISOString(), tracker, ...summary };
}

module.exports = { runMailSync, appToSourceFolder, monthFolderName, sourceBase, planTrackerCells, STATE_PATH, INDEX_PATH };
