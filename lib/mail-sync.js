const { encodeGraphPath, sanitizeSegment, readJsonFile, writeJsonFile, uploadFileContent, resolveArchiveRoot, archiveFile, graphFetch, itemIdByPath, ensureFolder, moveItem } = require('./graph');
const excel = require('./excel');
const { openSpendSheet, readAliasMap, cellValue, appendLog } = require('./spend-sheet');
const { buildResolver } = require('./vendor-map');
const mail = require('./mail');
const { extractInvoiceTotal, extractInvoiceRef, readPdfText } = require('./invoice-amount');
const { invoiceMonth, extractInvoiceDate } = require('./invoice-period');
const { monthFromFileName, detectDayFirst, parseMonthFolder } = require('./invoices/inventory');
// The tracker write is shared with the checklist's backfill, which ticks what
// the archive already holds rather than what a run just filed. Same rules.
const { planTrackerCells, markTracker } = require('./invoices/tracker');
const rulesLib = require('./invoices/rules');
const pending = require('./invoices/pending');
const slack = require('./slack');

// Files invoices from the shared mailbox. Called by the daily invoice cron
// rather than living at its own api/ path, because each file under api/ counts
// against the Hobby plan's 12-Serverless-Function limit for a deployment.

const UNMATCHED_FOLDER = '_Unmatched';
const FIRST_RUN_LOOKBACK_DAYS = 60;
const DEADLINE_MS = 45 * 1000;

// Invoices are archived in the same place they have always been filed by hand,
// under the app's own folder — not in a store of the app's own. That folder is
// located at run time by resolveArchiveRoot rather than assumed: it was once
// hardcoded as "Procurment bills", and when it was renamed this kept filing into
// the old name, recreating it and splitting the archive in two. See lib/graph.js.

// Same normalization the vendor resolver uses, so a folder and an app name are
// compared the way the rest of the app compares them.
const normName = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// The saved import mapping is sourceFolder -> app; invoices need the reverse.
//
// Several folders can map to one app — "Luzmo" and "Cumul(Luzmo)" both mean
// Cumul(Luzmo), and "Bubble" and "Bubble Starter" both mean Bubble Starter —
// so the reverse has to choose. Taking the first listed picked "Bubble", a
// folder that had since been renamed to "Bubble Starter": filing there would
// have recreated the old name and split the vendor's invoices across two
// folders, with only one of them ever totalled.
//
// A folder that exists therefore beats one that does not. `existingFolders` is
// what the archive actually contains; without it the old first-wins rule stands,
// so callers that cannot cheaply list the archive are unaffected.
//
// And a folder the archive spells exactly like the app beats the mapping
// outright. The mapping is a snapshot of the folder names at the last import,
// so renaming a folder to match the sheet silently strands it: the saved entry
// still says "click up invoices" long after that folder became
// "Mango technology(Clickup)", and filing by the stale name recreates the old
// folder beside the real one. The sheet's own name is the one the archive is
// being kept in step with, so where the archive already carries it, it wins.
function appToSourceFolder(mapping, existingFolders, appNames) {
  const present = new Set((existingFolders || []).map(n => String(n).toLowerCase()));
  const has = name => present.has(String(name).toLowerCase());
  const out = {};
  for (const [folder, app] of Object.entries(mapping || {})) {
    if (!app) continue;
    const chosen = out[app];
    if (chosen === undefined) { out[app] = folder; continue; }
    // Only ever trade up: a real folder replaces one that is not there.
    if (present.size && !has(chosen) && has(folder)) out[app] = folder;
  }
  // Case and punctuation drift between the sheet and the archive ("apollo" for
  // "Apollo", "Google Cloud" for "Google cloud"), so match on the same
  // normalized form the vendor resolver uses, and file under the folder's own
  // spelling rather than the sheet's.
  const byNorm = new Map();
  for (const name of (existingFolders || [])) byNorm.set(normName(name), name);
  for (const app of (appNames || [])) {
    const folder = byNorm.get(normName(app));
    if (folder) out[app] = folder;
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
//
// A missing folder is not an error — it simply has no children yet. Any OTHER
// failure is, and used to be swallowed as "empty": a throttled listing then read
// as "this vendor has no month folders", and the run created `Aug-26` beside the
// existing `Aug`, splitting the month's invoices across two folders and totalling
// each of them short. It throws now, and the run reports the file it could not
// place. Paged, for the same reason a truncated listing is worse than no listing.
async function folderChildren(token, driveId, path) {
  const files = new Set();
  const fileItems = [];
  const folders = [];
  let url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/children?$select=name,folder,file,createdDateTime&$top=200`;
  let pages = 0;

  while (url && pages < 20) {
    const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (res.status === 404) return { files, fileItems, folders, missing: true };
    if (!res.ok) throw new Error(`could not list "${path}" (${res.status})`);
    const json = await res.json();
    for (const c of json.value || []) {
      if (c.folder) folders.push(c.name);
      else { files.add(c.name); fileItems.push({ name: c.name, createdDateTime: c.createdDateTime }); }
    }
    url = json['@odata.nextLink'] || null;
    pages++;
  }
  return { files, fileItems, folders };
}

// Invoices reach the archive by more than one route. The mailbox pass only ever
// learns about an app-month because mail arrived for it, so an invoice dropped
// into a folder by hand — or mirrored in from a source folder — was never
// totalled and never ticked until some unrelated mail happened to land on the
// same app-month. Nothing said so; the sheet just stayed blank.
//
// So each run also sweeps the archive's own folders for recent months and hands
// back the app-months holding invoices, to be totalled and ticked by exactly the
// path the mailbox pass uses. Bounded three ways, because this shares one
// function timeout with the mailbox pass and the folder mirror: only the last
// few months, only folders that resolve to a row in the sheet, and it stops at
// the deadline rather than running long. Stopping early loses nothing — the
// next run picks up whatever was missed.
const RECONCILE_MONTHS = 3;

function recentMonths(count, now) {
  const at = now ? new Date(now) : new Date();
  const out = [];
  for (let i = 0; i < count; i++) {
    const d = new Date(Date.UTC(at.getUTCFullYear(), at.getUTCMonth() - i, 1));
    out.push(`${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}

async function reconcileArchive(token, driveId, base, opts) {
  const options = opts || {};
  const folders = options.folders || [];
  const wanted = new Set(recentMonths(options.months || RECONCILE_MONTHS, options.now));
  const deadline = options.deadline || (Date.now() + 15 * 1000);
  const found = [];
  const errors = [];
  let timedOut = false;

  let cursor = 0;
  const POOL = options.pool || 4;
  await Promise.all(Array.from({ length: Math.min(POOL, folders.length) }, async () => {
    while (cursor < folders.length) {
      if (Date.now() > deadline) { timedOut = true; return; }
      const folder = folders[cursor++];
      const app = options.appFor(folder);
      if (!app) continue; // no row in the sheet: the checklist reports it, this does not guess
      const path = `${base}/${folder}`;
      try {
        const { folders: months } = await folderChildren(token, driveId, path);
        for (const name of months) {
          const month = parseMonthFolder(name, options.now);
          if (month && wanted.has(month)) found.push({ app, month, folder: `${path}/${name}` });
        }
      } catch (e) {
        errors.push(`${folder}: ${e.message}`);
      }
    }
  }));
  return { found, errors, timedOut };
}

// Invoices for a month are not all in that month's folder. Apollo keeps
// "Invoice-A0589F17-0016-Aug 2026.pdf" loose in the vendor folder while
// Aug-26/ holds only the 27th's invoice, so totalling the month folder alone
// reported 85.00 for a month that really cost 138.12 — and, being higher than
// the 53.12 already in the cell, it overwrote and lost the 4 August charge.
//
// Loose files are dated by their name, exactly as the checklist dates them, so
// the two agree on which month an invoice belongs to.
async function looseFilesForMonth(token, driveId, appFolderPath, month) {
  const { fileItems } = await folderChildren(token, driveId, appFolderPath);
  const pdfs = (fileItems || []).filter(f => /\.pdf$/i.test(f.name));
  // "03-12-2025" only reads once another name in the folder settles the order.
  const dayFirst = detectDayFirst(pdfs.map(f => f.name));
  return pdfs
    .filter(f => monthFromFileName(f.name, f.createdDateTime, dayFirst) === month)
    .map(f => ({ name: f.name, path: `${appFolderPath}/${f.name}` }));
}

async function folderFileNames(token, driveId, path) {
  return (await folderChildren(token, driveId, path)).files;
}

// Ticks the app's month in the "Invoices tracker" sheet. That sheet has the same
// app-rows/month-columns shape as the spend sheet but holds TRUE/FALSE.
// Totals every invoice in one app-month folder, not just the ones that happened
// to arrive by email. Invoices reach the archive by several routes, so an
// email-only total undercounts — badly, and silently.
//
// Each PDF is parsed once and its result cached in the invoice index by path, so
// repeat runs cost a folder listing rather than a re-download of everything.
async function sumFolderInvoices(token, driveId, folderPath, cache, budget, extraFiles) {
  const { files } = await folderChildren(token, driveId, folderPath);
  // `extraFiles` are invoices for this same month that sit outside its folder —
  // see looseFilesForMonth. They are totalled with it because they are the
  // month's spend wherever they happen to be filed.
  const entries = [...files].filter(n => /\.pdf$/i.test(n))
    .map(name => ({ name, key: `${folderPath}/${name}` }))
    .concat((extraFiles || []).map(f => ({ name: f.name, key: f.path })));

  let total = 0;
  let counted = 0;
  const unusable = [];
  const unread = [];
  const duplicates = [];
  const seenRefs = new Map(); // invoice number -> the file that first claimed it

  for (const { name, key } of entries) {
    let entry = cache.get(key);

    // An entry parsed by the old reader (nothing recorded which one read it)
    // is read again: it may hold a truncated invoice number, a first-of-four
    // total, or no total at all for a file the new reader handles. Once per
    // file — the re-read records its reader, so it is not repeated.
    if (entry && !entry.reader) entry = null;

    if (!entry) {
      if (budget.parsed >= budget.maxParse || Date.now() > budget.deadline) {
        budget.exhausted = true;
        break;
      }
      budget.parsed++;
      try {
        const res = await graphFetch(
          `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(key)}:/content`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        if (!res.ok) throw new Error(`download ${res.status}`);
        const bytes = Buffer.from(await res.arrayBuffer());
        const { text, error, reader } = await readPdfText(bytes);
        entry = error ? { amount: null, currency: null, usable: false, note: error } : extractInvoiceTotal(text, budget.options);
        // Kept so an invoice and its own payment receipt can be recognised as
        // one charge rather than two.
        if (!error) entry.ref = entry.refs ? null : extractInvoiceRef(text);
        entry.reader = reader || 'none';
      } catch (e) {
        entry = { amount: null, currency: null, usable: false, note: e.message, reader: 'none' };
      }
      cache.set(key, entry);
      budget.fresh.push({
        path: key, amount: entry.amount, currency: entry.currency, usable: !!entry.usable,
        ref: entry.ref || null, refs: entry.refs || null, note: entry.note || '', reader: entry.reader, via: entry.via || null,
      });
    }

    if (entry.usable) {
      // An invoice and its payment receipt are one charge sent twice. They carry
      // the same invoice number, so the second one seen is a duplicate of the
      // first, not more spend. Only a reference actually read counts: without
      // one, two files that merely happen to share an amount are two charges.
      // A file holding several invoices carries several numbers.
      const refs = entry.refs || (entry.ref ? [entry.ref] : []);
      const dup = refs.find(r => seenRefs.has(r));
      if (dup) {
        duplicates.push({ file: name, amount: entry.amount, ref: dup, of: seenRefs.get(dup) });
        continue;
      }
      for (const r of refs) seenRefs.set(r, name);
      total += entry.amount;
      counted++;
    }
    else if (entry.amount !== null) unusable.push({ file: name, amount: entry.amount, currency: entry.currency, note: entry.note });
    else unread.push({ file: name, note: entry.note });
  }

  return {
    total: Math.round(total * 100) / 100, counted,
    // A duplicate is accounted for, not missing, so it must not make the folder
    // look partially read — that would stop the total being written at all.
    pdfCount: entries.length - duplicates.length,
    unusable, unread, duplicates,
  };
}

// Decides what each Spendings cell should hold, given the invoices on file.
//
// The rule, as the sheet's owner set it: a cell IS the total of that month's
// invoices. Several vendors send three or four invoices a month; they are added
// up and the sum goes in the cell, replacing whatever is there — higher or
// lower. Writing the whole month's total, rather than adding an invoice on, is
// what makes this safe to run again and again: the same folder writes the same
// number twice, never twice the number.
//
// The one thing that may not be written is a total that is not the total. When
// some PDF in the month could not be read, the sum of the rest is a lower
// bound, and putting a lower bound in a cell that claims to be the month's
// spend is exactly the kind of guess the owner asked never to make. Those
// months are held and reported, with the file that needs fixing, and the cell
// is written on the next run once every invoice reads.
const CELL_EPSILON = 0.005;

function planAmountCells(amounts, grid, used, cellValueOf, priorWrites, partialTotals, locks) {
  const rowByApp = new Map(grid.apps.map(a => [a.name, a.rowIdx]));
  const byNorm = new Map(grid.apps.map(a => [normName(a.name), a.rowIdx]));
  const prior = priorWrites || {};
  const partial = partialTotals || {};
  const write = [];
  const updated = [];
  const skippedFilled = [];
  const locked = [];
  const lockList = Array.isArray(locks) ? locks : [];
  const lockOf = (app, month) => lockList.find(l => normName(l.app) === normName(app) && l.month === month) || null;

  // A locked cell holds the owner's figure, whatever the month's invoices
  // come to. It is enforced first, and then left out of everything below.
  const lockedKeys = new Set();
  for (const l of lockList) {
    const rowIdx = byNorm.get(normName(l.app));
    const colIdx = grid.monthCols[l.month];
    if (rowIdx === undefined || colIdx === undefined) continue;
    const appName = grid.apps.find(a => a.rowIdx === rowIdx).name;
    lockedKeys.add(`${appName}||${l.month}`);
    const current = cellValueOf(used.values, rowIdx, colIdx);
    const address = excel.cellAddress(used.start, rowIdx, colIdx);
    const invoiceTotal = amounts && amounts[appName] ? amounts[appName][l.month] : undefined;
    const held = current === null || current === 0 ? 0 : current;
    if (Math.abs(held - l.value) < CELL_EPSILON) {
      locked.push({ app: appName, month: l.month, address, value: l.value, current, invoiceTotal, enforced: false, note: l.note });
    } else {
      locked.push({ app: appName, month: l.month, address, value: l.value, current, invoiceTotal, enforced: true, note: l.note });
      updated.push({ app: appName, month: l.month, address, value: l.value, previous: current, direction: l.value > held ? 'up' : 'down', wasOurs: true, locked: true });
    }
  }

  for (const [app, byMonth] of Object.entries(amounts || {})) {
    const rowIdx = rowByApp.get(app);
    if (rowIdx === undefined) continue;
    for (const [month, amount] of Object.entries(byMonth)) {
      const colIdx = grid.monthCols[month];
      if (colIdx === undefined) continue;
      if (lockedKeys.has(`${app}||${month}`) || lockOf(app, month)) continue;
      const current = cellValueOf(used.values, rowIdx, colIdx);
      const address = excel.cellAddress(used.start, rowIdx, colIdx);

      // A month whose invoices have all been moved elsewhere totals nothing.
      // Its cell is cleared only when the figure there is one this sync wrote —
      // a statement figure for a month that never had invoices is not touched.
      if (!(amount > 0)) {
        const mine = prior[`${app}||${month}`];
        if (amount === 0 && current && mine !== undefined && Math.abs(current - mine) < CELL_EPSILON) {
          updated.push({ app, month, address, value: 0, previous: current, direction: 'down', wasOurs: true, emptied: true });
        }
        continue;
      }

      if (partial[`${app}||${month}`] === true) {
        skippedFilled.push({ app, month, address, current, invoiceTotal: amount, reason: 'folder-total-incomplete' });
        continue;
      }

      if (current === null || current === 0) {
        write.push({ app, month, address, value: amount, partial: false });
        continue;
      }

      if (Math.abs(amount - current) < CELL_EPSILON) continue; // already right

      const mine = prior[`${app}||${month}`];
      updated.push({
        app, month, address, value: amount, previous: current,
        direction: amount > current ? 'up' : 'down',
        // false when the figure being replaced is not one this sync wrote — a
        // hand correction or a statement figure. It is still replaced, because
        // the cell is defined as the invoice total; but replacing someone's
        // number is never silent, so the summary and the audit log say so.
        wasOurs: mine !== undefined && Math.abs(current - mine) < CELL_EPSILON,
      });
    }
  }
  return { write, updated, skippedFilled, locked };
}

// --- The rules file ---------------------------------------------------------

// The owner's rules, seeded on first use so there is a file to edit.
async function readRules(token, driveId, root) {
  const path = archiveFile(root, rulesLib.RULES_FILE);
  const raw = await readJsonFile(token, driveId, path);
  if (raw && Array.isArray(raw.vendors)) return rulesLib.normalizeRules(raw);
  const seeded = rulesLib.normalizeRules(rulesLib.SEED_RULES);
  await writeJsonFile(token, driveId, path, { ...seeded, savedAt: new Date().toISOString(), note: 'Edit from the dashboard, or here. See lib/invoices/rules.js for the shape.' });
  return seeded;
}

async function writeRules(token, driveId, root, rules) {
  const clean = rulesLib.normalizeRules(rules);
  await writeJsonFile(token, driveId, archiveFile(root, rulesLib.RULES_FILE), { ...clean, savedAt: new Date().toISOString() });
  return clean;
}

// --- Where an invoice goes ---------------------------------------------------

// Everything the sync needs to place a file, resolved once per run: the sheet
// (which rows exist), the archive's vendor folders, and a cache of month
// folders already looked up. Shared by the mailbox pass and by answering a
// held question, so both file a PDF by exactly the same path.
async function openPlacement(token, driveId, opts) {
  const root = opts.root || await resolveArchiveRoot(token, driveId);
  const base = root.path;
  const sheet = await openSpendSheet(token);
  const appNames = sheet.grid.apps.map(a => a.name);
  const errors = [];
  let archiveFolders = [];
  try {
    archiveFolders = (await folderChildren(token, driveId, base)).folders;
  } catch (e) {
    errors.push(`could not list "${base}" to check which vendor folders exist (${e.message})`);
  }
  const folderForApp = appToSourceFolder(opts.mapping, archiveFolders, appNames);

  // Where one vendor-folder/month pair lives, and what is already in it. Each
  // pair costs two Graph listings, so they are resolved once and remembered.
  const places = new Map();
  const placeFor = async (vendorFolder, monthKey) => {
    const key = `${vendorFolder}||${monthKey}`;
    if (!places.has(key)) {
      const { folders } = await folderChildren(token, driveId, `${base}/${vendorFolder}`);
      const folder = `${base}/${vendorFolder}/${monthFolderName(folders, monthKey) || monthKey}`;
      places.set(key, { folder, files: await folderFileNames(token, driveId, folder) });
    }
    return places.get(key);
  };
  const vendorFolderFor = (app) => folderForApp[app] || sanitizeSegment(app);

  return {
    root, base, sheet, appNames, archiveFolders, folderForApp, placeFor, vendorFolderFor, errors,
    INDEX_PATH: archiveFile(root, '_invoice-index.json'),
  };
}

// The month an invoice bills, by the owner's rules: the month its billing
// period starts in; failing a stated period, the month it was issued; failing
// both, nothing — and nothing means ask, never the month the mail arrived.
// `receivedMonth` is only used to catch a period so far from the mail that it
// must be a misread (a contract term, a renewal pair), which is also a question.
function monthForInvoice(text, receivedMonth, options) {
  const placement = invoiceMonth(text, receivedMonth, { usageRange: !!(options && options.usageRange) });
  if (placement.period && !placement.ignoredPeriod) {
    return { month: placement.month, via: placement.via, period: placement.period, issued: null };
  }
  const issued = extractInvoiceDate(text);
  if (placement.period && placement.ignoredPeriod) {
    return {
      month: null, via: 'period-far', period: placement.period, issued,
      why: `the invoice states a period ${placement.period.start} → ${placement.period.end}, months away from when the mail arrived (${receivedMonth}) — which month is it for?`,
    };
  }
  if (issued) return { month: issued.month, via: 'invoice-date', period: null, issued };
  return { month: null, via: 'unknown', period: null, issued: null, why: 'nothing in the invoice says which month it bills' };
}

// --- The run -----------------------------------------------------------------

function newSummary(extra) {
  return {
    filed: 0, alreadyPresent: 0, skippedNotInvoice: 0, held: 0,
    perApp: {}, heldItems: [], errors: [], timedOut: false,
    newFolders: [], reroutedByPeriod: [],
    amountsWritten: [], amountsUpdated: [], amountsNeedingReview: [], amountsUnread: [], amountsSkippedFilled: [],
    folderTotals: [], duplicateInvoices: [], looseIncluded: [], reconciled: [],
    answered: [], learned: [], ignored: [],
    ...(extra || {}),
  };
}

// Returns a summary of what was filed. `deadline` is a timestamp the run stops
// at, so the caller can share one function timeout across several jobs.
async function runMailSync(token, driveId, options) {
  const opts = options || {};
  const mailbox = mail.mailboxAddress();
  const ctx = await openPlacement(token, driveId, opts);
  const { root, base, sheet, appNames, archiveFolders, placeFor, vendorFolderFor, INDEX_PATH } = ctx;
  const STATE_PATH = archiveFile(root, '_mail-sync.json');
  const state = (await readJsonFile(token, driveId, STATE_PATH)) || {};
  const seen = new Set(Array.isArray(state.seenMessageIds) ? state.seenMessageIds : []);
  // Normal runs only look at mail newer than the last successful run. A rescan
  // has to widen the window too — skipping the seen-list alone achieves nothing,
  // because the Graph query would still be asking only for mail that arrived
  // after the last run.
  const lookback = new Date(Date.now() - FIRST_RUN_LOOKBACK_DAYS * 86400000).toISOString();
  const since = opts.rescan ? lookback : (state.lastRunAt || lookback);

  const rules = await readRules(token, driveId, root);
  const held = await pending.readPending(token, driveId, root);
  // The old fuzzy resolver is kept for one thing: a suggestion beside a
  // question. It never files anything.
  const suggest = buildResolver(await readAliasMap(token, driveId), appNames);

  // Invoices held because the PDF would not read are tried again first: the
  // reader has been fixed more than once, and a file the rules can now place
  // should be filed by the rules, not by an answer typed for want of them.
  const summary = newSummary({ mailbox, since, scanned: 0, base });
  try {
    const reread = await rereadHeld(token, driveId, { root, rules, appNames, held, deadline: opts.deadline || (Date.now() + DEADLINE_MS) });
    summary.reread = reread.items;
    if (reread.answers.length) {
      await pending.writePending(token, driveId, root, held);
      const filed = await resolvePending(token, driveId, reread.answers, { ...opts, root });
      summary.rereadFiled = filed.answered;
      for (const k of ['answered', 'amountsWritten', 'amountsUpdated', 'amountsSkippedFilled', 'amountsNeedingReview', 'amountsUnread', 'errors', 'learned']) {
        summary[k] = (summary[k] || []).concat(filed[k] || []);
      }
      summary.filed += filed.filed || 0;
      for (const [app, n] of Object.entries(filed.perApp || {})) summary.perApp[app] = (summary.perApp[app] || 0) + n;
      // The state on disk moved on; read it again before the mailbox pass adds to it.
      Object.assign(held, await pending.readPending(token, driveId, root));
    } else if (reread.changed) {
      await pending.writePending(token, driveId, root, held);
    }
  } catch (e) {
    summary.errors.push(`re-reading held invoices: ${e.message}`);
  }

  const messages = await mail.listMessages(token, mailbox, since.replace(/\.\d+Z$/, 'Z'), 50);
  summary.scanned = messages.length;
  for (const e of ctx.errors) summary.errors.push(e);

  const indexEntries = [];
  const marks = [];
  const markKeys = new Set();
  const touched = new Map(); // app||month -> { app, month, folder } seen this run
  const deadline = opts.deadline || (Date.now() + DEADLINE_MS);
  let heldChanged = false;

  for (const message of messages) {
    if (Date.now() > deadline) { summary.timedOut = true; break; }
    if (!opts.rescan && seen.has(message.id)) continue;

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
    const receivedMonth = mail.messageMonth(message);
    const pdfs = attachments.filter(mail.isPdf);
    const signals = {
      address: identity.address, originalAddress: identity.originalAddress, senderName: identity.name,
      subject: message.subject || '', attachmentNames: pdfs.map(a => a.name || ''),
    };

    // Read every PDF in the mail before placing any of them. Stripe sends an
    // invoice and its receipt in one mail; the receipt reads "Date paid" and
    // the invoice "Date of issue", and if one of them is read and the other
    // not, they must not part company — a receipt filed under September while
    // its invoice waits to be answered as August is one charge in two months.
    // Documents sharing an invoice number therefore share a month and an app,
    // and are held together when either is in doubt.
    const docs = [];
    for (const att of pdfs) {
      if (Date.now() > deadline) { summary.timedOut = true; break; }
      const fileName = sanitizeFileName(att.name);
      try {
        // Already held from an earlier run (a rescan re-reads the same mail):
        // the question stands, do not park a second copy.
        if (pending.findHeld(held, fileName, message.id)) continue;

        // Fetch even when the file is already archived: the bytes are what the
        // total is read from, and an invoice filed by hand still has an amount
        // worth picking up.
        const bytes = await mail.getAttachmentBytes(token, mailbox, message.id, att.id);

        // Read the PDF before choosing where it goes: the rules look at the
        // invoice's own wording, and the month comes from its billing period.
        const { text: pdfText, error: pdfError } = await readPdfText(bytes);
        const verdict = rulesLib.classify(rules, signals, pdfError ? '' : pdfText, appNames);
        const total = pdfError
          ? { amount: null, currency: null, usable: false, note: pdfError }
          : extractInvoiceTotal(pdfText, { currency: verdict.currency });
        const ref = pdfError ? null : extractInvoiceRef(pdfText);
        const when = pdfError
          ? { month: null, via: 'unreadable', period: null, issued: null, why: pdfError }
          : monthForInvoice(pdfText, receivedMonth, { usageRange: verdict.period === 'usage' });
        docs.push({ att, fileName, bytes, pdfError, verdict, total, ref, when, app: verdict.app || null, month: when.month, via: when.via });
      } catch (e) {
        summary.errors.push(`${fileName}: ${e.message}`);
      }
    }

    // Companions: same invoice number, one mail.
    for (const doc of docs) {
      if (!doc.ref) continue;
      const mates = docs.filter(d => d !== doc && d.ref === doc.ref);
      for (const mate of mates) {
        if (!doc.app && mate.app) doc.app = mate.app;
        if (!doc.month && mate.month) { doc.month = mate.month; doc.via = `${mate.via} (its ${/receipt/i.test(mate.fileName) ? 'receipt' : 'invoice'})`; }
      }
    }
    for (const doc of docs) {
      if (!doc.ref) continue;
      const group = docs.filter(d => d.ref === doc.ref);
      if (group.some(d => !d.app || !d.month)) for (const d of group) d.holdWithGroup = true;
    }

    for (const doc of docs) {
      if (Date.now() > deadline) { summary.timedOut = true; break; }
      const { att, fileName, bytes, pdfError, verdict, total, ref, when } = doc;
      try {
        const base_meta = {
          file: fileName, messageId: message.id, attachmentId: att.id,
          subject: message.subject || '', from: identity.originalAddress || identity.address,
          senderName: identity.name, domain: rulesLib.domainOf(identity.originalAddress || identity.address) || null,
          receivedAt: message.receivedDateTime, receivedMonth, webLink: message.webLink || null,
          month: doc.month, monthVia: doc.via,
          periodStart: when.period ? when.period.start : null, periodEnd: when.period ? when.period.end : null,
          invoiceDate: when.issued ? when.issued.date : null,
          amount: total.amount, currency: total.currency, amountUsable: !!total.usable, amountNote: total.note || '',
          ref, vendor: verdict.vendor || null,
        };

        // Not sure of the app, or of the month: hold it and ask. The two
        // questions are combined when both apply, so one answer settles it.
        if (!doc.app || !doc.month || doc.holdWithGroup) {
          const questions = [];
          let reason = doc.app ? 'no-month' : verdict.reason;
          if (!doc.app) questions.push(verdict.question);
          if (!doc.month) questions.push(`Which month is it for? (${when.why || 'nothing in the invoice says which month it bills'})`);
          if (doc.app && doc.month && doc.holdWithGroup) { reason = 'with-companion'; questions.push('Held with the other document in this mail that carries the same invoice number, so the two are filed together.'); }
          if (pdfError) reason = 'unreadable';
          const guess = !doc.app ? [signals.subject, signals.attachmentNames.join(' '), signals.senderName].map(t => suggest('', t)).find(h => h.app || h.suggestion) : null;
          const { item, already } = await pending.holdInvoice(token, driveId, root, held, bytes, {
            ...base_meta, app: doc.app || null, reason, question: questions.join(' '),
            options: verdict.options || [], suggestion: guess ? (guess.app || guess.suggestion) : null,
          });
          if (!already) { heldChanged = true; summary.held++; summary.heldItems.push(item); }
          continue;
        }

        const attApp = doc.app;
        const attMonth = doc.month;
        const attVendor = vendorFolderFor(attApp);
        if (!ctx.folderForApp[attApp] && !summary.newFolders.includes(attVendor)) summary.newFolders.push(attVendor);

        if (attMonth !== receivedMonth && !summary.reroutedByPeriod.some(r => r.file === fileName)) {
          // An earlier run — before this rule existed — may have filed the same
          // PDF under the mail's month. Nothing is deleted here, but a leftover
          // copy has to be reported: both folders would be totalled, and the
          // same charge would count in two months.
          const priorPlace = await placeFor(attVendor, receivedMonth);
          summary.reroutedByPeriod.push({
            file: fileName, app: attApp, from: receivedMonth, to: attMonth, via: doc.via,
            periodStart: base_meta.periodStart, periodEnd: base_meta.periodEnd, invoiceDate: base_meta.invoiceDate,
            alsoStillAt: priorPlace.files.has(fileName) ? priorPlace.folder : null,
          });
        }

        const place = await placeFor(attVendor, attMonth);
        const attFolder = place.folder;

        if (place.files.has(fileName)) {
          summary.alreadyPresent++;
        } else {
          await uploadFileContent(token, driveId, `${attFolder}/${fileName}`, bytes, 'application/pdf');
          // Remember it, so a second mail carrying the same attachment name in
          // the same run is reported as already present rather than re-uploaded.
          place.files.add(fileName);
          summary.filed++;
          summary.perApp[attApp] = (summary.perApp[attApp] || 0) + 1;
        }

        // Tick the tracker and total the folder for the app and month this PDF
        // was actually filed under. Several invoices for one app-month are
        // normal; they tick one cell, not one per message.
        const markKey = `${attApp}||${attMonth}`;
        if (!markKeys.has(markKey)) { markKeys.add(markKey); marks.push({ app: attApp, month: attMonth }); }
        // The amount is NOT summed here: the whole folder is totalled after the
        // loop, so invoices that arrived by any other route count too.
        touched.set(markKey, { app: attApp, month: attMonth, folder: attFolder });

        if (total.amount !== null && !total.usable) {
          summary.amountsNeedingReview.push({ app: attApp, month: attMonth, file: fileName, amount: total.amount, currency: total.currency, note: total.note });
        } else if (total.amount === null) {
          summary.amountsUnread.push({ app: attApp, month: attMonth, file: fileName, note: total.note });
        }

        indexEntries.push({
          app: attApp, month: attMonth, file: fileName, folder: attFolder,
          amount: total.amount, currency: total.currency, amountUsable: !!total.usable, ref,
          monthVia: doc.via, periodStart: base_meta.periodStart, periodEnd: base_meta.periodEnd, invoiceDate: base_meta.invoiceDate,
          receivedMonth, matchedVia: verdict.via, vendor: verdict.vendor,
          subject: message.subject || '', from: identity.originalAddress || identity.address,
          receivedAt: message.receivedDateTime, webLink: message.webLink || null,
        });
      } catch (e) {
        summary.errors.push(`${fileName}: ${e.message}`);
      }
    }
    seen.add(message.id);
  }

  if (heldChanged) await pending.writePending(token, driveId, root, held);

  // Sweep the archive for recent months too, so an invoice that arrived by any
  // route other than this mailbox is totalled and ticked like the rest. Folded
  // into `touched` and `marks` rather than given a path of its own: everything
  // downstream — the folder total, the cell, the tracker tick — then treats it
  // identically, and there is one set of rules to get right.
  if (!opts.skipReconcile) {
    const byNorm = new Map();
    for (const name of appNames) byNorm.set(normName(name), name);
    const appFor = (folder) => {
      if (folder === UNMATCHED_FOLDER || String(folder).startsWith('_')) return null;
      const direct = byNorm.get(normName(folder));
      if (direct) return direct;
      const mapped = opts.mapping && opts.mapping[folder];
      if (mapped && byNorm.get(normName(mapped))) return byNorm.get(normName(mapped));
      const hit = suggest(folder, '');
      return hit.confident ? hit.app : null;
    };
    try {
      const swept = await reconcileArchive(token, driveId, base, {
        folders: archiveFolders, appFor, deadline: Math.min(deadline, Date.now() + 15 * 1000),
      });
      for (const e of swept.errors) summary.errors.push(`archive sweep: ${e}`);
      if (swept.timedOut) summary.reconcileTimedOut = true;
      for (const { app, month, folder } of swept.found) {
        const key = `${app}||${month}`;
        const at = touched.get(key);
        if (!at) {
          touched.set(key, { app, month, folder });
          summary.reconciled.push({ app, month });
        } else if (at.folder !== folder && !(at.alsoFolders || []).includes(folder)) {
          // A vendor can name one month twice — Luzmo has "Jul-26" and a "July"
          // lifted out of a nested "Luzmo/". Totalling only the first reports
          // half the month's spend as if it were all of it.
          (at.alsoFolders || (at.alsoFolders = [])).push(folder);
        }
        if (!markKeys.has(key)) { markKeys.add(key); marks.push({ app, month }); }
      }
    } catch (e) {
      summary.errors.push(`archive sweep: ${e.message}`);
    }
  }

  await settleMonths(token, driveId, { sheet, INDEX_PATH, deadline, locks: rules.locks }, touched, marks, indexEntries, summary);

  await writeJsonFile(token, driveId, STATE_PATH, {
    mailbox,
    // On a timed-out run, don't advance the watermark past what was processed —
    // the next run picks the remainder up.
    lastRunAt: summary.timedOut ? state.lastRunAt || since : new Date().toISOString(),
    seenMessageIds: [...seen].slice(-500),
    updatedAt: new Date().toISOString(),
  });

  // One PDF can be read twice in a run — once as the mail attachment, once when
  // its folder is totalled — so a single bad file was reported as two problems.
  // Report each file once, keeping the first reason given.
  summary.amountsUnread = dedupeByFile(summary.amountsUnread);
  summary.amountsNeedingReview = dedupeByFile(summary.amountsNeedingReview);
  summary.pendingCount = held.items.length;

  return { ok: true, ranAt: new Date().toISOString(), ...summary };
}

// Reads held-as-unreadable PDFs again with the current reader. An item that
// now yields an app and a month becomes an answer (filed by the rules, so
// nothing is guessed); one that reads but still lacks either has its question
// and figures refreshed so the owner sees what was read. Bounded, because it
// shares the run's time with the mailbox.
const REREAD_MAX = 12;

async function rereadHeld(token, driveId, ctx) {
  const { root, rules, appNames, held, deadline } = ctx;
  const answers = [];
  const items = [];
  let changed = false;
  const candidates = held.items.filter(i => i.reason === 'unreadable' || /could not read the pdf/i.test(String(i.question || ''))).slice(0, REREAD_MAX);
  for (const item of candidates) {
    if (Date.now() > deadline - 10 * 1000) break;
    try {
      const res = await graphFetch(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(item.heldPath)}:/content`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!res.ok) throw new Error(`download ${res.status}`);
      const bytes = Buffer.from(await res.arrayBuffer());
      const { text, error } = await readPdfText(bytes);
      if (error) { items.push({ id: item.id, file: item.file, result: 'still unreadable', note: error }); continue; }

      const signals = { address: item.from, originalAddress: item.from, senderName: item.senderName, subject: item.subject, attachmentNames: [item.file] };
      const verdict = rulesLib.classify(rules, signals, text, appNames);
      const total = extractInvoiceTotal(text, { currency: verdict.currency });
      const when = monthForInvoice(text, item.receivedMonth, { usageRange: verdict.period === 'usage' });
      Object.assign(item, {
        amount: total.amount, currency: total.currency, amountUsable: !!total.usable, amountNote: total.note || '',
        ref: extractInvoiceRef(text), month: when.month, monthVia: when.via,
        periodStart: when.period ? when.period.start : null, periodEnd: when.period ? when.period.end : null,
        invoiceDate: when.issued ? when.issued.date : null,
        app: verdict.app || item.app || null, vendor: verdict.vendor || item.vendor || null, options: verdict.options || item.options || [],
      });
      changed = true;
      if (item.app && item.month) {
        answers.push({ id: item.id, app: item.app, month: item.month });
        items.push({ id: item.id, file: item.file, result: 'read and filed', app: item.app, month: item.month, amount: total.amount });
      } else {
        const questions = [];
        item.reason = item.app ? 'no-month' : (verdict.reason || item.reason);
        if (!item.app) questions.push(verdict.question);
        if (!item.month) questions.push(`Which month is it for? (${when.why || 'nothing in the invoice says which month it bills'})`);
        item.question = questions.join(' ');
        items.push({ id: item.id, file: item.file, result: 'read, still a question', question: item.question });
      }
    } catch (e) {
      items.push({ id: item.id, file: item.file, result: 'could not re-read', note: e.message });
    }
  }
  return { answers, items, changed };
}

// Totals every app-month that this run touched, sets the sheet cells to those
// totals, ticks the tracker, and records what was parsed and written. Shared
// by the mailbox pass and by answering a held question, so a file reaches the
// sheet by one path whichever way it arrived.
async function settleMonths(token, driveId, ctx, touched, marks, indexEntries, summary) {
  const { sheet, INDEX_PATH, deadline } = ctx;
  const amounts = {};        // app -> month -> summed USD invoice total
  const partial = {};        // app||month -> true when some PDF in that month could not be used
  let tracker = { sheet: null, marked: 0 };

  // Total each app-month folder that this run touched. Reuse whatever the index
  // already knows so a PDF is parsed once, not on every run.
  if (touched.size) {
    const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
    const cache = new Map();
    for (const e of (Array.isArray(index.amounts) ? index.amounts : [])) {
      if (e && e.path) cache.set(e.path, { amount: e.amount, currency: e.currency, usable: !!e.usable, note: e.note || '', ref: e.ref || null });
    }
    const budget = { parsed: 0, maxParse: 40, deadline, fresh: [], exhausted: false };

    for (const { app, month, folder, alsoFolders, emptied } of touched.values()) {
      try {
        // The vendor folder holding this month's folder — where loose invoices
        // for the same month sit.
        const appFolder = folder.slice(0, folder.lastIndexOf('/'));
        let loose = [];
        try {
          loose = await looseFilesForMonth(token, driveId, appFolder, month);
        } catch (e) {
          summary.errors.push(`${app} ${month}: could not check "${appFolder}" for loose invoices (${e.message})`);
        }
        // A second folder naming the same month is totalled with the first, the
        // same way loose files are: the month's spend is every invoice for it,
        // wherever it happens to be filed. Duplicates are still recognised by
        // invoice number, so a file copied into both is counted once.
        for (const other of alsoFolders || []) {
          try {
            const { files } = await folderChildren(token, driveId, other);
            for (const name of files) {
              if (/\.pdf$/i.test(name)) loose.push({ name, path: `${other}/${name}` });
            }
          } catch (e) {
            summary.errors.push(`${app} ${month}: could not read "${other}" (${e.message})`);
          }
        }
        const res = await sumFolderInvoices(token, driveId, folder, cache, budget, loose);
        if (loose.length) summary.looseIncluded.push({ app, month, files: loose.map(f => f.name) });
        if (res.counted > 0 || (emptied && res.pdfCount === 0)) {
          const byMonth = amounts[app] || (amounts[app] = {});
          byMonth[month] = res.total;
          // A month with PDFs nobody could use has not been totalled, only
          // partially added up. Carry that through so the writer holds it.
          partial[`${app}||${month}`] = res.counted < res.pdfCount;
        }
        summary.folderTotals.push({
          app, month, folder, total: res.total,
          invoicesCounted: res.counted, pdfsInFolder: res.pdfCount,
        });
        for (const d of res.duplicates) summary.duplicateInvoices.push({ app, month, ...d });
        for (const u of res.unusable) summary.amountsNeedingReview.push({ app, month, ...u });
        for (const u of res.unread) summary.amountsUnread.push({ app, month, ...u });
      } catch (e) {
        summary.errors.push(`${app} ${month}: could not total the folder (${e.message})`);
      }
      if (budget.exhausted) { summary.timedOut = true; break; }
    }

    // Persist the freshly parsed totals so the next run does not re-download them.
    if (budget.fresh.length || indexEntries.length) {
      const freshPaths = new Set(budget.fresh.map(f => f.path));
      const existing = (Array.isArray(index.amounts) ? index.amounts : []).filter(e => !freshPaths.has(e.path));
      // Spread the index rather than rebuilding it: it also carries `written`
      // (which figures this sync put in the sheet) and `periods` (every billing
      // period read out of a PDF).
      await writeJsonFile(token, driveId, INDEX_PATH, {
        ...index,
        entries: (Array.isArray(index.entries) ? index.entries : []).concat(indexEntries).slice(-500),
        amounts: existing.concat(budget.fresh).slice(-2000),
        updatedAt: new Date().toISOString(),
      });
      indexEntries.length = 0; // already written above
    }
  }

  // Tick the invoice tracker and set the cells. Locked cells are checked on
  // every run, whether or not their month was touched.
  const hasAmounts = Object.keys(amounts).length > 0 || (Array.isArray(ctx.locks) && ctx.locks.length > 0);
  if (marks.length || hasAmounts) {
    let sessionId = null;
    try {
      sessionId = await excel.createSession(token, sheet.driveId, sheet.itemId);

      if (marks.length) {
        try {
          tracker = await markTracker(token, sheet.driveId, sheet.itemId, marks, sessionId);
        } catch (e) {
          summary.errors.push(`tracker update: ${e.message}`);
        }
      }

      if (hasAmounts) {
        try {
          // Re-read inside the session so the row/column positions match the copy
          // being edited — the sheet gains rows from time to time.
          const live = await openSpendSheet(token, sessionId);
          const priorIndex = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
          const priorWrites = (priorIndex && priorIndex.written) || {};
          const planned = planAmountCells(amounts, live.grid, { values: live.values, start: live.start }, cellValue, priorWrites, partial, ctx.locks);
          summary.amountsSkippedFilled = (summary.amountsSkippedFilled || []).concat(planned.skippedFilled);
          summary.amountsLocked = (summary.amountsLocked || []).concat(planned.locked);
          const toWrite = planned.write.concat(planned.updated);
          if (toWrite.length) {
            const results = await excel.writeCells(token, live.driveId, live.itemId, live.sheetName, toWrite, sessionId, 4);
            const ok = results.filter(r => r.ok);
            for (const r of ok) {
              const u = planned.updated.find(x => x.address === r.address);
              if (u) {
                summary.amountsUpdated.push({
                  app: r.app, month: r.month, address: r.address, amount: r.value,
                  previous: u.previous, direction: u.direction, emptied: !!u.emptied, locked: !!u.locked,
                  // false when the figure being replaced was not one this sync
                  // wrote — a hand correction or a statement figure. Surfaced so
                  // replacing someone's number is never silent.
                  wasOurs: !!u.wasOurs,
                });
              } else {
                summary.amountsWritten.push({ app: r.app, month: r.month, address: r.address, amount: r.value });
              }
            }
            // Remember what we wrote, so a later run can tell our own figure
            // from one a human has since corrected.
            const written = { ...priorWrites };
            for (const r of ok) written[`${r.app}||${r.month}`] = r.value;
            await writeJsonFile(token, driveId, INDEX_PATH, { ...priorIndex, written, updatedAt: new Date().toISOString() });
            for (const bad of results.filter(r => !r.ok)) summary.errors.push(`${bad.app} ${bad.month}: ${bad.error}`);
            if (ok.length) {
              await appendLog(token, driveId, {
                at: new Date().toISOString(), by: 'invoice-sync', source: 'invoice',
                statement: null, attribution: 'invoice', sheet: live.sheetName,
                cells: ok.map(c => ({ app: c.app, month: c.month, address: c.address, before: c.before === undefined ? null : c.before, after: c.value })),
                failed: [],
              });
            }
          }
        } catch (e) {
          summary.errors.push(`amount update: ${e.message}`);
        }
      }
    } catch (e) {
      summary.errors.push(`workbook session: ${e.message}`);
    } finally {
      if (sessionId) await excel.closeSession(token, sheet.driveId, sheet.itemId, sessionId);
    }
  }

  if (indexEntries.length) {
    const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
    const merged = (Array.isArray(index.entries) ? index.entries : []).concat(indexEntries).slice(-500);
    await writeJsonFile(token, driveId, INDEX_PATH, { ...index, entries: merged, updatedAt: new Date().toISOString() });
    indexEntries.length = 0;
  }

  summary.tracker = tracker;
  return summary;
}

// The other documents this mail carried, as the index recorded them: same
// subject and arrival time, filed somewhere other than where the answer puts
// this one. Entries are returned live so the caller can update them in place.
function companionsOf(index, item, destFolder) {
  if (!item || !item.subject || !item.receivedAt) return [];
  return (Array.isArray(index.entries) ? index.entries : []).filter(e =>
    e && e.subject === item.subject && e.receivedAt === item.receivedAt
    && e.file !== item.file && e.folder && e.folder !== destFolder)
    .map(e => { e.movedFrom = `${e.folder}/${e.file}`; return e; });
}

// --- Answering a held question ----------------------------------------------

// `answers` is a list of { id, app, month, amount, ignore }, from the dashboard
// or parsed out of a Slack reply. Each one files the held PDF where the answer
// says, remembers the answer as a rule where one can be inferred, and then
// totals and ticks the month exactly as a mailed invoice would be.
async function resolvePending(token, driveId, answers, options) {
  const opts = options || {};
  const ctx = await openPlacement(token, driveId, opts);
  const { root, base, sheet, appNames, placeFor, vendorFolderFor, INDEX_PATH } = ctx;
  const held = await pending.readPending(token, driveId, root);
  let rules = await readRules(token, driveId, root);
  const summary = newSummary({ base });
  for (const e of ctx.errors) summary.errors.push(e);

  const touched = new Map();
  const marks = [];
  const indexEntries = [];
  const manualAmounts = [];
  let rulesChanged = false;
  let stateChanged = false;
  let indexChanged = false;
  const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};

  for (const ans of Array.isArray(answers) ? answers : []) {
    if (!ans || !ans.id) continue;
    const item = held.items.find(i => i.id === String(ans.id).trim().toUpperCase());
    if (!item) { summary.errors.push(`${ans.id}: not an open question`); continue; }
    try {
      if (ans.ignore) {
        const dest = `${base}/_Ignored`;
        await pending.moveHeld(token, driveId, item, dest);
        pending.removeItem(held, item.id);
        stateChanged = true;
        summary.ignored.push({ id: item.id, file: item.file });
        continue;
      }
      const app = pending.matchApp(ans.app || item.app || '', appNames);
      if (!app.app) { summary.errors.push(`${item.id}: ${app.error}`); continue; }
      const month = ans.month || item.month || null;
      if (!month || !/^\d{4}-\d{2}$/.test(month)) { summary.errors.push(`${item.id}: which month? (say e.g. "${item.id} = ${app.app}, Aug-26")`); continue; }

      const vendorFolder = vendorFolderFor(app.app);
      const place = await placeFor(vendorFolder, month);
      let path;
      if (place.files.has(item.file)) {
        // Already there — filed by hand meanwhile. The held copy is not needed.
        path = `${place.folder}/${item.file}`;
        await pending.moveHeld(token, driveId, item, `${base}/_Ignored`).catch(() => null);
        summary.alreadyPresent++;
      } else {
        path = await pending.moveHeld(token, driveId, item, place.folder);
        place.files.add(item.file);
        summary.filed++;
        summary.perApp[app.app] = (summary.perApp[app.app] || 0) + 1;
      }
      pending.removeItem(held, item.id);
      stateChanged = true;
      summary.answered.push({ id: item.id, file: item.file, app: app.app, month, folder: place.folder });

      // A companion from the same mail — the receipt for this invoice, or the
      // invoice for this receipt — that an earlier run filed under another
      // month follows the answer, and the month it leaves is totalled again
      // so its cell does not keep the charge.
      for (const mate of companionsOf(index, item, place.folder)) {
        try {
          const from = `${mate.folder}/${mate.file}`;
          const itemId = await itemIdByPath(token, driveId, from);
          if (!itemId) continue;
          if (place.files.has(mate.file)) continue;
          const parentId = await ensureFolder(token, driveId, place.folder);
          await moveItem(token, driveId, itemId, parentId, mate.file);
          place.files.add(mate.file);
          const oldKey = `${mate.app}||${mate.month}`;
          if (!touched.has(oldKey)) { touched.set(oldKey, { app: mate.app, month: mate.month, folder: mate.folder, emptied: true }); }
          mate.folder = place.folder; mate.month = month; mate.app = app.app; mate.monthVia = 'answered (with its companion)';
          indexChanged = true;
          summary.answered.push({ id: item.id, file: mate.file, app: app.app, month, folder: place.folder, companion: true });
        } catch (e) {
          summary.errors.push(`${item.id}: could not move ${mate.file} alongside it (${e.message})`);
        }
      }

      // An amount typed with the answer stands in for one the PDF would not
      // give up. It is recorded against the file, as if it had been read.
      const typed = ans.amount === undefined || ans.amount === null || ans.amount === '' ? null : Number(ans.amount);
      if (typed !== null && Number.isFinite(typed) && typed >= 0) {
        manualAmounts.push({ path, amount: typed, currency: 'USD', usable: true, ref: item.ref || null, note: 'typed by the owner when answering', manual: true });
      }

      const learned = rulesLib.learn(rules, { ...item, senderName: item.senderName }, app.app);
      if (learned.learned) { rules = learned.rules; rulesChanged = true; summary.learned.push(learned.learned); }

      const key = `${app.app}||${month}`;
      if (!touched.has(key)) { touched.set(key, { app: app.app, month, folder: place.folder }); marks.push({ app: app.app, month }); }
      indexEntries.push({
        app: app.app, month, file: item.file, folder: place.folder,
        amount: typed !== null ? typed : item.amount, currency: typed !== null ? 'USD' : item.currency, amountUsable: typed !== null ? true : !!item.amountUsable,
        monthVia: ans.month ? 'answered' : item.monthVia, periodStart: item.periodStart, periodEnd: item.periodEnd,
        receivedMonth: item.receivedMonth, matchedVia: 'answered', vendor: item.vendor || null,
        subject: item.subject, from: item.from, receivedAt: item.receivedAt, webLink: item.webLink || null,
      });
    } catch (e) {
      summary.errors.push(`${item.id}: ${e.message}`);
    }
  }

  if (manualAmounts.length || indexChanged) {
    const amountsList = (Array.isArray(index.amounts) ? index.amounts : []).filter(e => !manualAmounts.some(m => m.path === e.path));
    // A moved companion's parse record is keyed by its old path; drop it so the
    // new path is read afresh rather than reported missing.
    const movedFrom = new Set((index.entries || []).filter(e => e.monthVia === 'answered (with its companion)').map(e => e.movedFrom).filter(Boolean));
    await writeJsonFile(token, driveId, INDEX_PATH, {
      ...index,
      amounts: amountsList.filter(e => !movedFrom.has(e.path)).concat(manualAmounts).slice(-2000),
      updatedAt: new Date().toISOString(),
    });
  }
  if (rulesChanged) await writeRules(token, driveId, root, rules);
  if (stateChanged) await pending.writePending(token, driveId, root, held);

  await settleMonths(token, driveId, { sheet, INDEX_PATH, deadline: opts.deadline || (Date.now() + DEADLINE_MS), locks: rules.locks }, touched, marks, indexEntries, summary);
  summary.amountsUnread = dedupeByFile(summary.amountsUnread);
  summary.amountsNeedingReview = dedupeByFile(summary.amountsNeedingReview);
  summary.pendingCount = held.items.length;
  return { ok: true, ranAt: new Date().toISOString(), ...summary };
}

// What the owner has typed into the Slack DM since the last look, applied as
// answers. Returns the resolve summary plus the lines that could not be read.
async function collectSlackAnswers(token, driveId, options) {
  const opts = options || {};
  if (!slack.configured()) return null;
  const root = opts.root || await resolveArchiveRoot(token, driveId);
  const held = await pending.readPending(token, driveId, root);
  const replies = await slack.readReplies(held.slack.lastTs);
  if (!replies.messages.length) {
    if (replies.latestTs && replies.latestTs !== held.slack.lastTs) {
      held.slack.lastTs = replies.latestTs;
      await pending.writePending(token, driveId, root, held);
    }
    return null;
  }
  const sheet = await openSpendSheet(token);
  const appNames = sheet.grid.apps.map(a => a.name);
  const answers = [];
  const unreadable = [];
  for (const m of replies.messages) {
    for (const parsed of pending.parseReply(m.text, held.items, appNames)) {
      if (parsed.error) unreadable.push({ line: parsed.line, error: parsed.error });
      else answers.push(parsed);
    }
  }
  // Mark the DM read before acting, so a failure part-way does not replay the
  // same answers on the next run — re-answering a closed question is refused
  // anyway, but there is no point trying.
  held.slack.lastTs = replies.latestTs;
  await pending.writePending(token, driveId, root, held);

  const result = answers.length ? await resolvePending(token, driveId, answers, { ...opts, root }) : null;
  return { answers: answers.length, unreadable, result };
}

// --- Telling the owner -------------------------------------------------------

const money = n => Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// The DM. One message per run that changed anything or has a question
// outstanding; a run that filed nothing, wrote nothing and asks nothing says
// nothing. Slack mrkdwn, not Markdown: *bold*, • bullets.
function formatRunReport(summary, heldItems, extra) {
  const s = summary || {};
  const lines = [];
  const changed = (s.filed || 0) + (s.amountsWritten || []).length + (s.amountsUpdated || []).length
    + ((s.tracker && s.tracker.marked) || 0) + (s.answered || []).length + (s.ignored || []).length;
  const asking = (heldItems || []).length;
  if (!changed && !asking && !(s.errors || []).length) return null;

  lines.push(`*Invoice sync — ${new Date().toISOString().slice(0, 16).replace('T', ' ')} UTC*`);
  if (s.filed) {
    lines.push(`Filed ${s.filed} invoice${s.filed === 1 ? '' : 's'}:`);
    for (const [app, n] of Object.entries(s.perApp || {})) lines.push(`• ${app}: ${n}`);
  }
  for (const a of s.answered || []) lines.push(`• ${a.id} ${a.file} → ${a.app} ${a.month}${(s.rereadFiled || []).some(r => r.id === a.id) ? ' (read again and filed by the rules)' : ' (your answer)'}`);
  for (const i of s.ignored || []) lines.push(`• ${i.id} ${i.file} left out, as you said`);
  for (const l of s.learned || []) lines.push(`• remembered: ${l}`);
  const wrote = s.amountsWritten || [];
  const upd = s.amountsUpdated || [];
  if (wrote.length || upd.length) {
    lines.push(`Spendings sheet:`);
    for (const w of wrote) lines.push(`• ${w.app} ${w.month}: set to ${money(w.amount)}`);
    for (const u of upd) lines.push(`• ${u.app} ${u.month}: ${money(u.previous == null ? 0 : u.previous)} → ${money(u.amount)}${u.locked ? ' (restored to the figure you locked)' : u.emptied ? ' (its invoices moved to another month)' : u.direction === 'down' ? ' (lowered — are all that month\'s invoices in the folder?)' : ''}${u.wasOurs || u.locked ? '' : ' (replaced a figure not from invoices)'}`);
  }
  const kept = (s.amountsLocked || []).filter(l => !l.enforced && l.invoiceTotal !== undefined && Math.abs((l.invoiceTotal || 0) - l.value) >= 0.005);
  if (kept.length) {
    lines.push(`Locked cells left as you set them (the invoices on file say otherwise, and were ignored):`);
    for (const l of kept) lines.push(`• ${l.app} ${l.month}: ${money(l.value)} kept; invoices come to ${money(l.invoiceTotal)}`);
  }
  if (s.tracker && s.tracker.marked) lines.push(`Ticked ${s.tracker.marked} cell${s.tracker.marked === 1 ? '' : 's'} in the invoice tracker.`);
  const heldCells = (s.amountsSkippedFilled || []).filter(f => f.reason === 'folder-total-incomplete');
  if (heldCells.length) {
    lines.push(`Not written — an invoice in the month could not be used, so the total is not known:`);
    for (const f of heldCells) lines.push(`• ${f.app} ${f.month}: readable invoices come to ${money(f.invoiceTotal)}`);
  }
  const review = (s.amountsNeedingReview || []);
  if (review.length) {
    lines.push(`Invoices whose amount cannot go in the USD sheet as-is:`);
    for (const r of review.slice(0, 10)) lines.push(`• ${r.app} ${r.month || ''} ${r.file}: ${r.amount !== undefined && r.amount !== null ? money(r.amount) : '?'} ${r.currency || ''} — ${r.note || ''}`);
  }
  const unread = (s.amountsUnread || []);
  if (unread.length) {
    lines.push(`Invoices with no readable total:`);
    for (const u of unread.slice(0, 10)) lines.push(`• ${u.app || '?'} ${u.month || ''} ${u.file} — ${u.note || ''}`);
  }
  const rerouted = s.reroutedByPeriod || [];
  if (rerouted.length) {
    lines.push(`Filed by billing period rather than the month the mail arrived:`);
    for (const r of rerouted.slice(0, 10)) lines.push(`• ${r.app} ${r.file}: ${r.from} → ${r.to}${r.alsoStillAt ? ` (a copy is still in ${r.alsoStillAt} — delete it or the month counts twice)` : ''}`);
  }
  if (extra && extra.unreadable && extra.unreadable.length) {
    lines.push(`Replies I could not read:`);
    for (const u of extra.unreadable.slice(0, 10)) lines.push(`• "${u.line}" — ${u.error}`);
  }
  if ((s.errors || []).length) {
    lines.push(`Errors:`);
    for (const e of s.errors.slice(0, 8)) lines.push(`• ${e}`);
  }
  if (s.timedOut) lines.push(`Ran out of time part-way; the next run picks up the rest.`);
  if (asking) {
    lines.push('');
    lines.push(pending.describeQuestions(heldItems));
  }
  return lines.join('\n');
}

// Posts the report, never throwing: the run's result stands whether or not
// the DM got through, and the summary says which.
async function notify(token, driveId, summary, extra, root) {
  try {
    const r = root || await resolveArchiveRoot(token, driveId);
    const held = await pending.readPending(token, driveId, r);
    const text = formatRunReport(summary, held.items, extra);
    if (!text) return { sent: false, why: 'nothing to report' };
    return await slack.postDm(text);
  } catch (e) {
    return { sent: false, why: e.message };
  }
}

function dedupeByFile(items) {
  const seen = new Set();
  const out = [];
  for (const item of items || []) {
    const key = `${item.app || ''}||${item.month || ''}||${item.file || ''}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(item);
  }
  return out;
}

module.exports = { runMailSync, resolvePending, collectSlackAnswers, notify, formatRunReport, readRules, writeRules, monthForInvoice, companionsOf, rereadHeld, appToSourceFolder, reconcileArchive, recentMonths, RECONCILE_MONTHS, monthFolderName, planTrackerCells, planAmountCells, folderChildren, sumFolderInvoices };
