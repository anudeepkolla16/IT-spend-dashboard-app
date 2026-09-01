const { encodeGraphPath, sanitizeSegment, readJsonFile, writeJsonFile, uploadFileContent, resolveArchiveRoot, archiveFile, graphFetch } = require('./graph');
const excel = require('./excel');
const { openSpendSheet, readAliasMap, cellValue, appendLog } = require('./spend-sheet');
const { buildResolver, refineAnthropic } = require('./vendor-map');
const mail = require('./mail');
const { extractInvoiceTotal, extractInvoiceRef, readPdfText } = require('./invoice-amount');
const { invoiceMonth } = require('./invoice-period');
const { monthFromFileName, detectDayFirst, parseMonthFolder } = require('./invoices/inventory');
// The tracker write is shared with the checklist's backfill, which ticks what
// the archive already holds rather than what a run just filed. Same rules.
const { planTrackerCells, markTracker } = require('./invoices/tracker');

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
        const { text, error } = await readPdfText(bytes);
        entry = error ? { amount: null, currency: null, usable: false, note: error } : extractInvoiceTotal(text);
        // Kept so an invoice and its own payment receipt can be recognised as
        // one charge rather than two.
        if (!error) entry.ref = extractInvoiceRef(text);
      } catch (e) {
        entry = { amount: null, currency: null, usable: false, note: e.message };
      }
      cache.set(key, entry);
      budget.fresh.push({ path: key, amount: entry.amount, currency: entry.currency, usable: !!entry.usable, ref: entry.ref || null });
    }

    if (entry.usable) {
      // An invoice and its payment receipt are one charge sent twice. They carry
      // the same invoice number, so the second one seen is a duplicate of the
      // first, not more spend. Only a reference actually read counts: without
      // one, two files that merely happen to share an amount are two charges.
      if (entry.ref && seenRefs.has(entry.ref)) {
        duplicates.push({ file: name, amount: entry.amount, ref: entry.ref, of: seenRefs.get(entry.ref) });
        continue;
      }
      if (entry.ref) seenRefs.set(entry.ref, name);
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
// The model: a cell holds the invoices on file PLUS whatever spend the sheet
// already knows about that no invoice has explained yet. An empty cell takes the
// folder total; a cell above the folder total is left alone, because that excess
// is anticipating invoices still to come and a newly arrived one is absorbed by
// it rather than added on top; a cell below it has been overtaken by real spend
// and is raised. Which is to say `max(cell, folderTotal)`, and never a sum — the
// daily cron would grow an additive figure on every run.
//
// `amount` is the total of the WHOLE app-month folder, not one invoice, which is
// what makes that safe: the folder total is the entire figure, so writing it
// twice writes the same number twice.
const CELL_EPSILON = 0.005;

function planAmountCells(amounts, grid, used, cellValueOf, priorWrites, partialTotals) {
  const rowByApp = new Map(grid.apps.map(a => [a.name, a.rowIdx]));
  const prior = priorWrites || {};
  const partial = partialTotals || {};
  const write = [];
  const updated = [];
  const skippedFilled = [];

  for (const [app, byMonth] of Object.entries(amounts || {})) {
    const rowIdx = rowByApp.get(app);
    if (rowIdx === undefined) continue;
    for (const [month, amount] of Object.entries(byMonth)) {
      const colIdx = grid.monthCols[month];
      if (colIdx === undefined || !(amount > 0)) continue;
      const current = cellValueOf(used.values, rowIdx, colIdx);
      const address = excel.cellAddress(used.start, rowIdx, colIdx);
      // Some PDF in this folder could not be read, so `amount` is not the
      // folder's total — it is everything that happened to parse, and the real
      // figure is at least that and probably more.
      const isPartial = partial[`${app}||${month}`] === true;

      if (current === null || current === 0) {
        write.push({ app, month, address, value: amount, partial: isPartial });
        continue;
      }

      // The comparison below asks "has the folder grown past the sheet?", and a
      // lower bound cannot answer it. Apollo's August folder held two invoices,
      // one of them unreadable; the 85.00 that did parse looked bigger than the
      // 53.12 already in the cell, so it replaced a figure this sync had never
      // written with a half-read one. A partial total may fill an empty cell,
      // where something beats nothing and the gap is reported — but it may
      // never overwrite a figure somebody else put there.
      if (isPartial) {
        skippedFilled.push({ app, month, address, current, invoiceTotal: amount, reason: 'folder-total-incomplete' });
        continue;
      }

      // Invoices for a month arrive across it, so the cell has to keep up: when
      // the folder total has grown, a new invoice has landed and the figure is
      // updated whoever wrote it.
      if (amount > current + CELL_EPSILON) {
        const mine = prior[`${app}||${month}`];
        const wasOurs = mine !== undefined && Math.abs(current - mine) < CELL_EPSILON;
        updated.push({ app, month, address, value: amount, previous: current, wasOurs });
        continue;
      }

      // A total LOWER than the cell is the one case left alone, and this is the
      // deliberate top-up model rather than mere caution.
      //
      // A cell holds the invoices on file PLUS whatever spend the sheet already
      // knows about that no invoice has explained yet. When a new invoice
      // arrives it is ABSORBED by that excess rather than added on top of it,
      // because the excess was anticipating it: Bubble's cell held 524.27
      // against eight invoices totalling 492.27, and the ninth was 32.00 —
      // 492.27 + 32 = 524.27. The sheet was never wrong, only early. Adding
      // would have made it 556.27.
      //
      // Written out, that is `remainder_after + folderTotal_after`, which is
      // exactly `max(cell, folderTotal)` — so the cell simply never moves until
      // the folder passes it. It also makes the write idempotent, which no
      // additive rule could be: the daily cron would grow the figure every run.
      // See the "top-up model" tests in test/invoice-amount.test.js.
      if (amount < current - CELL_EPSILON) {
        skippedFilled.push({ app, month, address, current, invoiceTotal: amount, reason: 'invoice-total-lower' });
      }
    }
  }
  return { write, updated, skippedFilled };
}

// Returns a summary of what was filed. `deadline` is a timestamp the run stops
// at, so the caller can share one function timeout across several jobs.
async function runMailSync(token, driveId, options) {
  const opts = options || {};
  const mailbox = mail.mailboxAddress();
  const root = await resolveArchiveRoot(token, driveId);
  const base = root.path;
  const STATE_PATH = archiveFile(root, '_mail-sync.json');
  const INDEX_PATH = archiveFile(root, '_invoice-index.json');
  const state = (await readJsonFile(token, driveId, STATE_PATH)) || {};
  const seen = new Set(Array.isArray(state.seenMessageIds) ? state.seenMessageIds : []);
  // Normal runs only look at mail newer than the last successful run. A rescan
  // has to widen the window too — skipping the seen-list alone achieves nothing,
  // because the Graph query would still be asking only for mail that arrived
  // after the last run.
  const lookback = new Date(Date.now() - FIRST_RUN_LOOKBACK_DAYS * 86400000).toISOString();
  const since = opts.rescan ? lookback : (state.lastRunAt || lookback);

  const sheet = await openSpendSheet(token);
  const appNames = sheet.grid.apps.map(a => a.name);
  const resolve = buildResolver(await readAliasMap(token, driveId), appNames);

  const messages = await mail.listMessages(token, mailbox, since.replace(/\.\d+Z$/, 'Z'), 50);

  const summary = {
    mailbox, since, scanned: messages.length,
    filed: 0, alreadyPresent: 0, skippedNotInvoice: 0, unmatched: 0,
    perApp: {}, unmatchedItems: [], errors: [], timedOut: false,
    base, newFolders: [], reroutedByContent: [], reroutedByPeriod: [],
    amountsWritten: [], amountsUpdated: [], amountsNeedingReview: [], amountsUnread: [], amountsSkippedFilled: [], folderTotals: [], duplicateInvoices: [], looseIncluded: [], reconciled: [],
  };
  // Reuse the folder->app mapping the invoice import already saved, so invoices
  // land in the vendor folder they have always been filed under, alongside the
  // archive's own folder list — so a mapping that names a folder which no longer
  // exists does not resurrect it, and a folder renamed to match the sheet wins.
  // Resolved here rather than at the top of the run because a failed listing is
  // reported through `summary`.
  let archiveFolders = [];
  try {
    archiveFolders = (await folderChildren(token, driveId, base)).folders;
  } catch (e) {
    summary.errors.push(`could not list "${base}" to check which vendor folders exist (${e.message})`);
  }
  const folderForApp = appToSourceFolder(opts.mapping, archiveFolders, appNames);

  const indexEntries = [];
  const marks = [];
  const markKeys = new Set();
  const amounts = {};        // app -> month -> summed USD invoice total
  const touched = new Map(); // app||month -> { app, month, folder } seen this run
  const deadline = opts.deadline || (Date.now() + DEADLINE_MS);

  // Where one vendor-folder/month pair lives, and what is already in it. Each
  // pair costs two Graph listings, and a run now asks for more of them than it
  // used to — an invoice whose billing period moves it lands in a second month
  // folder of the same vendor — so they are resolved once and remembered.
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

    if (newFolder && !summary.newFolders.includes(vendorFolder)) summary.newFolders.push(vendorFolder);

    // The month folder is not resolved here any more: which month an invoice
    // belongs to is settled per PDF, once its billing period has been read.
    let lastFolder = `${base}/${vendorFolder}`;

    for (const att of pdfs) {
      if (Date.now() > deadline) { summary.timedOut = true; break; }
      const fileName = sanitizeFileName(att.name);
      try {
        // Fetch even when the file is already archived: the bytes are what the
        // total is read from, and an invoice filed by hand still has an amount
        // worth picking up.
        const bytes = await mail.getAttachmentBytes(token, mailbox, message.id, att.id);

        // Read the PDF before choosing where it goes. Some vendors bill one
        // company across several rows of the sheet with identical senders,
        // subjects and filenames — Anthropic's API console and Claude seats —
        // and only the invoice body tells them apart.
        const { text: pdfText, error: pdfError } = await readPdfText(bytes);
        const total = pdfError
          ? { amount: null, currency: null, usable: false, note: pdfError }
          : extractInvoiceTotal(pdfText);

        let attApp = app;
        let attVendor = vendorFolder;
        const refined = app && !pdfError ? refineAnthropic(pdfText, appNames) : null;
        if (refined && refined !== app) {
          attApp = refined;
          attVendor = folderForApp[refined] || sanitizeSegment(refined);
          if (!summary.reroutedByContent.some(r => r.file === fileName)) {
            summary.reroutedByContent.push({ file: fileName, from: app, to: refined });
          }
        }

        // An invoice belongs to the month its billing period mostly covers, not
        // the month the mail happened to arrive in. Luzmo's August mail bills
        // "period from 2026-08-26 until 2026-09-26", which is September's
        // charge; only an explicitly stated period moves anything, so an
        // invoice that names none is filed by its mail date exactly as before.
        const placement = invoiceMonth(pdfError ? '' : pdfText, month);
        const attMonth = placement.month || month;
        if (attMonth !== month && !summary.reroutedByPeriod.some(r => r.file === fileName)) {
          // An earlier run — before this rule existed — may have filed the same
          // PDF under the mail's month. Nothing is deleted here, but a leftover
          // copy has to be reported: both folders would be totalled, and the
          // same charge would count in two months.
          const priorPlace = await placeFor(attVendor, month);
          summary.reroutedByPeriod.push({
            file: fileName, app: attApp || null, from: month, to: attMonth,
            periodStart: placement.period ? placement.period.start : null,
            periodEnd: placement.period ? placement.period.end : null,
            alsoStillAt: priorPlace.files.has(fileName) ? priorPlace.folder : null,
          });
        }

        const place = await placeFor(attVendor, attMonth);
        const attFolder = place.folder;
        lastFolder = attFolder;

        if (place.files.has(fileName)) {
          summary.alreadyPresent++;
        } else {
          await uploadFileContent(token, driveId, `${attFolder}/${fileName}`, bytes, 'application/pdf');
          // Remember it, so a second mail carrying the same attachment name in
          // the same run is reported as already present rather than re-uploaded.
          place.files.add(fileName);
          summary.filed++;
          if (attApp) summary.perApp[attApp] = (summary.perApp[attApp] || 0) + 1;
        }

        // Tick the tracker for the app and month this PDF was actually filed
        // under — both can differ from the message's own, once the invoice text
        // has been read. Several invoices for one app-month are normal; they
        // tick one cell, not one per message.
        if (attApp && attMonth) {
          const markKey = `${attApp}||${attMonth}`;
          if (!markKeys.has(markKey)) {
            markKeys.add(markKey);
            marks.push({ app: attApp, month: attMonth });
          }
        }

        // The amount is NOT summed here. Invoices reach the archive by more than
        // one route — hand-filed, mirrored from the source folders, or emailed —
        // and totalling only the emailed ones undercounts badly: Bubble's August
        // mail carried 2 of its 9 charges, so this path alone reported 64.00
        // against an actual 524.27. The folder is totalled after the loop.
        if (attApp && attMonth) touched.set(`${attApp}||${attMonth}`, { app: attApp, month: attMonth, folder: attFolder });

        if (attApp && total.amount !== null && !total.usable) {
          summary.amountsNeedingReview.push({ app: attApp, month: attMonth, file: fileName, amount: total.amount, currency: total.currency, note: total.note });
        } else if (attApp && total.amount === null) {
          summary.amountsUnread.push({ app: attApp, month: attMonth, file: fileName, note: total.note });
        }

        indexEntries.push({
          app: attApp || null, month: attMonth, file: fileName, folder: attFolder,
          amount: total.amount, currency: total.currency, amountUsable: !!total.usable,
          // Kept so a month can be traced back: which rule chose it, the period
          // it was read from, and the month the mail itself arrived in.
          monthVia: placement.via,
          periodStart: placement.period ? placement.period.start : null,
          periodEnd: placement.period ? placement.period.end : null,
          receivedMonth: month,
          subject: message.subject || '', from: identity.originalAddress || identity.address,
          receivedAt: message.receivedDateTime, webLink: message.webLink || null,
        });
      } catch (e) {
        summary.errors.push(`${fileName}: ${e.message}`);
      }
    }

    // Tracker ticks are made per PDF above, where the app and month it was
    // actually filed under are known.
    if (!app) {
      summary.unmatched++;
      summary.unmatchedItems.push({
        subject: (message.subject || '').slice(0, 90),
        from: identity.originalAddress || identity.address,
        month, folder: lastFolder,
      });
    }
    seen.add(message.id);
  }

  // Sweep the archive for recent months too, so an invoice that arrived by any
  // route other than this mailbox is totalled and ticked like the rest. Folded
  // into `touched` and `marks` rather than given a path of its own: everything
  // downstream — the folder total, the top-up rule, the tracker tick — then
  // treats it identically, and there is one set of rules to get right.
  if (!opts.skipReconcile) {
    const byNorm = new Map();
    for (const name of appNames) byNorm.set(normName(name), name);
    const appFor = (folder) => {
      if (folder === UNMATCHED_FOLDER || String(folder).startsWith('_')) return null;
      const direct = byNorm.get(normName(folder));
      if (direct) return direct;
      const mapped = opts.mapping && opts.mapping[folder];
      if (mapped && byNorm.get(normName(mapped))) return byNorm.get(normName(mapped));
      const hit = resolve(folder, '');
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

  // Total each app-month folder that this run touched. Reuse whatever the index
  // already knows so a PDF is parsed once, not on every run.
  if (touched.size) {
    const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
    const cache = new Map();
    for (const e of (Array.isArray(index.amounts) ? index.amounts : [])) {
      if (e && e.path) cache.set(e.path, { amount: e.amount, currency: e.currency, usable: !!e.usable, note: e.note || '' });
    }
    const budget = { parsed: 0, maxParse: 40, deadline, fresh: [], exhausted: false };
    // app||month -> true when some PDF in that folder could not be read.
    const partial = {};

    for (const { app, month, folder, alsoFolders } of touched.values()) {
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
        if (res.counted > 0) {
          const byMonth = amounts[app] || (amounts[app] = {});
          byMonth[month] = res.total;
          // A folder with PDFs nobody could read has not been totalled, only
          // partially added up. Carry that through so the writer can tell a
          // total from a lower bound.
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
    if (budget.fresh.length) {
      const existing = (Array.isArray(index.amounts) ? index.amounts : []).filter(e => !budget.fresh.some(f => f.path === e.path));
      // Spread the index rather than rebuilding it: it also carries `written`
      // (which figures this sync put in the sheet) and `periods` (every billing
      // period read out of a PDF). Writing only the two keys this block knows
      // about used to drop both, so the next run re-parsed the whole archive and
      // lost track of which cells were its own.
      await writeJsonFile(token, driveId, INDEX_PATH, {
        ...index,
        entries: (Array.isArray(index.entries) ? index.entries : []).concat(indexEntries).slice(-500),
        amounts: existing.concat(budget.fresh).slice(-2000),
        updatedAt: new Date().toISOString(),
      });
      indexEntries.length = 0; // already written above
    }
  }

  // Tick the invoice tracker for everything filed against a known app.
  let tracker = { sheet: null, marked: 0 };
  const hasAmounts = Object.keys(amounts).length > 0;
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
          const planned = planAmountCells(amounts, live.grid, { values: live.values, start: live.start }, cellValue, priorWrites, partial);
          summary.amountsSkippedFilled = planned.skippedFilled;
          const toWrite = planned.write.concat(planned.updated);
          if (toWrite.length) {
            const results = await excel.writeCells(token, live.driveId, live.itemId, live.sheetName, toWrite, sessionId, 4);
            const ok = results.filter(r => r.ok);
            summary.amountsWritten = ok.filter(r => !planned.updated.some(u => u.address === r.address))
              .map(r => ({ app: r.app, month: r.month, address: r.address, amount: r.value }));
            summary.amountsUpdated = ok.filter(r => planned.updated.some(u => u.address === r.address))
              .map(r => {
                const u = planned.updated.find(x => x.address === r.address);
                return {
                  app: r.app, month: r.month, address: r.address, amount: r.value,
                  previous: u && u.previous,
                  // false when the figure being replaced was not one this sync
                  // wrote — a hand correction or a statement figure. Surfaced so
                  // replacing someone's number is never silent.
                  wasOurs: !!(u && u.wasOurs),
                };
              });
            // Remember what we wrote, so a later top-up can tell our own figure
            // from one a human has since corrected.
            const written = { ...priorWrites };
            for (const r of ok) written[`${r.app}||${r.month}`] = r.value;
            await writeJsonFile(token, driveId, INDEX_PATH, { ...priorIndex, written, updatedAt: new Date().toISOString() });
            const results2 = results;
            for (const bad of results2.filter(r => !r.ok)) summary.errors.push(`${bad.app} ${bad.month}: ${bad.error}`);
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
  }

  await writeJsonFile(token, driveId, STATE_PATH, {
    mailbox,
    // On a timed-out run, don't advance the watermark past what was processed —
    // the next run picks the remainder up.
    lastRunAt: summary.timedOut ? state.lastRunAt || since : new Date().toISOString(),
    seenMessageIds: [...seen].slice(-500),
    updatedAt: new Date().toISOString(),
  });

  // One PDF can be read twice in a run — once as the mail attachment, once when
  // its folder is totalled — and pdf-parse does not always fail the same way
  // twice, so a single bad file was reported as two problems with two different
  // errors. Report each file once, keeping the first reason given.
  summary.amountsUnread = dedupeByFile(summary.amountsUnread);
  summary.amountsNeedingReview = dedupeByFile(summary.amountsNeedingReview);

  return { ok: true, ranAt: new Date().toISOString(), tracker, ...summary };
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

module.exports = { runMailSync, appToSourceFolder, reconcileArchive, recentMonths, RECONCILE_MONTHS, monthFolderName, planTrackerCells, planAmountCells, folderChildren, sumFolderInvoices };
