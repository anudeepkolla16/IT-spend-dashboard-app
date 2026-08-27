const { encodeGraphPath, sanitizeSegment, readJsonFile, writeJsonFile, uploadFileContent, resolveArchiveRoot, archiveFile, graphFetch } = require('./graph');
const excel = require('./excel');
const { openSpendSheet, readAliasMap, cellValue, appendLog } = require('./spend-sheet');
const { buildResolver, refineAnthropic } = require('./vendor-map');
const mail = require('./mail');
const { extractInvoiceTotal, readPdfText } = require('./invoice-amount');
const { invoiceMonth } = require('./invoice-period');

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
//
// A missing folder is not an error — it simply has no children yet. Any OTHER
// failure is, and used to be swallowed as "empty": a throttled listing then read
// as "this vendor has no month folders", and the run created `Aug-26` beside the
// existing `Aug`, splitting the month's invoices across two folders and totalling
// each of them short. It throws now, and the run reports the file it could not
// place. Paged, for the same reason a truncated listing is worse than no listing.
async function folderChildren(token, driveId, path) {
  const files = new Set();
  const folders = [];
  let url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/children?$select=name,folder,file&$top=200`;
  let pages = 0;

  while (url && pages < 20) {
    const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (res.status === 404) return { files, folders, missing: true };
    if (!res.ok) throw new Error(`could not list "${path}" (${res.status})`);
    const json = await res.json();
    for (const c of json.value || []) {
      if (c.folder) folders.push(c.name);
      else files.add(c.name);
    }
    url = json['@odata.nextLink'] || null;
    pages++;
  }
  return { files, folders };
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
async function sumFolderInvoices(token, driveId, folderPath, cache, budget) {
  const { files } = await folderChildren(token, driveId, folderPath);
  const pdfNames = [...files].filter(n => /\.pdf$/i.test(n));

  let total = 0;
  let counted = 0;
  const unusable = [];
  const unread = [];

  for (const name of pdfNames) {
    const key = `${folderPath}/${name}`;
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
      } catch (e) {
        entry = { amount: null, currency: null, usable: false, note: e.message };
      }
      cache.set(key, entry);
      budget.fresh.push({ path: key, amount: entry.amount, currency: entry.currency, usable: !!entry.usable });
    }

    if (entry.usable) { total += entry.amount; counted++; }
    else if (entry.amount !== null) unusable.push({ file: name, amount: entry.amount, currency: entry.currency, note: entry.note });
    else unread.push({ file: name, note: entry.note });
  }

  return { total: Math.round(total * 100) / 100, counted, pdfCount: pdfNames.length, unusable, unread };
}

// Decides which Spendings cells an invoice total may fill. The rule is narrow on
// purpose: only an empty cell is written. A cell that already holds a figure is
// left exactly as it is and reported, because the existing value may be the
// bank-statement figure or a hand correction, and an invoice total is not
// authoritative over either.
const CELL_EPSILON = 0.005;

function planAmountCells(amounts, grid, used, cellValueOf, priorWrites) {
  const rowByApp = new Map(grid.apps.map(a => [a.name, a.rowIdx]));
  const prior = priorWrites || {};
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

      if (current === null || current === 0) {
        write.push({ app, month, address, value: amount });
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

      // A total LOWER than the cell is the one case left alone. A new invoice can
      // only add, so a shortfall means invoices are missing from the folder, not
      // that less was spent — Bubble sat at 492.27 against a correct 524.27 with
      // its ninth invoice still pending. Overwriting there would replace a right
      // figure with a wrong one, so it is reported instead.
      if (amount < current - CELL_EPSILON) {
        skippedFilled.push({ app, month, address, current, invoiceTotal: amount, reason: 'invoice-total-lower' });
      }
    }
  }
  return { write, updated, skippedFilled };
}

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
  const root = await resolveArchiveRoot(token, driveId);
  const base = root.path;
  const STATE_PATH = archiveFile(root, '_mail-sync.json');
  const INDEX_PATH = archiveFile(root, '_invoice-index.json');
  // Reuse the folder->app mapping the invoice import already saved, so invoices
  // land in the vendor folder they have always been filed under.
  const folderForApp = appToSourceFolder(opts.mapping);

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
    amountsWritten: [], amountsUpdated: [], amountsNeedingReview: [], amountsUnread: [], amountsSkippedFilled: [], folderTotals: [],
  };
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

  // Total each app-month folder that this run touched. Reuse whatever the index
  // already knows so a PDF is parsed once, not on every run.
  if (touched.size) {
    const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
    const cache = new Map();
    for (const e of (Array.isArray(index.amounts) ? index.amounts : [])) {
      if (e && e.path) cache.set(e.path, { amount: e.amount, currency: e.currency, usable: !!e.usable, note: e.note || '' });
    }
    const budget = { parsed: 0, maxParse: 40, deadline, fresh: [], exhausted: false };

    for (const { app, month, folder } of touched.values()) {
      try {
        const res = await sumFolderInvoices(token, driveId, folder, cache, budget);
        if (res.counted > 0) {
          const byMonth = amounts[app] || (amounts[app] = {});
          byMonth[month] = res.total;
        }
        summary.folderTotals.push({
          app, month, folder, total: res.total,
          invoicesCounted: res.counted, pdfsInFolder: res.pdfCount,
        });
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
          const planned = planAmountCells(amounts, live.grid, { values: live.values, start: live.start }, cellValue, priorWrites);
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

  return { ok: true, ranAt: new Date().toISOString(), tracker, ...summary };
}

module.exports = { runMailSync, appToSourceFolder, monthFolderName, planTrackerCells, planAmountCells, folderChildren, sumFolderInvoices };
