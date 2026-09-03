// Re-checks the billing period of every invoice already in the archive, and
// puts the ones that were filed by their arrival date into the month they
// actually pay for — in the folders and in the spend sheet.
//
// The mail sync now files by the invoice's billing period (lib/invoice-period.js),
// but everything filed before that rule existed sits where its email happened to
// land: Luzmo's Aug-26 invoice, which bills 2026-08-26 → 2026-09-26, is in the
// August folder and its 557.28 is in the August cell.
//
// Two halves, deliberately separated:
//   scanPeriods()  reads every PDF's period and reports what is misfiled.
//                  It writes nothing at all.
//   applyBackfill() moves the files the caller approved and writes the cells the
//                  caller approved, recomputing each affected month from the
//                  folder as it stands after the moves.
//
// Nothing is deleted, ever. A move is a Graph move, so the file keeps its
// identity, its history and anyone's links to it, and moving it back undoes it.
//
// The archive is one folder, found by resolveArchiveRoot rather than assumed —
// see the note in lib/graph.js for what assuming it cost last time. Every path
// here is built from the root the caller resolved, so a rename moves the whole
// backfill with it instead of pointing it at a folder that is no longer there.

const {
  encodeGraphPath, graphFetch, graphListAll, listFilesRecursive, readJsonFile, writeJsonFile,
  itemIdByPath, ensureFolder, moveItem, archiveFile,
} = require('../graph');
const { readPdfText, extractInvoiceTotal } = require('../invoice-amount');
const { extractBillingPeriod, extractInvoiceDate, monthForPeriod } = require('../invoice-period');
const { monthFolderName, appToSourceFolder, sumFolderInvoices } = require('../mail-sync');
const { buildResolver, isIgnored } = require('../vendor-map');
const { parseMonthFolder, monthFromFileName, detectDayFirst } = require('./inventory');
const excel = require('../excel');
const { openSpendSheet, cellValue, appendLog, readAliasMap } = require('../spend-sheet');

const CELL_EPSILON = 0.005;

// How many PDFs one run may download and read. Each is a download plus a parse,
// and the function has 45 seconds; the rest are picked up by the next run, which
// costs nothing extra because every read is cached in the invoice index.
// Bounded by the run's deadline rather than a count: reads happen in parallel,
// so a 45-second run gets through a few hundred rather than the twenty-five a
// serial loop managed. The cap is only a backstop against a pathological run.
const DEFAULT_MAX_PARSE = 600;
const LIST_POOL = 6;
const READ_POOL = 6;
// Stop reading with time in hand: the run still has to total the folders it
// touched and write the cache of what it read.
const READ_RESERVE_MS = 6000;

const childrenOf = (token, driveId, path) => graphListAll(
  token,
  `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/children?$select=id,name,folder&$top=200`,
  `listing "${path}"`
);

// What the index remembers about each PDF it has already read, by path.
function cachedReads(index) {
  const periods = new Map();
  for (const e of (Array.isArray(index && index.periods) ? index.periods : [])) {
    if (e && e.path) periods.set(e.path, e);
  }
  return periods;
}

// Read one invoice: its billing period and its total, in a single parse.
async function readInvoice(token, driveId, path) {
  const res = await graphFetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`download ${res.status}`);
  const bytes = Buffer.from(await res.arrayBuffer());
  const { text, pages, error } = await readPdfText(bytes);
  if (error) return { path, read: false, note: error };

  const period = extractBillingPeriod(text);
  const total = extractInvoiceTotal(text, { pages });
  // Only wanted for the invoices with no month anywhere else. Read here anyway
  // so it lands in the cache with everything else — going back for it later
  // would mean downloading and parsing the same PDF a second time.
  const issued = extractInvoiceDate(text);
  return {
    path,
    read: true,
    periodStart: period ? period.start : null,
    periodEnd: period ? period.end : null,
    invoiceDate: issued ? issued.date : null,
    invoiceMonth: issued ? issued.month : null,
    amount: total.amount,
    currency: total.currency,
    usable: !!total.usable,
    note: total.note || '',
  };
}

// One skipped invoice, as the plan reports it. The dashboard renders these by
// name so the few that need filing by hand can be, which only works while both
// sides agree on the field names — see test/period-backfill.test.js, which pins
// them against index.html after a rename here left the page printing a list of
// bare dashes.
// `kind` groups the list: 84 individually-worded lines say nothing about the
// shape of the problem, while "62 name disagrees, 12 unreadable, 10 nested"
// says which one is worth an afternoon.
const SKIP_KINDS = ['unreadable', 'undated', 'name-disagrees', 'nested', 'other'];

function skipEntry(file, folder, why, kind) {
  return {
    file: file.name,
    folder: folder.vendorFolder,
    why,
    kind: SKIP_KINDS.includes(kind) ? kind : 'other',
  };
}

// Does this invoice belong in another month, and where would it go?
//
// Returns null to leave it exactly where it is, { skip } for one that cannot be
// judged, or { move }. The month folder it moves into reuses whatever this
// vendor already calls that month ("Aug", "Aug-26", "August 2026") rather than
// adding a second folder for the same month beside it.
function planMove(file, folder, base) {
  // Not read yet — the run ran out of time before reaching it. The next one
  // picks it up; saying anything about it now would be a guess.
  if (file.pending) return null;
  if (!file.read) return { skip: file.note || 'could not be read', kind: 'unreadable' };

  // An invoice with no month anywhere — not in a folder, not in its name — is
  // the one case where moving is unambiguously an improvement. It counts
  // towards no month at all today, so ClickUp reads as a gap for a month whose
  // invoice is sitting right there in the folder. There is no name or folder to
  // contradict, so file it by what the PDF itself says: its billing period if
  // it states one, otherwise the date it was issued.
  if (!file.folderMonth && !file.nameMonth) {
    const byPeriod = file.periodStart
      ? monthForPeriod({ start: file.periodStart, end: file.periodEnd }, null).month
      : null;
    const month = byPeriod || file.invoiceMonth || null;
    if (!month) {
      return { skip: `is filed loose in "${folder.vendorFolder}" with no month in its name and none stated inside — file it under a month folder by hand`, kind: 'undated' };
    }
    const into = monthFolderName(folder.monthFolderNames, month) || month;
    return {
      move: {
        app: folder.app,
        vendorFolder: folder.vendorFolder,
        file: file.name,
        fromMonth: null, // it belonged to no month; that is the point
        toMonth: month,
        fromPath: file.path,
        toFolderPath: `${base}/${folder.vendorFolder}/${into}`,
        periodStart: file.periodStart,
        periodEnd: file.periodEnd,
        invoiceDate: file.invoiceDate,
        amount: file.usable ? file.amount : null,
        via: byPeriod ? 'period' : 'invoice-date',
        undated: true,
      },
    };
  }

  // No stated period is not a misfiling: most invoices state none, and their
  // arrival month is the best thing anyone knows about them.
  if (!file.periodStart) return null;

  // Much of the archive keeps invoices flat in the app folder with the month in
  // the name ("jan 26.pdf") — there is no month folder to move them out of, and
  // renaming somebody's files to impose one is not this job. Report the ones
  // whose period disagrees with the name the checklist dates them by, and leave
  // them where they are.
  if (!file.folderMonth) {
    const byPeriod = monthForPeriod({ start: file.periodStart, end: file.periodEnd }, file.nameMonth);
    if (!byPeriod.month || byPeriod.month === file.nameMonth) return null;
    return { skip: `is filed loose in "${folder.vendorFolder}" and its name reads as ${file.nameMonth}, but it bills ${byPeriod.month} — rename it or file it under a month folder`, kind: 'name-disagrees' };
  }

  if (String(file.relPath || '').split('/').length > 1) {
    return { skip: `sits in "${file.relPath}", not a plain month folder`, kind: 'nested' };
  }

  const resolved = monthForPeriod({ start: file.periodStart, end: file.periodEnd }, file.currentMonth);
  if (!resolved.month || resolved.month === file.currentMonth) return null;

  const toFolder = monthFolderName(folder.monthFolderNames, resolved.month) || resolved.month;
  return {
    move: {
      app: folder.app,
      vendorFolder: folder.vendorFolder,
      file: file.name,
      fromMonth: file.currentMonth,
      toMonth: resolved.month,
      fromPath: file.path,
      toFolderPath: `${base}/${folder.vendorFolder}/${toFolder}`,
      periodStart: file.periodStart,
      periodEnd: file.periodEnd,
      amount: file.usable ? file.amount : null,
      via: resolved.via,
    },
  };
}

// Every PDF in the archive, and whether its period puts it somewhere else.
//
// Returns proposals (a file to move), the sheet cells those moves would change,
// and how much of the archive is still unread — the caller runs it again until
// `unread` is 0.
async function scanPeriods(token, driveId, opts) {
  const options = opts || {};
  const deadline = options.deadline || (Date.now() + 40 * 1000);
  const maxParse = options.maxParse || DEFAULT_MAX_PARSE;
  const root = options.root;
  if (!root || !root.resolved) {
    return {
      scannedAt: new Date().toISOString(), vendors: 0, invoices: 0, parsed: 0, unread: 0,
      moves: [], cells: [], skipped: [], truncated: false,
      errors: [`The invoice archive could not be found. Tried: ${(root && root.candidates || []).join(', ') || 'nothing'}`],
    };
  }
  const base = root.path;
  const INDEX_PATH = archiveFile(root, '_invoice-index.json');
  const folderForApp = appToSourceFolder(options.mapping);
  const mapped = new Map(Object.entries(folderForApp).map(([app, folder]) => [folder, app]));

  const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
  const periods = cachedReads(index);

  const summary = {
    scannedAt: new Date().toISOString(),
    base, vendors: 0, invoices: 0, parsed: 0, unread: 0,
    moves: [], cells: [], skipped: [], errors: [],
    truncated: false,
  };

  // The archive's folders carry vendor names and the sheet's rows carry app
  // names, and the two often disagree — "Cumul" is "Cumul(Luzmo)", "Bubble" is
  // "Bubble Starter". The saved folder->app mapping settles most of it; the rest
  // goes through the same resolver the mail sync files by, so a vendor folder is
  // not left without a row (and its month without a total) over a spelling.
  let sheet = null;
  let appOf = folder => mapped.get(folder) || null;
  try {
    sheet = await openSpendSheet(token);
    const appNames = sheet.grid.apps.map(a => a.name);
    const resolve = buildResolver(await readAliasMap(token, driveId), appNames);
    appOf = (folder) => {
      const viaMapping = mapped.get(folder);
      if (viaMapping && appNames.includes(viaMapping)) return viaMapping;
      const hit = resolve(folder, '');
      return hit.app || viaMapping || null;
    };
  } catch (e) {
    summary.errors.push(`could not read the spend sheet, so no cell changes are proposed: ${e.message}`);
  }

  const fresh = [];
  // vendor||month -> { app, vendorFolder, monthFolder, files: [] }
  const folders = new Map();

  let vendorFolders;
  try {
    // The folders deliberately left out of the checklist are left out here too.
    // "Courier bills" and "Laptops sold" have no row in the sheet and never
    // will, so every file in them came back as an invoice that could not be
    // dated — a cancelled cheque from 2022 among them. Reporting those buries
    // the ones that are real, and reading them spends the parse budget of a run
    // on files no month will ever hold.
    vendorFolders = (await childrenOf(token, driveId, base))
      .filter(c => c.folder && !String(c.name).startsWith('_') && !isIgnored(c.name));
  } catch (e) {
    return { ...summary, errors: [`could not list the archive: ${e.message}`] };
  }
  summary.vendors = vendorFolders.length;

  // Pass 1a: list what is there. One listing per vendor folder, run through a
  // small pool — forty of them in series spent a fifth of the run just waiting.
  const listed = [];
  let vendorCursor = 0;
  await Promise.all(Array.from({ length: Math.min(LIST_POOL, vendorFolders.length) }, async () => {
    while (vendorCursor < vendorFolders.length) {
      if (Date.now() > deadline) { summary.truncated = true; return; }
      const vendor = vendorFolders[vendorCursor++];
      try {
        listed.push({ vendor, files: await listFilesRecursive(token, driveId, vendor.id) });
      } catch (e) {
        summary.errors.push(`${vendor.name}: ${e.message}`);
      }
    }
  }));

  // Pass 1b: work out where each PDF sits, and which ones still need reading.
  const pending = [];
  for (const { vendor, files } of listed) {
    const monthFolderNames = [...new Set(files.map(f => String(f.relPath || '').split('/')[0]).filter(Boolean))];
    // "03-12-2025" is only readable once some other file in the folder settles
    // the order, so the whole vendor's names are weighed together (#20).
    const dayFirst = detectDayFirst(files.map(f => f.name));

    for (const file of files) {
      if (!/\.pdf$/i.test(file.name)) continue;
      summary.invoices++;

      const relPath = String(file.relPath || '');
      const path = `${base}/${vendor.name}${relPath ? '/' + relPath : ''}/${file.name}`;
      // The folder it sits in comes first — somebody put it there deliberately —
      // and the file name second, which is how the checklist dates the rest.
      const folderMonth = relPath ? parseMonthFolder(relPath.split('/')[0], file.createdDateTime) : null;
      const nameMonth = monthFromFileName(file.name, file.createdDateTime, dayFirst);

      const bucket = `${vendor.name}||${relPath.split('/')[0] || ''}`;
      if (!folders.has(bucket)) {
        folders.set(bucket, {
          app: appOf(vendor.name),
          vendorFolder: vendor.name,
          // The bucket is one folder on disk. `month` is the month that folder
          // IS; a bucket of loose files is not a month folder at all, whatever
          // the names inside it happen to read as.
          monthFolder: relPath.split('/')[0] || '',
          month: relPath ? folderMonth : null,
          monthFolderNames,
          files: [],
        });
      }
      const placed = {
        name: file.name, path, relPath, id: file.id,
        currentMonth: folderMonth || nameMonth, folderMonth, nameMonth,
        invoiceDate: null, invoiceMonth: null,
      };
      folders.get(bucket).files.push(placed);
      if (!periods.has(path)) pending.push(placed);
    }
  }

  // Pass 1c: read the ones the index has never seen. This is the whole cost of
  // a run — a download and a parse each — so they go through a pool, and the
  // deadline rather than a fixed count decides how far one run gets. Whatever is
  // left is reported and picked up by the next run, which re-reads nothing.
  let readCursor = 0;
  await Promise.all(Array.from({ length: Math.min(READ_POOL, pending.length) }, async () => {
    while (readCursor < pending.length) {
      if (summary.parsed >= maxParse || Date.now() > deadline - READ_RESERVE_MS) return;
      const target = pending[readCursor++];
      summary.parsed++;
      let entry;
      try {
        entry = await readInvoice(token, driveId, target.path);
      } catch (e) {
        entry = { path: target.path, read: false, note: e.message };
      }
      periods.set(target.path, entry);
      fresh.push(entry);
    }
  }));

  // Everything read, from this run or a previous one, joins its file.
  for (const folder of folders.values()) {
    for (const file of folder.files) {
      const entry = periods.get(file.path);
      if (entry) Object.assign(file, entry, { path: file.path });
      else { file.pending = true; summary.unread++; summary.truncated = true; }
    }
  }

  // Pass 2: which of them belong somewhere else.
  const movedOut = new Map();  // "vendor||month" -> the files leaving it
  const movedIn = new Map();   // "vendor||month" -> the files arriving in it

  for (const folder of folders.values()) {
    for (const file of folder.files) {
      const verdict = planMove(file, folder, base);
      if (!verdict) continue;
      if (verdict.skip) { summary.skipped.push(skipEntry(file, folder, verdict.skip, verdict.kind)); continue; }

      summary.moves.push(verdict.move);
      const outKey = `${folder.vendorFolder}||${file.currentMonth}`;
      const inKey = `${folder.vendorFolder}||${verdict.move.toMonth}`;
      if (!movedOut.has(outKey)) movedOut.set(outKey, []);
      if (!movedIn.has(inKey)) movedIn.set(inKey, []);
      movedOut.get(outKey).push(file);
      movedIn.get(inKey).push(file);
    }
  }

  // Pass 3: what those moves would do to the sheet. Predicted here, and
  // recomputed from the folders for real before anything is written.
  if (summary.moves.length) {
    try {
      summary.cells = sheet ? predictCells(sheet, folders, movedOut, movedIn, appOf, index.written, opts.locks) : [];
    } catch (e) {
      summary.errors.push(`sheet preview: ${e.message}`);
    }
  }

  // Remember every read, so the next run costs a listing rather than a download.
  if (fresh.length) {
    const keep = (Array.isArray(index.periods) ? index.periods : []).filter(e => !fresh.some(f => f.path === e.path));
    await writeJsonFile(token, driveId, INDEX_PATH, {
      ...index,
      periods: keep.concat(fresh).slice(-3000),
      updatedAt: new Date().toISOString(),
    });
  }

  return summary;
}

const sumUsable = files => Math.round(
  (files || []).filter(f => f.usable).reduce((sum, f) => sum + f.amount, 0) * 100
) / 100;

const near = (a, b) => a !== null && a !== undefined && b !== null && b !== undefined && Math.abs(a - b) < CELL_EPSILON;

// Whether a month's cell may be rewritten from its folder, and to what.
//
// This is the one path in the app that writes a SMALLER number than it found,
// on the reasoning that an invoice which moved out of a month was never that
// month's. That reasoning holds for the invoices — it says nothing about a
// figure that did not come from them.
//
// A live run proved the difference. Cumul(Luzmo)'s August cell held 14,081.00
// and its August folder held one 557.28 invoice; moving that invoice out left
// the folder empty, and the rule as written offered to replace 14,081.00 with
// 0.00. The 14,081.00 is a bank-statement or hand-entered figure, and the
// backfill has no idea what it is made of — only that it is not the invoices.
//
// The cell is the month's invoice total, so once the moves are done it is set
// to what the month's folder comes to — the same rule the mail sync applies.
// A figure the invoices do not account for (a statement total, a hand-typed
// number) is still replaced, because the owner asked for the cell to BE the
// invoice total; but it is flagged, with both figures, so the confirmation
// shows whose number is going. The user ticks the cell before anything is
// written, so nothing here is a guess made on their behalf.
function decideCell({ current, before, after, ourLastWrite }) {
  if (near(current, after)) return { write: false, why: 'already right' };
  if (current === null || current === 0) return { write: true, value: after };
  if (near(current, before)) return { write: true, value: after };
  if (near(current, ourLastWrite)) return { write: true, value: after };
  return {
    write: true,
    value: after,
    replaces: `the sheet holds ${current.toLocaleString('en-US', { minimumFractionDigits: 2 })}, which is not what the invoices in that month come to`
      + `${before === 0 ? '' : ` (${before.toLocaleString('en-US', { minimumFractionDigits: 2 })} before this change)`}`,
  };
}

// The month totals the moves imply, against what the sheet holds now. A month
// with any unread invoice in it is left out: a total built on a PDF nobody could
// read would be short, and writing it would look authoritative.
function predictCells(sheet, folders, movedOut, movedIn, appOf, priorWrites, locks) {
  const rowByApp = new Map(sheet.grid.apps.map(a => [a.name, a.rowIdx]));
  const written = priorWrites || {};
  const normKey = (x) => String(x == null ? '' : x).toLowerCase().replace(/[^a-z0-9]/g, '');
  const isLocked = (app, month) => (Array.isArray(locks) ? locks : []).some(l => normKey(l.app) === normKey(app) && l.month === month);

  // vendor||month -> the files that folder will hold once the moves are done.
  //
  // Only real month folders count. A vendor's loose files are one bucket with no
  // month of its own, and its key would otherwise collide with a month a file is
  // moving INTO — putting every loose invoice the vendor has into that month's
  // predicted total. They are excluded from the totals for the same reason the
  // sync's own totals exclude them: a month's figure is the sum of its folder.
  // One vendor can have two folders for the same month — Luzmo has "Jul-26" and
  // a "July" that Tidy lifts out of a nested "Luzmo/". They share a key, so
  // setting it twice dropped whichever came second and the month was totalled
  // from half its invoices. They are merged instead: a month is the sum of every
  // folder that names it.
  const after = new Map();
  for (const folder of folders.values()) {
    if (!folder.monthFolder || !folder.month) continue;
    const key = `${folder.vendorFolder}||${folder.month}`;
    const leaving = new Set((movedOut.get(key) || []).map(f => f.path));
    const staying = folder.files.filter(f => !leaving.has(f.path));
    const at = after.get(key);
    if (at) {
      at.before += sumUsable(folder.files);
      at.files.push(...staying);
      continue;
    }
    after.set(key, {
      app: folder.app,
      month: folder.month,
      // What the folder totals now, before anything moves — the evidence that
      // the cell's figure is the invoices' and not somebody else's.
      before: sumUsable(folder.files),
      files: staying,
    });
  }
  for (const [key, files] of movedIn) {
    const [vendorFolder, month] = key.split('||');
    if (!after.has(key)) after.set(key, { app: appOf(vendorFolder), month, before: 0, files: [] });
    after.get(key).files.push(...files);
  }

  const cells = [];
  for (const [key, state] of after) {
    if (!state.app || !state.month) continue;
    const touched = movedOut.has(key) || movedIn.has(key);
    if (!touched) continue;

    const rowIdx = rowByApp.get(state.app);
    const colIdx = sheet.grid.monthCols[state.month];
    if (rowIdx === undefined || colIdx === undefined) {
      cells.push({ app: state.app, month: state.month, blocked: 'the sheet has no cell for that app and month' });
      continue;
    }
    if (isLocked(state.app, state.month)) {
      cells.push({ app: state.app, month: state.month, blocked: 'you locked this cell; the sync leaves it as you set it' });
      continue;
    }
    if (state.files.some(f => !f.read || (!f.usable && f.amount === null))) {
      cells.push({ app: state.app, month: state.month, blocked: 'an invoice in that month could not be read' });
      continue;
    }

    const total = sumUsable(state.files);
    const current = cellValue(sheet.values, rowIdx, colIdx);
    const verdict = decideCell({
      current, before: state.before, after: total,
      ourLastWrite: written[`${state.app}||${state.month}`],
    });
    if (!verdict.write) continue;

    cells.push({
      app: state.app,
      month: state.month,
      address: excel.cellAddress(sheet.start, rowIdx, colIdx),
      current,
      // Carried to the apply, which re-derives the figure and flags the same
      // replacement the scan did.
      before: state.before,
      value: total,
      invoices: state.files.filter(f => f.usable).length,
      direction: current === null || total > current ? 'up' : 'down',
      // "the invoices come to 585.15" does not say whether the archive is short
      // or the sheet is stale: Luzmo's July was three invoices totalling
      // 1,727.58 against a 13,439.00 charge (the archive is missing some),
      // Shopify's February one invoice for 19.00 against 20.00 (the sheet is a
      // little off). The files are listed so the confirmation makes that a
      // glance rather than a judgement call.
      ...(verdict.replaces ? {
        replaces: verdict.replaces,
        files: state.files.map(f => ({ file: f.name, amount: f.usable ? f.amount : null })),
      } : {}),
    });
  }
  return cells.sort((a, b) => a.app.localeCompare(b.app) || a.month.localeCompare(b.month));
}

// Move the approved files, then write the approved cells.
//
// The values written are NOT the ones previewed: each affected month is totalled
// again from the folder as it stands once the moves are done, so a file that
// failed to move cannot leave a figure behind that assumes it did.
async function applyBackfill(token, driveId, payload, opts) {
  const options = opts || {};
  const deadline = options.deadline || (Date.now() + 40 * 1000);
  const root = options.root;
  if (!root || !root.resolved) {
    return { moved: [], failed: [], cellsWritten: [], cellsSkipped: [], timedOut: false,
      errors: ['The invoice archive could not be found, so nothing was moved.'] };
  }
  const INDEX_PATH = archiveFile(root, '_invoice-index.json');
  const moves = Array.isArray(payload && payload.moves) ? payload.moves : [];
  // Keyed by app||month, and carrying what the scan saw: the figure in the cell
  // and what the folder totalled before the moves. The apply re-derives the new
  // figure from the folder, but it re-applies the same test before writing —
  // approval says which months to consider, never that a cell may be replaced.
  const approved = new Map((Array.isArray(payload && payload.cells) ? payload.cells : [])
    .map(c => [`${c.app}||${c.month}`, c]));

  const summary = { moved: [], failed: [], cellsWritten: [], cellsSkipped: [], errors: [], timedOut: false };
  const touched = new Map(); // app||month -> { app, month, folder }
  const renames = [];        // [fromPath, toPath] for the index caches

  for (const move of moves) {
    if (Date.now() > deadline) { summary.timedOut = true; break; }
    try {
      const itemId = await itemIdByPath(token, driveId, move.fromPath);
      if (!itemId) throw new Error('the file is no longer where the scan found it');
      const targetId = await ensureFolder(token, driveId, move.toFolderPath);
      await moveItem(token, driveId, itemId, targetId);
      renames.push([move.fromPath, `${move.toFolderPath}/${move.file}`]);
      summary.moved.push(move);
      if (move.app) {
        touched.set(`${move.app}||${move.fromMonth}`, { app: move.app, month: move.fromMonth, folder: dirOf(move.fromPath) });
        touched.set(`${move.app}||${move.toMonth}`, { app: move.app, month: move.toMonth, folder: move.toFolderPath });
      }
    } catch (e) {
      summary.failed.push({ file: move.file, from: move.fromMonth, to: move.toMonth, error: e.message });
    }
  }

  if (renames.length) await renameIndexPaths(token, driveId, INDEX_PATH, renames);
  if (!summary.moved.length) return summary;

  // Re-total every month a move touched, from the folders as they now stand.
  const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
  const cache = new Map();
  for (const e of (Array.isArray(index.amounts) ? index.amounts : [])) {
    if (e && e.path) cache.set(e.path, { amount: e.amount, currency: e.currency, usable: !!e.usable, note: e.note || '' });
  }
  const budget = { parsed: 0, maxParse: 40, deadline, fresh: [], exhausted: false };
  const totals = [];
  for (const { app, month, folder } of touched.values()) {
    const ticked = approved.get(`${app}||${month}`);
    if (!ticked) { summary.cellsSkipped.push({ app, month, why: 'not approved' }); continue; }
    try {
      const res = await sumFolderInvoices(token, driveId, folder, cache, budget);
      if (res.unread.length || res.unusable.length) {
        summary.cellsSkipped.push({ app, month, why: 'an invoice in that month could not be read as a USD total' });
        continue;
      }
      totals.push({ app, month, value: res.total, before: ticked.before, ourLastWrite: (index.written || {})[`${app}||${month}`] });
    } catch (e) {
      summary.errors.push(`${app} ${month}: could not total the folder (${e.message})`);
    }
  }

  if (totals.length) await writeCells(token, driveId, totals, summary);
  return summary;
}

const dirOf = p => String(p || '').split('/').slice(0, -1).join('/');

// Point the index's cached reads at the paths the files now live under, so the
// next run does not re-download every invoice it has already read.
async function renameIndexPaths(token, driveId, INDEX_PATH, renames) {
  const index = (await readJsonFile(token, driveId, INDEX_PATH)) || {};
  const map = new Map(renames);
  const move = list => (Array.isArray(list) ? list : []).map(e => (e && e.path && map.has(e.path) ? { ...e, path: map.get(e.path) } : e));
  await writeJsonFile(token, driveId, INDEX_PATH, {
    ...index,
    periods: move(index.periods),
    amounts: move(index.amounts),
    updatedAt: new Date().toISOString(),
  });
}

// Write the recomputed totals, in one workbook session, and log what changed.
// Unlike the invoice sync, this DOES lower a cell: an invoice that has moved out
// of a month is not a missing invoice, it is one that was never that month's.
async function writeCells(token, driveId, totals, summary) {
  let sessionId = null;
  const sheet = await openSpendSheet(token);
  try {
    sessionId = await excel.createSession(token, sheet.driveId, sheet.itemId);
    const live = await openSpendSheet(token, sessionId);
    const rowByApp = new Map(live.grid.apps.map(a => [a.name, a.rowIdx]));

    const cells = [];
    for (const t of totals) {
      const rowIdx = rowByApp.get(t.app);
      const colIdx = live.grid.monthCols[t.month];
      if (rowIdx === undefined || colIdx === undefined) {
        summary.cellsSkipped.push({ app: t.app, month: t.month, why: 'the sheet has no cell for that app and month' });
        continue;
      }
      if ((Array.isArray(opts.locks) ? opts.locks : []).some(l => String(l.app).toLowerCase().replace(/[^a-z0-9]/g, '') === String(t.app).toLowerCase().replace(/[^a-z0-9]/g, '') && l.month === t.month)) {
        summary.cellsSkipped.push({ app: t.app, month: t.month, why: 'you locked this cell; the sync leaves it as you set it' });
        continue;
      }
      // The same test the scan applied, against the sheet as it is right now.
      // The scan's answer is not taken on trust: the sheet may have been edited
      // in between, and this is the one path that writes a lower figure.
      const current = cellValue(live.values, rowIdx, colIdx);
      const verdict = decideCell({ current, before: t.before, after: t.value, ourLastWrite: t.ourLastWrite });
      if (!verdict.write) continue;
      cells.push({ app: t.app, month: t.month, address: excel.cellAddress(live.start, rowIdx, colIdx), value: t.value, before: current, replaces: verdict.replaces || null });
    }
    if (!cells.length) return;

    // writeCells carries each cell through to its result, so `before` — what the
    // sheet held — comes back with it and lands in the audit log unchanged.
    const results = await excel.writeCells(token, live.driveId, live.itemId, live.sheetName, cells, sessionId, 4);
    const ok = results.filter(r => r.ok);
    for (const r of ok) {
      summary.cellsWritten.push({ app: r.app, month: r.month, address: r.address, value: r.value, previous: r.before, replaces: r.replaces || null });
    }
    for (const bad of results.filter(r => !r.ok)) summary.errors.push(`${bad.app} ${bad.month}: ${bad.error}`);

    if (ok.length) {
      await appendLog(token, driveId, {
        at: new Date().toISOString(), by: 'period-backfill', source: 'invoice',
        statement: null, attribution: 'billing-period', sheet: live.sheetName,
        cells: ok.map(c => ({ app: c.app, month: c.month, address: c.address, before: c.before === undefined ? null : c.before, after: c.value })),
        failed: [],
      });
    }
  } catch (e) {
    summary.errors.push(`sheet update: ${e.message}`);
  } finally {
    if (sessionId) await excel.closeSession(token, sheet.driveId, sheet.itemId, sessionId);
  }
}

module.exports = { skipEntry, SKIP_KINDS, scanPeriods, applyBackfill, planMove, predictCells, decideCell, readInvoice };
