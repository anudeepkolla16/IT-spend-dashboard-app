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
const { extractBillingPeriod, monthForPeriod } = require('../invoice-period');
const { monthFolderName, appToSourceFolder, sumFolderInvoices } = require('../mail-sync');
const { buildResolver } = require('../vendor-map');
const { parseMonthFolder, monthFromFileName, detectDayFirst } = require('./inventory');
const excel = require('../excel');
const { openSpendSheet, cellValue, appendLog, readAliasMap } = require('../spend-sheet');

const CELL_EPSILON = 0.005;

// How many PDFs one run may download and read. Each is a download plus a parse,
// and the function has 45 seconds; the rest are picked up by the next run, which
// costs nothing extra because every read is cached in the invoice index.
const DEFAULT_MAX_PARSE = 25;

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
  const { text, error } = await readPdfText(bytes);
  if (error) return { path, read: false, note: error };

  const period = extractBillingPeriod(text);
  const total = extractInvoiceTotal(text);
  return {
    path,
    read: true,
    periodStart: period ? period.start : null,
    periodEnd: period ? period.end : null,
    amount: total.amount,
    currency: total.currency,
    usable: !!total.usable,
    note: total.note || '',
  };
}

// Does this invoice belong in another month, and where would it go?
//
// Returns null to leave it exactly where it is, { skip } for one that cannot be
// judged, or { move }. The month folder it moves into reuses whatever this
// vendor already calls that month ("Aug", "Aug-26", "August 2026") rather than
// adding a second folder for the same month beside it.
function planMove(file, folder, base) {
  if (!file.read) return { skip: file.note || 'could not be read' };
  // No stated period is not a misfiling: most invoices state none, and their
  // arrival month is the best thing anyone knows about them.
  if (!file.periodStart) return null;

  // Much of the archive keeps invoices flat in the app folder with the month in
  // the name ("jan 26.pdf") — there is no month folder to move them out of, and
  // renaming somebody's files to impose one is not this job. Report the ones
  // whose period disagrees with the name the checklist dates them by, and leave
  // them where they are.
  if (!file.folderMonth) {
    if (!file.nameMonth) return null;
    const byPeriod = monthForPeriod({ start: file.periodStart, end: file.periodEnd }, file.nameMonth);
    if (!byPeriod.month || byPeriod.month === file.nameMonth) return null;
    return { skip: `is filed loose in "${folder.vendorFolder}" and its name reads as ${file.nameMonth}, but it bills ${byPeriod.month} — rename it or file it under a month folder` };
  }

  if (String(file.relPath || '').split('/').length > 1) {
    return { skip: `sits in "${file.relPath}", not a plain month folder` };
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
    vendorFolders = (await childrenOf(token, driveId, base))
      .filter(c => c.folder && !String(c.name).startsWith('_'));
  } catch (e) {
    return { ...summary, errors: [`could not list the archive: ${e.message}`] };
  }
  summary.vendors = vendorFolders.length;

  // Pass 1: collect every PDF, reading the ones the index has not seen yet.
  for (const vendor of vendorFolders) {
    if (Date.now() > deadline) { summary.truncated = true; break; }
    let files;
    try {
      files = await listFilesRecursive(token, driveId, vendor.id);
    } catch (e) {
      summary.errors.push(`${vendor.name}: ${e.message}`);
      continue;
    }

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
      const currentMonth = folderMonth || nameMonth;

      let entry = periods.get(path);
      if (!entry) {
        if (summary.parsed >= maxParse || Date.now() > deadline) { summary.unread++; summary.truncated = true; continue; }
        summary.parsed++;
        try {
          entry = await readInvoice(token, driveId, path);
        } catch (e) {
          entry = { path, read: false, note: e.message };
        }
        periods.set(path, entry);
        fresh.push(entry);
      }

      const bucket = `${vendor.name}||${relPath.split('/')[0] || ''}`;
      if (!folders.has(bucket)) {
        folders.set(bucket, {
          app: appOf(vendor.name),
          vendorFolder: vendor.name,
          month: currentMonth,
          monthFolderNames,
          files: [],
        });
      }
      folders.get(bucket).files.push({ ...entry, name: file.name, path, relPath, id: file.id, currentMonth, folderMonth, nameMonth });
    }
  }

  // Pass 2: which of them belong somewhere else.
  const movedOut = new Map();  // "vendor||month" -> the files leaving it
  const movedIn = new Map();   // "vendor||month" -> the files arriving in it

  for (const folder of folders.values()) {
    for (const file of folder.files) {
      const verdict = planMove(file, folder, base);
      if (!verdict) continue;
      if (verdict.skip) { summary.skipped.push({ file: file.name, folder: folder.vendorFolder, why: verdict.skip }); continue; }

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
      summary.cells = sheet ? predictCells(sheet, folders, movedOut, movedIn, appOf) : [];
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

// The month totals the moves imply, against what the sheet holds now. A month
// with any unread invoice in it is left out: a total built on a PDF nobody could
// read would be short, and writing it would look authoritative.
function predictCells(sheet, folders, movedOut, movedIn, appOf) {
  const rowByApp = new Map(sheet.grid.apps.map(a => [a.name, a.rowIdx]));

  // vendor||month -> the files that folder will hold once the moves are done
  const after = new Map();
  for (const folder of folders.values()) {
    if (!folder.month) continue;
    const key = `${folder.vendorFolder}||${folder.month}`;
    const leaving = new Set((movedOut.get(key) || []).map(f => f.path));
    after.set(key, {
      app: folder.app,
      month: folder.month,
      files: folder.files.filter(f => !leaving.has(f.path)),
    });
  }
  for (const [key, files] of movedIn) {
    const [vendorFolder, month] = key.split('||');
    if (!after.has(key)) after.set(key, { app: appOf(vendorFolder), month, files: [] });
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
    if (state.files.some(f => !f.read || (!f.usable && f.amount === null))) {
      cells.push({ app: state.app, month: state.month, blocked: 'an invoice in that month could not be read' });
      continue;
    }

    const total = Math.round(state.files.filter(f => f.usable).reduce((sum, f) => sum + f.amount, 0) * 100) / 100;
    const current = cellValue(sheet.values, rowIdx, colIdx);
    if (current !== null && Math.abs(current - total) < CELL_EPSILON) continue;

    cells.push({
      app: state.app,
      month: state.month,
      address: excel.cellAddress(sheet.start, rowIdx, colIdx),
      current,
      value: total,
      invoices: state.files.filter(f => f.usable).length,
      direction: current === null || total > current ? 'up' : 'down',
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
  const approved = new Set((Array.isArray(payload && payload.cells) ? payload.cells : [])
    .map(c => `${c.app}||${c.month}`));

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
    if (!approved.has(`${app}||${month}`)) { summary.cellsSkipped.push({ app, month, why: 'not approved' }); continue; }
    try {
      const res = await sumFolderInvoices(token, driveId, folder, cache, budget);
      if (res.unread.length || res.unusable.length) {
        summary.cellsSkipped.push({ app, month, why: 'an invoice in that month could not be read as a USD total' });
        continue;
      }
      totals.push({ app, month, value: res.total });
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
      const current = cellValue(live.values, rowIdx, colIdx);
      if (current !== null && Math.abs(current - t.value) < CELL_EPSILON) continue;
      cells.push({ app: t.app, month: t.month, address: excel.cellAddress(live.start, rowIdx, colIdx), value: t.value, before: current });
    }
    if (!cells.length) return;

    // writeCells carries each cell through to its result, so `before` — what the
    // sheet held — comes back with it and lands in the audit log unchanged.
    const results = await excel.writeCells(token, live.driveId, live.itemId, live.sheetName, cells, sessionId, 4);
    const ok = results.filter(r => r.ok);
    for (const r of ok) {
      summary.cellsWritten.push({ app: r.app, month: r.month, address: r.address, value: r.value, previous: r.before });
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

module.exports = { scanPeriods, applyBackfill, planMove, predictCells, readInvoice };
