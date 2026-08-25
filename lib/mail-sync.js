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

function sanitizeFileName(name) {
  const dot = String(name || '').lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : (name || 'invoice');
  const ext = dot > 0 ? name.slice(dot).replace(/[<>:"/\\|?*\x00-\x1F]/g, '') : '.pdf';
  return sanitizeSegment(base) + ext;
}

async function folderFileNames(token, driveId, path) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/children?$select=name&$top=200`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return new Set();
  if (!res.ok) return new Set();
  const json = await res.json();
  return new Set((json.value || []).map(f => f.name));
}

// Ticks the app's month in the "Invoices tracker" sheet. That sheet has the same
// app-rows/month-columns shape as the spend sheet but holds TRUE/FALSE.
async function markTracker(token, driveId, itemId, marks, sessionId) {
  const sheets = await excel.listWorksheets(token, driveId, itemId, sessionId);
  const tracker = sheets.map(s => s.name).find(n => /tracker/i.test(n));
  if (!tracker) return { sheet: null, marked: 0 };

  const used = await excel.readUsedRange(token, driveId, itemId, tracker, sessionId);
  const grid = excel.locateGrid(used.values, used.text);
  const rowByApp = new Map(grid.apps.map(a => [a.name, a.rowIdx]));

  const cells = [];
  for (const { app, month } of marks) {
    const rowIdx = rowByApp.get(app);
    const colIdx = grid.monthCols[month];
    if (rowIdx === undefined || colIdx === undefined) continue;
    const existing = String((used.values[rowIdx] || [])[colIdx] == null ? '' : (used.values[rowIdx] || [])[colIdx]).trim();
    if (/^true$/i.test(existing)) continue; // already ticked
    cells.push({ app, month, address: excel.cellAddress(used.start, rowIdx, colIdx), value: true });
  }
  if (!cells.length) return { sheet: tracker, marked: 0 };

  const results = await excel.writeCells(token, driveId, itemId, tracker, cells, sessionId, 4);
  return { sheet: tracker, marked: results.filter(r => r.ok).length };
}

// Returns a summary of what was filed. `deadline` is a timestamp the run stops
// at, so the caller can share one function timeout across several jobs.
async function runMailSync(token, driveId, options) {
  const opts = options || {};
  const mailbox = mail.mailboxAddress();

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
  };
  const indexEntries = [];
  const marks = [];
  const deadline = opts.deadline || (Date.now() + DEADLINE_MS);

  for (const message of messages) {
    if (Date.now() > deadline) { summary.timedOut = true; break; }
    if (seen.has(message.id)) continue;

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
    const targetFolder = app
      ? `Invoices/${sanitizeSegment(app)}/${month}`
      : `Invoices/${UNMATCHED_FOLDER}/${month}`;
    const existing = await folderFileNames(token, driveId, targetFolder);

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
      if (month) marks.push({ app, month });
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

module.exports = { runMailSync, STATE_PATH, INDEX_PATH };
