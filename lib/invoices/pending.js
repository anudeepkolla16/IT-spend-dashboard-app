// Invoices the app is not sure about, held until the owner answers.
//
// Rule 4 of the owner's process: "all the invoices should go to particular
// month folder only, never guess or mismatch — if you are not clear ask me".
// So an invoice that no rule places, or whose month cannot be read, is not
// filed anywhere it might be wrong. Its bytes go into `_Pending/` in the
// archive — a mail can drop out of the sync window before it is answered —
// and a question goes into `_pending.json`, which the dashboard shows and the
// Slack DM repeats. An answer, from either place, files it properly.
//
// The state file:
//   { items: [ { id: 'P12', file, heldPath, question, options, … } ],
//     nextId: 13,
//     slack: { lastTs } }        — how far the DM has been read
const { readJsonFile, writeJsonFile, uploadFileContent, archiveFile, itemIdByPath, ensureFolder, moveItem } = require('../graph');
const { norm } = require('../vendor-map');
const { parseMonthFolder } = require('./inventory');

const PENDING_FILE = '_pending.json';
const HOLD_FOLDER = '_Pending';
const MAX_ITEMS = 500;

function emptyState() {
  return { items: [], nextId: 1, slack: { lastTs: null } };
}

async function readPending(token, driveId, root) {
  const raw = (await readJsonFile(token, driveId, archiveFile(root, PENDING_FILE))) || {};
  const state = emptyState();
  if (Array.isArray(raw.items)) state.items = raw.items.filter(i => i && typeof i === 'object' && i.id);
  if (Number.isInteger(raw.nextId) && raw.nextId > 0) state.nextId = raw.nextId;
  if (raw.slack && typeof raw.slack === 'object') state.slack = { lastTs: raw.slack.lastTs || null };
  return state;
}

async function writePending(token, driveId, root, state) {
  await writeJsonFile(token, driveId, archiveFile(root, PENDING_FILE), {
    items: state.items.slice(-MAX_ITEMS), nextId: state.nextId, slack: state.slack || { lastTs: null },
    updatedAt: new Date().toISOString(),
  });
}

// One held file is one item, and a file held twice (the same mail read again
// on a rescan) is still one item.
function findHeld(state, fileName, messageId) {
  return state.items.find(i => i.file === fileName && (messageId ? i.messageId === messageId : true)) || null;
}

// Parks the PDF and records the question. `meta` is everything the run knew
// about the invoice, so the dashboard and the DM can show it and the answer
// can be learned from it (see rules.learn).
async function holdInvoice(token, driveId, root, state, bytes, meta) {
  const existing = findHeld(state, meta.file, meta.messageId);
  if (existing) return { item: existing, already: true };
  const id = `P${state.nextId++}`;
  const heldPath = `${root.path}/${HOLD_FOLDER}/${id}-${meta.file}`;
  await uploadFileContent(token, driveId, heldPath, bytes, 'application/pdf');
  const item = { id, heldPath, createdAt: new Date().toISOString(), ...meta };
  state.items.push(item);
  return { item, already: false };
}

function removeItem(state, id) {
  const idx = state.items.findIndex(i => i.id === id);
  if (idx === -1) return null;
  return state.items.splice(idx, 1)[0];
}

// Moves a held PDF into its proper folder. The folder is created if it is not
// there yet; the file keeps its original name.
async function moveHeld(token, driveId, item, destFolderPath) {
  const itemId = await itemIdByPath(token, driveId, item.heldPath);
  if (!itemId) throw new Error(`the held file is no longer at "${item.heldPath}"`);
  const parentId = await ensureFolder(token, driveId, destFolderPath);
  await moveItem(token, driveId, itemId, parentId, item.file);
  return `${destFolderPath}/${item.file}`;
}

// --- Reading an answer -----------------------------------------------------
//
// From the dashboard an answer arrives as { id, app, month } and needs no
// parsing. From Slack it is a line of text; the DM says how to write it:
//
//   P12 = Google Voice
//   P12 = 2                 (the second option the question listed)
//   P12 = Google Voice, Aug-26
//   P12 = ignore            (not an invoice for the sheet; leave it in _Pending)
//
// Anything that does not read cleanly is returned as an error naming the line,
// never guessed at, and the next DM says so.

const REPLY_LINE = /^\s*#?\s*(P\s*\d+)\s*(?:=|:|->|→|-|–|—)\s*(.+?)\s*$/i;

function parseMonthToken(tok, now) {
  const s = String(tok || '').trim();
  if (!s) return null;
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  return parseMonthFolder(s, now) || null;
}

// An app name, matched the way the rest of the app matches names: normalized
// exact first, then a single unambiguous containment. Two apps containing the
// typed text resolve to neither.
function matchApp(text, appNames) {
  const key = norm(text);
  if (!key) return { app: null, error: 'no app named' };
  const exact = appNames.find(a => norm(a) === key);
  if (exact) return { app: exact };
  const contains = appNames.filter(a => norm(a).includes(key) || key.includes(norm(a)));
  if (contains.length === 1) return { app: contains[0] };
  if (contains.length > 1) return { app: null, error: `"${text}" could be ${contains.slice(0, 4).join(', ')} — say which` };
  return { app: null, error: `"${text}" is not a row in the sheet` };
}

function parseReply(text, items, appNames, now) {
  const out = [];
  for (const line of String(text || '').split(/\r?\n/)) {
    const m = line.match(REPLY_LINE);
    if (!m) continue;
    const id = m[1].replace(/\s+/g, '').toUpperCase();
    const item = items.find(i => i.id === id);
    if (!item) { out.push({ id, error: `${id} is not an open question`, line }); continue; }
    const answer = m[2].trim();
    if (/^(ignore|skip|not an invoice|drop)$/i.test(answer)) { out.push({ id, ignore: true, line }); continue; }

    // "Google Voice, Aug-26" / "Google Voice Aug-26" / "2 Aug 2026"
    let appText = answer;
    let month = null;
    const parts = answer.split(/\s*[,;]\s*/);
    if (parts.length === 2) {
      month = parseMonthToken(parts[1], now);
      if (month) appText = parts[0];
    }
    if (!month) {
      const tail = answer.match(/^(.*?)\s+((?:[A-Za-z]{3,9}[-\s]?\d{2,4})|(?:\d{4}-\d{2}))$/);
      if (tail) {
        const mm = parseMonthToken(tail[2], now);
        if (mm) { month = mm; appText = tail[1]; }
      }
    }

    let app = null;
    const n = Number(appText);
    if (Number.isInteger(n) && n >= 1 && Array.isArray(item.options) && item.options[n - 1]) {
      app = item.options[n - 1];
    } else {
      const r = matchApp(appText, appNames);
      if (!r.app) { out.push({ id, error: r.error, line }); continue; }
      app = r.app;
    }
    out.push({ id, app, month: month || item.month || null, line });
  }
  return out;
}

// The DM's wording for the open questions, appended to every run report while
// anything is waiting. Kept short and the reply grammar stated once.
function describeQuestions(items) {
  if (!items.length) return '';
  const lines = [`*${items.length} invoice${items.length === 1 ? '' : 's'} waiting for your answer:*`];
  for (const it of items.slice(0, 20)) {
    const when = it.receivedAt ? ` · ${String(it.receivedAt).slice(0, 10)}` : '';
    lines.push(`• *${it.id}* — ${it.file}${when}${it.subject ? ` · "${String(it.subject).slice(0, 60)}"` : ''}${it.from ? ` · from ${it.from}` : ''}`);
    lines.push(`   ${it.question}`);
    if (Array.isArray(it.options) && it.options.length) lines.push(`   options: ${it.options.map((o, i) => `${i + 1}) ${o}`).join('   ')}`);
    if (!it.month) lines.push(`   (and the month: nothing in the invoice says which month it bills)`);
  }
  if (items.length > 20) lines.push(`… and ${items.length - 20} more on the dashboard`);
  lines.push('');
  lines.push('Reply one per line, e.g. `P12 = Google Voice` or `P12 = 2` (option number), add the month if asked: `P12 = Google Voice, Aug-26`. `P12 = ignore` leaves it out of the sheet.');
  return lines.join('\n');
}

module.exports = { PENDING_FILE, HOLD_FOLDER, readPending, writePending, holdInvoice, findHeld, removeItem, moveHeld, parseReply, matchApp, describeQuestions };
