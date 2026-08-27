const { resolveDriveId, readJsonFile, writeJsonFile, resolveArchiveRoot, archiveFile } = require('./graph');

const excel = require('./excel');

// The alias map and audit log live beside the invoices, wherever the archive is.
const mapPath = root => archiveFile(root, '_amount-map.json');
const logPath = root => archiveFile(root, '_amount-log.json');

function spendFilePath() {
  return (process.env.TARGET_FILE_PATH || 'Anudeep Excel sheets/Saras Apps & Subscriptions Purchase from Jan 26 .xlsx').trim();
}

// Opens the spend workbook and works out where everything lives: which sheet
// holds the amounts, which row each app is on, which column each month is.
// Both preview and apply re-derive this from the live file rather than trusting
// anything the browser sends, so a row inserted between the two never causes a
// write to land on the wrong app.
async function openSpendSheet(token, sessionId) {
  const upn = (process.env.TARGET_USER_UPN || '').trim();
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

  const driveId = await resolveDriveId(token, upn);
  const itemId = await excel.resolveItemId(token, driveId, spendFilePath());
  const sheets = await excel.listWorksheets(token, driveId, itemId, sessionId);
  if (!sheets.length) throw new Error('The spend workbook has no worksheets');
  const sheetName = excel.pickSpendSheet(sheets.map(s => s.name));

  const used = await excel.readUsedRange(token, driveId, itemId, sheetName, sessionId);
  const grid = excel.locateGrid(used.values, used.text);
  return { driveId, itemId, sheetName, values: used.values, text: used.text, formulas: used.formulas, start: used.start, grid };
}

function cellValue(values, rowIdx, colIdx) {
  const row = values[rowIdx] || [];
  const raw = row[colIdx];
  if (typeof raw === 'number') return raw;
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  const n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return Number.isNaN(n) ? null : n;
}

// Trailing average of the months before `month` for one app, used to spot a
// figure that is wildly out of line with what this app normally costs.
function trailingAverage(values, grid, rowIdx, month, lookback) {
  const months = Object.keys(grid.monthCols).sort();
  const idx = months.indexOf(month);
  if (idx <= 0) return null;
  const prior = months.slice(Math.max(0, idx - (lookback || 3)), idx);
  const nums = prior.map(m => cellValue(values, rowIdx, grid.monthCols[m])).filter(v => typeof v === 'number' && v > 0);
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

async function readAliasMap(token, driveId) {
  const root = await resolveArchiveRoot(token, driveId);
  const json = await readJsonFile(token, driveId, mapPath(root));
  return (json && json.aliases) || {};
}

async function writeAliasMap(token, driveId, aliases) {
  const root = await resolveArchiveRoot(token, driveId);
  await writeJsonFile(token, driveId, mapPath(root), { aliases, savedAt: new Date().toISOString() });
}

// Append-only audit trail: every write the dashboard makes to the spend sheet is
// recorded with who did it and what the cell held before, so a wrong number can
// always be traced and undone.
async function appendLog(token, driveId, entry) {
  const root = await resolveArchiveRoot(token, driveId);
  let log = null;
  try {
    log = await readJsonFile(token, driveId, logPath(root));
  } catch (_) {
    log = null;
  }
  const entries = (log && Array.isArray(log.entries) ? log.entries : []).concat([entry]);
  // Keep the file small enough to stay a cheap read.
  const trimmed = entries.slice(-200);
  await writeJsonFile(token, driveId, logPath(root), { entries: trimmed, updatedAt: new Date().toISOString() });
}

async function readLog(token, driveId) {
  const root = await resolveArchiveRoot(token, driveId);
  return readJsonFile(token, driveId, logPath(root));
}

module.exports = { openSpendSheet, cellValue, trailingAverage, readAliasMap, writeAliasMap, appendLog, readLog, spendFilePath, mapPath, logPath };
