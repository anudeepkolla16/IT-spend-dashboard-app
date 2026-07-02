const XLSX = require('xlsx');

const MONTH_RE = /^([A-Za-z]{3})-(\d{2})$/;
const MONTH_MAP = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };

function normMonthHeader(h) {
  const m = String(h || '').trim().match(MONTH_RE);
  if (!m) return null;
  const mon = MONTH_MAP[m[1].toLowerCase()];
  if (!mon) return null;
  return `20${m[2]}-${String(mon).padStart(2, '0')}`;
}

function cycleFromRow(recurringOnetime, frequency) {
  const ro = String(recurringOnetime || '').toLowerCase();
  const f = String(frequency || '').toLowerCase();
  if (ro.includes('one')) return 'One-time';
  const numMatch = f.match(/(\d+)\s*year/);
  if (numMatch) return `${numMatch[1]} Years`;
  if (f.includes('year') || f.includes('annual')) return 'Annual';
  if (f.includes('half')) return 'Half-Yearly';
  if (f.includes('quarter')) return 'Quarterly';
  if (f.includes('week')) return 'Weekly';
  return 'Monthly';
}

function inferCurrency(paymentMethod) {
  return /\bUS\b/i.test(String(paymentMethod || '')) ? 'USD' : 'INR';
}

async function getGraphToken() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET env vars');
  }
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
  });
  const res = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph token request failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const json = await res.json();
  return json.access_token;
}

async function downloadWorkbook(token) {
  const upn = process.env.TARGET_USER_UPN;
  const filePath = process.env.TARGET_FILE_PATH || 'Anudeep Excel sheets/Saras Apps & Subscriptions Purchase from Jan 26 .xlsx';
  if (!upn) throw new Error('Missing TARGET_USER_UPN env var');
  const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/drive/root:/${encodedPath}:/content`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph file download failed (${res.status}): ${text.slice(0, 300)}`);
  }
  return Buffer.from(await res.arrayBuffer());
}

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  let headerRowIdx = -1;
  let headers = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const hasMonthCol = row.some(c => normMonthHeader(c));
    const hasNameCol = row.some(c => /application|sw\s*\/\s*license/i.test(String(c)));
    if (hasMonthCol && hasNameCol) { headerRowIdx = i; headers = row; break; }
  }
  if (headerRowIdx === -1) throw new Error('Could not locate header row in the sheet');

  const colIdx = (patterns) => headers.findIndex(h => patterns.some(p => p.test(String(h))));
  const nameCol = colIdx([/application|sw\s*\/\s*license/i]);
  const deptCol = colIdx([/department/i]);
  const pocCol = colIdx([/poc/i]);
  const renewalCol = colIdx([/renewal/i]);
  const roCol = colIdx([/recurring\s*\/\s*onetime/i]);
  const freqCol = colIdx([/frequency/i]);
  const payCol = colIdx([/payment\s*method/i]);
  const monthCols = headers.map((h, idx) => ({ idx, month: normMonthHeader(h) })).filter(x => x.month);

  const records = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const name = String(row[nameCol] || '').trim();
    if (!name || /^total$/i.test(name)) continue;

    const dept = String(row[deptCol] || '').trim() || 'Unassigned';
    const poc = String(row[pocCol] || '').trim();
    const renewalDate = String(row[renewalCol] || '').trim();
    const recurringOnetime = String(row[roCol] || '').trim();
    const frequency = String(row[freqCol] || '').trim();
    const paymentMethod = String(row[payCol] || '').trim();
    const cycle = cycleFromRow(recurringOnetime, frequency);
    const cur = inferCurrency(paymentMethod);

    for (const { idx, month } of monthCols) {
      const raw = row[idx];
      const amt = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/,/g, ''));
      if (!amt || Number.isNaN(amt)) continue;
      records.push({ name, dept, poc, renewalDate, cycle, cur, month, amt, paymentMethod });
    }
  }
  return records;
}

module.exports = async (req, res) => {
  try {
    const token = await getGraphToken();
    const buffer = await downloadWorkbook(token);
    const rows = parseWorkbook(buffer);
    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json({ syncedAt: new Date().toISOString(), source: 'sharepoint', rowCount: rows.length, rows });
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: err.message || String(err) });
  }
};
