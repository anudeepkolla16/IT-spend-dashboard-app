// Run with: node --test
//
// The Agreements & KYC drawer. Two things matter here and nothing else does:
// that a file name is sorted into the right kind (an order form is not a bill,
// a cancelled cheque is not a contract), and that the folder's real contents
// come back grouped by vendor with the term dates read off the names.
//
// The fixtures below are the live folder's actual file names, misspellings and
// all — "Antropic Aggrement.pdf" is what is on the drive.

const test = require('node:test');
const assert = require('node:assert');

// Stub the Graph layer before lib/documents captures its exports.
const graph = require('../lib/graph');

const ROOT = 'Desktop/Anudeep files/Aggrements and Kyc';
const ROOT_ID = 'root-1';

// Vendor folder -> the files in it, as listFilesRecursive would report them.
const FOLDERS = {
  'Ar Enterprices': ['Fw_ Documents.zip'],
  'Claude': ['Antropic Aggrement.pdf'],
  'Clickup': ['Order Form may 2026 to may 2027.pdf'],
  'Hubspot': ['Your Hubspot Purchase Order #9373788.pdf'],
  'Microsoft': [
    'Aayusha IT Services & Solutions.pdf',
    'Cancelled cheque Aayusha.pdf',
    'Fw_ Need below KYC Documents .zip',
    'GST Certificate.pdf',
    'Kavita Pan Card.jpeg',
    'Microsoft 365 Business Basic with Email Security.pdf',
    'O365_Escalation_Matrix.pdf',
    'Udyam Certificate.pdf',
  ],
  'Slack': ['Order Form_Q-11968953.pdf'],
  'Sprinto': ['DealRoom_for_Saras_Analytics_Inc-Q-24888-8-Sep-2026-9-28-30-signed.pdf'],
};

// Only the real path resolves; the other candidates 404, as they would live.
// lib/documents destructures the Graph helpers when it loads, so this is the one
// stub for the whole file — the "no folder anywhere" case is keyed off a drive id
// rather than by swapping the stub afterwards, which would not take.
graph.graphFetch = async url => {
  const found = !url.includes('drive-docs-missing')
    && decodeURIComponent(url).includes(`root:/${ROOT}?`);
  return {
    status: found ? 200 : 404,
    ok: found,
    json: async () => ({ id: ROOT_ID, folder: {}, webUrl: 'https://sp/root' }),
    text: async () => '',
  };
};
graph.graphListAll = async () => Object.keys(FOLDERS).map((name, i) => ({
  id: `f${i}`, name, folder: {}, webUrl: `https://sp/${encodeURIComponent(name)}`,
}));
graph.listFilesRecursive = async (_t, _d, folderId) => {
  const name = Object.keys(FOLDERS)[Number(String(folderId).slice(1))];
  return (FOLDERS[name] || []).map((f, i) => ({
    id: `${folderId}-${i}`, name: f, size: 1024 * (i + 1),
    lastModifiedDateTime: '2026-09-08T07:00:00Z', webUrl: `https://sp/${folderId}/${i}`, relPath: '',
  }));
};

const { classify, extractTerm, buildDocumentIndex, documentCandidates } = require('../lib/documents');

test('every file in the live folder is sorted into the right kind', () => {
  const expected = {
    'Order Form may 2026 to may 2027.pdf': 'order',
    'Order Form_Q-11968953.pdf': 'order',
    'Your Hubspot Purchase Order #9373788.pdf': 'order',
    // Spelled "Aggrement" on the drive — the classifier has to read it anyway.
    'Antropic Aggrement.pdf': 'agreement',
    'DealRoom_for_Saras_Analytics_Inc-Q-24888-8-Sep-2026-9-28-30-signed.pdf': 'agreement',
    'GST Certificate.pdf': 'kyc',
    'Udyam Certificate.pdf': 'kyc',
    'Kavita Pan Card.jpeg': 'kyc',
    'Cancelled cheque Aayusha.pdf': 'kyc',
    'Fw_ Need below KYC Documents .zip': 'kyc',
    'Aayusha IT Services & Solutions.pdf': 'other',
    'O365_Escalation_Matrix.pdf': 'other',
    'Fw_ Documents.zip': 'other',
    'Microsoft 365 Business Basic with Email Security.pdf': 'other',
  };
  for (const [name, kind] of Object.entries(expected)) {
    assert.equal(classify(name), kind, `"${name}" should be ${kind}`);
  }
});

test('a cancelled cheque is never read as an order form', () => {
  // "cheque"/"PO"/"quote" overlap enough that the KYC check has to come first.
  assert.equal(classify('Cancelled cheque Aayusha.pdf'), 'kyc');
  assert.equal(classify('PAN card.pdf'), 'kyc');
});

test('a two-ended term is read off the file name', () => {
  assert.deepEqual(extractTerm('Order Form may 2026 to may 2027.pdf'), { start: '2026-05', end: '2027-05' });
  assert.deepEqual(extractTerm('MSA Jan 2026 - Dec 2026.pdf'), { start: '2026-01', end: '2026-12' });
  assert.deepEqual(extractTerm('Contract 2026-04-01 to 2027-03-31.pdf'), { start: '2026-04-01', end: '2027-03-31' });
});

test('a single date is not guessed at as an expiry', () => {
  // A lone date is as likely to be when it was signed as when it runs out, and
  // showing a wrong expiry is worse than showing none.
  assert.equal(extractTerm('DealRoom_for_Saras_Analytics_Inc-Q-24888-8-Sep-2026-9-28-30-signed.pdf'), null);
  assert.equal(extractTerm('Your Hubspot Purchase Order #9373788.pdf'), null);
  assert.equal(extractTerm('GST Certificate.pdf'), null);
  // End before start is not a term either.
  assert.equal(extractTerm('Order Form may 2027 to may 2026.pdf'), null);
});

test('the folder path is probed, not assumed', () => {
  const candidates = documentCandidates();
  assert.ok(candidates.includes(ROOT), 'the live folder must be among the candidates');
  assert.ok(candidates.length > 1, 'a single hardcoded path is what this is meant to avoid');
});

test('the index groups every vendor and counts the kinds', async () => {
  const data = await buildDocumentIndex('tok', 'drive-docs-1', { fresh: true });
  assert.equal(data.resolved, true);
  assert.equal(data.root, ROOT);
  assert.equal(data.truncated, false);
  assert.equal(data.totals.vendors, 7);
  assert.equal(data.totals.files, 14);
  assert.equal(data.totals.order, 3);
  assert.equal(data.totals.agreement, 2);
  assert.equal(data.totals.kyc, 5);
  assert.equal(data.totals.other, 4);

  assert.deepEqual(data.vendors.map(v => v.vendor),
    ['Ar Enterprices', 'Claude', 'Clickup', 'Hubspot', 'Microsoft', 'Slack', 'Sprinto']);

  const clickup = data.vendors.find(v => v.vendor === 'Clickup');
  assert.equal(clickup.files[0].type, 'order');
  assert.deepEqual(clickup.files[0].term, { start: '2026-05', end: '2027-05' });
  assert.ok(clickup.files[0].webUrl, 'a document is no use without a link to it');
});

test('a missing folder says so instead of reporting an empty drawer', async () => {
  const data = await buildDocumentIndex('tok', 'drive-docs-missing', { fresh: true });
  assert.equal(data.resolved, false);
  assert.equal(data.vendors.length, 0);
  assert.equal(data.totals.files, 0);
  assert.ok(data.candidates.length, 'the page tells you which paths were tried');
});
