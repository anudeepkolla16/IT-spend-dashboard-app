// Run with: node --test
//
// Where the invoice archive lives is resolved against the drive, never assumed.
// It was assumed once — every path hardcoded `Invoices/` at the drive root while
// the folder actually sat at `Desktop/Anudeep files/Procurment bills`, later
// renamed to `Desktop/Anudeep files/Invoices`. Nothing errored; reads just
// returned nothing, and the dashboard reported all 293 charged months as missing
// an invoice. These tests pin the behaviour that stops that recurring.

const test = require('node:test');
const assert = require('node:assert');

const { archiveCandidates, resolveArchiveRoot, archiveFile } = require('../lib/graph');

const ENV_KEYS = ['INVOICE_ARCHIVE_PATH', 'INVOICE_SOURCE_PATH'];
function clearEnv() { for (const k of ENV_KEYS) delete process.env[k]; }

// Stub Graph so only the named paths exist as folders. Records the paths asked
// for, in order, so the probe order itself can be asserted.
function stubDrive(existing, opts) {
  const original = global.fetch;
  const asked = [];
  const o = opts || {};
  global.fetch = async (url) => {
    const m = String(url).match(/\/root:\/(.+?)\?/);
    const path = m ? decodeURIComponent(m[1]).split('/').map(decodeURIComponent).join('/') : '';
    asked.push(path);
    if (o.failOn && o.failOn === path) {
      return { ok: false, status: 503, text: async () => 'service unavailable' };
    }
    if (Object.prototype.hasOwnProperty.call(existing, path)) {
      return { ok: true, status: 200, json: async () => existing[path] };
    }
    return { ok: false, status: 404, text: async () => 'itemNotFound' };
  };
  return { asked, restore: () => { global.fetch = original; } };
}

const folderItem = id => ({ id, folder: {} });

// Each test uses its own driveId so the module-level cache never leaks between them.
let n = 0;
const nextDrive = () => `drive-${++n}`;

test('candidate order puts explicit config first and the live folder before the old name', () => {
  clearEnv();
  assert.deepEqual(archiveCandidates(), [
    'Desktop/Anudeep files/Invoices',
    'Desktop/Anudeep files/Procurment bills',
    'Invoices',
  ]);

  process.env.INVOICE_ARCHIVE_PATH = 'Some/Explicit/Path/';
  process.env.INVOICE_SOURCE_PATH = 'Legacy/Path';
  assert.deepEqual(archiveCandidates(), [
    'Some/Explicit/Path',            // trailing slash trimmed
    'Legacy/Path',
    'Desktop/Anudeep files/Invoices',
    'Desktop/Anudeep files/Procurment bills',
    'Invoices',
  ]);
  clearEnv();
});

test('duplicate candidates collapse rather than being probed twice', () => {
  clearEnv();
  process.env.INVOICE_ARCHIVE_PATH = 'Invoices';
  process.env.INVOICE_SOURCE_PATH = 'Invoices';
  const c = archiveCandidates();
  assert.deepEqual(c.filter(p => p === 'Invoices').length, 1);
  clearEnv();
});

test('finds the archive at its real location, not the hardcoded root', async () => {
  clearEnv();
  const stub = stubDrive({ 'Desktop/Anudeep files/Invoices': folderItem('arch-1') });
  try {
    const root = await resolveArchiveRoot('tok', nextDrive());
    assert.equal(root.path, 'Desktop/Anudeep files/Invoices');
    assert.equal(root.itemId, 'arch-1');
    assert.equal(root.resolved, true);
  } finally { stub.restore(); }
});

test('still finds the archive under its pre-rename name', async () => {
  clearEnv();
  const stub = stubDrive({ 'Desktop/Anudeep files/Procurment bills': folderItem('arch-2') });
  try {
    const root = await resolveArchiveRoot('tok', nextDrive());
    assert.equal(root.path, 'Desktop/Anudeep files/Procurment bills');
    assert.equal(root.resolved, true);
    // The renamed location is checked first and simply isn't there.
    assert.deepEqual(stub.asked, ['Desktop/Anudeep files/Invoices', 'Desktop/Anudeep files/Procurment bills']);
  } finally { clearEnv(); stub.restore(); }
});

test('an explicit env path wins over both defaults', async () => {
  clearEnv();
  process.env.INVOICE_ARCHIVE_PATH = 'Elsewhere/Bills';
  const stub = stubDrive({
    'Elsewhere/Bills': folderItem('arch-env'),
    'Desktop/Anudeep files/Invoices': folderItem('arch-default'),
  });
  try {
    const root = await resolveArchiveRoot('tok', nextDrive());
    assert.equal(root.itemId, 'arch-env');
    assert.deepEqual(stub.asked, ['Elsewhere/Bills']);   // stops at the first hit
  } finally { clearEnv(); stub.restore(); }
});

test('a file of the archive name is not mistaken for the archive', async () => {
  clearEnv();
  const stub = stubDrive({
    'Desktop/Anudeep files/Invoices': { id: 'not-a-folder', file: {} },
    'Invoices': folderItem('arch-root'),
  });
  try {
    const root = await resolveArchiveRoot('tok', nextDrive());
    assert.equal(root.itemId, 'arch-root');
  } finally { stub.restore(); }
});

test('when nothing exists it reports the miss instead of pretending it resolved', async () => {
  clearEnv();
  const stub = stubDrive({});
  try {
    const root = await resolveArchiveRoot('tok', nextDrive());
    assert.equal(root.resolved, false);
    assert.equal(root.itemId, null);
    // A writer still needs somewhere to create; a reader must check `resolved`.
    assert.equal(root.path, 'Desktop/Anudeep files/Invoices');
    assert.deepEqual(root.candidates.length, 3);
  } finally { stub.restore(); }
});

test('an unresolved archive is not cached, so the folder appearing is picked up', async () => {
  clearEnv();
  const drive = nextDrive();
  let stub = stubDrive({});
  try {
    assert.equal((await resolveArchiveRoot('tok', drive)).resolved, false);
  } finally { stub.restore(); }

  stub = stubDrive({ 'Desktop/Anudeep files/Invoices': folderItem('arch-late') });
  try {
    const root = await resolveArchiveRoot('tok', drive);
    assert.equal(root.resolved, true, 'a negative result must not be cached');
    assert.equal(root.itemId, 'arch-late');
  } finally { stub.restore(); }
});

test('a transient Graph error fails loudly rather than filing into the next candidate', async () => {
  clearEnv();
  const stub = stubDrive(
    { 'Desktop/Anudeep files/Procurment bills': folderItem('arch-old') },
    { failOn: 'Desktop/Anudeep files/Invoices' }
  );
  try {
    await assert.rejects(
      () => resolveArchiveRoot('tok', nextDrive()),
      /failed \(503\)/,
      'falling through on a 503 is how invoices end up filed in the wrong folder'
    );
  } finally { stub.restore(); }
});

test('a resolved archive is cached, and `fresh` re-probes it', async () => {
  clearEnv();
  const drive = nextDrive();
  let stub = stubDrive({ 'Desktop/Anudeep files/Invoices': folderItem('arch-a') });
  try {
    await resolveArchiveRoot('tok', drive);
    const before = stub.asked.length;
    await resolveArchiveRoot('tok', drive);
    assert.equal(stub.asked.length, before, 'a cached hit must not re-probe');
    await resolveArchiveRoot('tok', drive, { fresh: true });
    assert.ok(stub.asked.length > before, '`fresh` must re-probe');
  } finally { stub.restore(); }
});

test('bookkeeping files sit beside the invoices, wherever those are', () => {
  const root = { path: 'Desktop/Anudeep files/Invoices' };
  assert.equal(archiveFile(root, '_invoice-index.json'), 'Desktop/Anudeep files/Invoices/_invoice-index.json');
  assert.equal(archiveFile(root, '_mail-sync.json'), 'Desktop/Anudeep files/Invoices/_mail-sync.json');
});
