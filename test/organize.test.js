// Run with: node --test
//
// Tidying the archive's shape. The case it exists for: a whole vendor folder
// dragged inside another, leaving its month folders one level too deep. The sync
// totals {vendor}/{month}/ and nothing below it, so Luzmo's July invoices sat at
// Cumul(Luzmo)/Luzmo/July/ and counted towards no month at all.

const test = require('node:test');
const assert = require('node:assert');

const { planTidy, applyTidy } = require('../lib/invoices/organize');

const DRIVE = 'drive-1';
const ARCHIVE = 'Desktop/Anudeep files/Invoices';

// Mirrors the Graph shapes the real calls return: the archive probe by path, and
// children listings by item id.
function stub(tree, opts) {
  const original = global.fetch;
  const calls = [];
  const o = opts || {};
  global.fetch = async (url, init) => {
    const u = String(url);
    calls.push({ url: u, method: (init && init.method) || 'GET', body: init && init.body });

    const probe = u.match(/\/root:\/(.+?)(?:\?|$)/);
    if (probe && (!init || (init.method || 'GET') === 'GET')) {
      const path = decodeURIComponent(probe[1]).split('/').map(decodeURIComponent).join('/');
      if (path === ARCHIVE) return { ok: true, status: 200, json: async () => ({ id: 'root', folder: {} }) };
      if (tree.paths && tree.paths[path]) return { ok: true, status: 200, json: async () => ({ id: tree.paths[path], folder: {} }) };
      return { ok: false, status: 404, text: async () => 'itemNotFound' };
    }

    const kids = u.match(/items\/([^/]+)\/children/);
    if (kids) {
      const id = decodeURIComponent(kids[1]);
      if (!(id in tree.children)) return { ok: false, status: 403, text: async () => 'accessDenied' };
      return { ok: true, status: 200, json: async () => ({ value: tree.children[id] }) };
    }

    const patch = u.match(/items\/([^/?]+)$/);
    if (patch && init && init.method === 'PATCH') {
      const id = decodeURIComponent(patch[1]);
      if (o.failMove && o.failMove === id) return { ok: false, status: 409, text: async () => 'nameAlreadyExists' };
      (o.moved = o.moved || []).push({ id, to: JSON.parse(init.body).parentReference.id });
      return { ok: true, status: 200, json: async () => ({ id }) };
    }
    return { ok: false, status: 404, text: async () => 'not found' };
  };
  return { calls, restore: () => { global.fetch = original; } };
}

const folder = (id, name) => ({ id, name, folder: {} });
const file = (id, name) => ({ id, name, file: {}, size: 1024 });

// The archive exactly as it stands: Luzmo dragged inside Cumul(Luzmo), its July
// invoices stranded, and one invoice present in two places under one vendor.
const LUZMO = {
  paths: { [`${ARCHIVE}/Cumul(Luzmo)`]: 'v-cumul' },
  children: {
    root: [folder('v-cumul', 'Cumul(Luzmo)'), folder('v-un', '_Unmatched')],
    'v-cumul': [folder('c-aug', 'Aug-26'), folder('c-sep', 'Sep-26'), folder('c-luzmo', 'Luzmo')],
    'c-aug': [file('f-aug', '20260826_20260258.pdf')],
    'c-sep': [],
    'c-luzmo': [folder('l-aug', 'Aug-26'), folder('l-jul', 'July'), folder('l-sep', 'Sep-26')],
    'l-aug': [],
    'l-jul': [file('f-j1', '20260709_20260212.pdf'), file('f-j2', '20260712_20260215.pdf'), file('f-j3', 'July 26.pdf')],
    'l-sep': [file('f-sep', '20260826_20260258.pdf')],
  },
};

test('lifts a stranded month folder up to where the sync looks', async () => {
  const s = stub(LUZMO);
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.equal(plan.root, ARCHIVE);
    const july = plan.moves.find(m => m.month === '2026-07');
    assert.ok(july, 'July has to be lifted');
    assert.equal(july.kind, 'folder', 'nothing of July exists up top, so the folder moves whole');
    assert.equal(july.from, 'Cumul(Luzmo)/Luzmo/July');
    assert.equal(july.to, 'Cumul(Luzmo)/July');
    assert.equal(july.files.length, 3);
    assert.equal(july.destPath, `${ARCHIVE}/Cumul(Luzmo)`);
  } finally { s.restore(); }
});

test('refuses to lift a copy of an invoice already filed under that vendor', async () => {
  // 20260826_20260258.pdf is in Aug-26 AND in the stranded Sep-26. Moving the
  // second up would put the identical charge in two months at once.
  const s = stub(LUZMO);
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.equal(plan.moves.some(m => m.month === '2026-09'), false, 'the September copy must not move');
    assert.equal(plan.duplicates.length, 1);
    assert.deepEqual(plan.duplicates[0].files, ['20260826_20260258.pdf']);
    assert.deepEqual(plan.duplicates[0].alreadyAt, ['Cumul(Luzmo)/Aug-26']);
  } finally { s.restore(); }
});

test('an empty stranded folder is reported, not moved', async () => {
  const s = stub(LUZMO);
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.deepEqual(plan.empties.map(e => e.path), ['Cumul(Luzmo)/Luzmo/Aug-26']);
    assert.equal(plan.moves.some(m => m.from.endsWith('Luzmo/Aug-26')), false);
  } finally { s.restore(); }
});

test('a folder that is not a month is left entirely alone', async () => {
  // "Quotations" under Laptop procurment holds dated subfolders that are not
  // months, and "Max accounts" under Claude Ai is a different app's invoices.
  // Neither is this job's business.
  const s = stub({
    paths: { [`${ARCHIVE}/Laptop procurment`]: 'v-lap' },
    children: {
      root: [folder('v-lap', 'Laptop procurment')],
      'v-lap': [folder('q', 'Quotations'), file('f', 'mac 29 apr 2026.pdf')],
      q: [folder('q1', '12-08-2026'), folder('q2', '05-08-2026')],
      q1: [file('qf1', 'quote.pdf')],
      q2: [],
    },
  });
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.deepEqual(plan.moves, []);
    assert.deepEqual(plan.duplicates, []);
  } finally { s.restore(); }
});

test('a vendor with no stray subfolders is not even descended into', async () => {
  const s = stub({
    children: {
      root: [folder('v', 'Adobe')],
      v: [folder('m', 'Aug-26'), file('f', 'jan 26.pdf')],
      m: [file('f2', 'x.pdf')],
    },
  });
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.deepEqual(plan.moves, []);
    // The month folder is never listed, because there is no stray to compare it to.
    assert.equal(s.calls.some(c => c.url.includes('items/m/children')), false);
  } finally { s.restore(); }
});

test('an unresolved archive plans nothing rather than guessing a root', async () => {
  const original = global.fetch;
  global.fetch = async () => ({ ok: false, status: 404, text: async () => 'itemNotFound' });
  try {
    const plan = await planTidy('tok', DRIVE);
    assert.equal(plan.root, null);
    assert.deepEqual(plan.moves, []);
    assert.ok(plan.triedPaths.length);
  } finally { global.fetch = original; }
});

/* ---------------- applying ---------------- */

test('applies only the moves it was given, and reports what it did', async () => {
  const o = {};
  const s = stub(LUZMO, o);
  try {
    const plan = await planTidy('tok', DRIVE);
    const out = await applyTidy('tok', DRIVE, plan.moves);
    assert.equal(out.moved.length, 1);
    assert.equal(out.failed.length, 0);
    assert.deepEqual(o.moved, [{ id: 'l-jul', to: 'v-cumul' }]);
  } finally { s.restore(); }
});

test('a destination outside the archive is refused', async () => {
  // The plan is echoed back by the browser, so an edited payload must not be
  // able to move an invoice somewhere else on the drive.
  const s = stub(LUZMO, {});
  try {
    const out = await applyTidy('tok', DRIVE, [
      { itemId: 'l-jul', destPath: 'Desktop/Anudeep files', from: 'a', to: 'b' },
      // Starts with the archive path and still escapes it — a prefix test alone
      // would wave this through.
      { itemId: 'l-jul', destPath: `${ARCHIVE}/../elsewhere`, from: 'a', to: 'b' },
      { itemId: 'l-jul', destPath: `${ARCHIVE}/./Cumul(Luzmo)`, from: 'a', to: 'b' },
      { itemId: 'l-jul', destPath: `${ARCHIVE}//Cumul(Luzmo)`, from: 'a', to: 'b' },
    ]);
    assert.equal(out.moved.length, 0);
    assert.equal(out.failed.length, 4);
    out.failed.forEach(f => assert.match(f.error, /outside the archive/));
    // And nothing reached Graph: the refusal happens before any request.
    assert.equal(s.calls.some(c => c.method === 'PATCH'), false);
  } finally { s.restore(); }
});

test('a clash at the destination fails that one move and no other', async () => {
  const s = stub(LUZMO, { failMove: 'l-jul' });
  try {
    const out = await applyTidy('tok', DRIVE, [
      { itemId: 'l-jul', destPath: `${ARCHIVE}/Cumul(Luzmo)`, from: 'a', to: 'b', kind: 'folder' },
      { itemId: 'f-j1', destPath: `${ARCHIVE}/Cumul(Luzmo)`, from: 'c', to: 'd', kind: 'file' },
    ]);
    assert.equal(out.moved.length, 1);
    assert.equal(out.failed.length, 1);
    assert.match(out.failed[0].error, /already in the destination/);
  } finally { s.restore(); }
});

test('nothing is deleted, ever', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'lib', 'invoices', 'organize.js'), 'utf8');
  assert.equal(/method:\s*'DELETE'/.test(src), false, 'tidying must never issue a delete');
  assert.equal(/deleteItem|removeItem/.test(src), false);
});
