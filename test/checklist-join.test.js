// Run with: node --test
//
// The checklist joins two different naming systems: the archive's folders carry
// VENDOR names ("Bubble", "Claude Api") and the spend sheet's rows carry APP
// names ("Bubble Starter", "Anthropic(Api Console)"). Joining on the raw name
// reports invoices the archive is holding as missing — which is exactly the
// failure this tab exists to catch, so the resolver is worth pinning down.
//
// The function under test lives in the dashboard's inline script, so it is
// lifted out of index.html by source rather than re-implemented here. A copy
// would drift from the page and prove nothing.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function lift(name) {
  const start = html.indexOf(`function ${name}(`);
  assert.ok(start > -1, `${name}() is no longer in index.html — this test is stale`);
  // Walk braces to the function's closing brace.
  let i = html.indexOf('{', start), depth = 0, end = -1;
  for (; i < html.length; i++) {
    if (html[i] === '{') depth++;
    else if (html[i] === '}' && --depth === 0) { end = i + 1; break; }
  }
  assert.ok(end > -1, `couldn't find the end of ${name}()`);
  return html.slice(start, end);
}

const invNorm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
// INV_FILING_WORDS is a const the function closes over, so it is lifted too.
const FILING_WORDS_SRC = html.match(/const INV_FILING_WORDS = [^\n]+/);
assert.ok(FILING_WORDS_SRC, 'INV_FILING_WORDS is no longer in index.html — this test is stale');
const invResolveFolder = new Function(
  'invNorm',
  `${FILING_WORDS_SRC[0]}\n${lift('invResolveFolder')}; return invResolveFolder;`
)(invNorm);

// The curated aliases the endpoint serves to the page.
const { SEED_ALIASES } = require('../lib/vendor-map');

// The spend sheet's real app names.
const APPS = ['Adobe', 'Anthropic(Api Console)', 'AWS', 'Bubble Starter', 'Claude Ai',
  'Claude Ai Max 6 Accounts', 'Cumul(Luzmo)', 'Cursor pro', 'Laptops Procurement',
  'Mango technology(Clickup)', 'OPENAI', 'windsurf pro'];
const sheetIndex = new Map(APPS.map(a => [invNorm(a), a]));

// The mapping the invoice import saves into the archive.
const MAPPING = {
  'Claude Api': 'Anthropic(Api Console)',
  'click up invoices': 'Mango technology(Clickup)',
  'Laptop procurment': 'Laptops Procurement',
};

const resolve = (folder, mapping, aliases) => invResolveFolder(folder, sheetIndex, mapping || {}, aliases || {});

test('an exact name needs no help', () => {
  assert.equal(resolve('Adobe'), 'Adobe');
  assert.equal(resolve('AWS'), 'AWS');
});

test('case and punctuation differences still match', () => {
  // The archive folder is "Open AI"; the sheet row is "OPENAI".
  assert.equal(resolve('Open AI'), 'OPENAI');
});

test('a vendor folder resolves to the app row that contains it', () => {
  assert.equal(resolve('Bubble'), 'Bubble Starter');
  assert.equal(resolve('Cursor'), 'Cursor pro');
  assert.equal(resolve('Luzmo'), 'Cumul(Luzmo)');
  assert.equal(resolve('windsurf'), 'windsurf pro');
});

test('an exact match wins over a longer row that also contains it', () => {
  // "Claude Ai" is its own row AND a prefix of "Claude Ai Max 6 Accounts".
  // Containment alone would be ambiguous; the exact test has to come first, or
  // Claude seat invoices land on the wrong row.
  assert.equal(resolve('Claude Ai'), 'Claude Ai');
});

test('an ambiguous folder is left unresolved rather than guessed', () => {
  const both = new Map([['acme', 'Acme'], ['acmecloud', 'Acme Cloud']]);
  // Exact still wins outright, punctuation and all.
  assert.equal(invResolveFolder('Acme Cloud', both, {}), 'Acme Cloud');
  assert.equal(invResolveFolder('acme', both, {}), 'Acme');
  // "acmeclo" contains "acme" and is contained by "acmecloud" — two defensible
  // answers, so it gets neither. Picking one would file invoices against a row
  // nobody chose, and a wrong tick is worse than a visible gap.
  assert.equal(invResolveFolder('acmeclo', both, {}), null);
});

test('the curated vendor aliases place folders the sheet names differently', () => {
  // No saved mapping needed: this is the same knowledge the statement importer
  // uses, so the checklist places "Claude Api" the moment the archive is read.
  assert.equal(resolve('Claude Api'), null, 'nothing about the name resembles the row');
  assert.equal(resolve('Claude Api', null, SEED_ALIASES), 'Anthropic(Api Console)');
  assert.equal(resolve('Laptop procurment', null, SEED_ALIASES), 'Laptops Procurement');
});

test('a trailing filing word is stripped before the alias lookup', () => {
  // The folder is "click up invoices"; the alias key is "clickup".
  assert.equal(resolve('click up invoices'), null);
  assert.equal(resolve('click up invoices', null, SEED_ALIASES), 'Mango technology(Clickup)');
});

test('an exact row name still wins over an alias pointing elsewhere', () => {
  // "Claude Ai" is its own row. An alias must never pull it onto another.
  assert.equal(resolve('Claude Ai', null, { claudeai: 'Anthropic(Api Console)' }), 'Claude Ai');
});

test('a folder that is only a filing word resolves to nothing', () => {
  // Stripping "bills" off "Courier bills" leaves "courier", which is no app.
  assert.equal(resolve('Courier bills', null, SEED_ALIASES), null);
  assert.equal(resolve('invoices', null, SEED_ALIASES), null);
});

test('the saved mapping settles what name-matching cannot', () => {
  // Nothing about "Claude Api" resembles "Anthropic(Api Console)".
  assert.equal(resolve('Claude Api'), null);
  assert.equal(resolve('Claude Api', MAPPING), 'Anthropic(Api Console)');
  assert.equal(resolve('click up invoices', MAPPING), 'Mango technology(Clickup)');
  // "Laptop procurment" vs "Laptops Procurement" — a typo and a plural apart.
  assert.equal(resolve('Laptop procurment'), null);
  assert.equal(resolve('Laptop procurment', MAPPING), 'Laptops Procurement');
});

test('the mapping wins over name-matching, but only if it names a real row', () => {
  assert.equal(resolve('Bubble', { Bubble: 'Claude Ai' }), 'Claude Ai');
  // A mapping pointing at a row that no longer exists falls back rather than
  // dropping the folder on the floor.
  assert.equal(resolve('Bubble', { Bubble: 'Deleted App' }), 'Bubble Starter');
});

test('a folder that is not an app at all stays unresolved', () => {
  // Real folders in the archive that are not subscriptions.
  for (const f of ['Courier bills', 'Laptop Repair', 'Laptops sold', 'Quotations']) {
    assert.equal(resolve(f, MAPPING), null, `"${f}" should not be forced onto a row`);
  }
});

test('very short folder names do not match by containment', () => {
  // Three letters or fewer inside a longer app name is coincidence, not a match.
  const idx = new Map([['adobecreativecloud', 'Adobe Creative Cloud']]);
  assert.equal(invResolveFolder('ado', idx, {}), null);
  assert.equal(invResolveFolder('adobe', idx, {}), 'Adobe Creative Cloud');
});

// --- A row that exists but has not been charged ---------------------------
//
// Posthog sits in the spend sheet with four invoices on file and no amounts
// entered yet. The pivot only carries rows with money in them, so the checklist
// joined against the pivot alone and reported the row as "not in sheet" while
// the sheet was plainly holding it.

test('the sheet index is built from every row, not only the charged ones', () => {
  // The pivot's rows plus the sheet's own list, which includes uncharged rows.
  assert.match(html, /window\.SHEET_APPS = json\.apps \|\| \[\]/,
    'the client must keep the full list of sheet rows');
  assert.match(html, /for \(const name of \(window\.SHEET_APPS \|\| \[\]\)\)/,
    'and resolve folders against it, not just the pivot');
});

test('an uncharged row is labelled as such, never as missing from the sheet', () => {
  assert.match(html, /inSheetUncharged: !!appName/,
    'a folder that resolved to a real row must be marked as one');
  assert.match(html, /no spend recorded/,
    'and shown with its own label rather than "not in sheet"');
});

test('the API hands over every app row it read', () => {
  const api = fs.readFileSync(path.join(__dirname, '..', 'api', 'spend-data.js'), 'utf8');
  assert.match(api, /if \(!apps\.includes\(name\)\) apps\.push\(name\)/,
    'every row is collected before the amounts are looked at');
  assert.match(api, /return \{ records, apps, appRows \}/);
  assert.match(api, /rowCount: rows\.length, rows, apps, appRows/, 'and all of them reach the payload');
});

test('an app with no spend is still listed, with its own details', () => {
  // Seven of the sheet's seventy-one rows have no figure in any month — Zapier,
  // Posthog, Sprinto and four more. A record is produced per month with money in
  // it, so those produced none and the table silently showed 64 apps.
  const api = fs.readFileSync(path.join(__dirname, '..', 'api', 'spend-data.js'), 'utf8');
  assert.match(api, /appRows\.push\(\{ name, dept, poc, renewalDate, cycle, cur, paymentMethod, kind \}\)/,
    'the row carries enough to render it without any amounts');
  // Collected before the month loop, so having no amounts cannot skip it.
  assert.ok(api.indexOf('appRows.push(') < api.indexOf('for (const { idx, month } of monthCols)'),
    'the row is recorded before the amounts are looked at, not inside that loop');

  assert.match(html, /window\.SHEET_ROWS = json\.appRows \|\| \[\]/,
    'the page keeps them');
  assert.match(html, /\(window\.SHEET_ROWS \|\| \[\]\)\.forEach/,
    'and seeds the pivot from them, so a row with no spend still appears');
  // They must obey the same filters, or a department or a search would show
  // apps it was not asked for.
  // Search forward from the seed block: `rows.forEach(r=>{` also appears earlier
  // in the file, and slicing to that one gives an empty range that asserts nothing.
  const seedStart = html.indexOf('(window.SHEET_ROWS || []).forEach');
  assert.ok(seedStart > -1, 'the pivot no longer seeds from the sheet rows — this test is stale');
  const seed = html.slice(seedStart, html.indexOf('rows.forEach(r=>{', seedStart));
  assert.ok(seed.length > 0 && seed.length < 2000, 'the seed block was not isolated');
  for (const guard of ['a.kind !== t', 'a.dept !== d', 'a.poc']) {
    assert.ok(seed.includes(guard), `the seeded rows must honour the ${guard} filter`);
  }
});


// --- The renewal-date parser reads every form the sheet uses ----------------
//
// Sprinto's renewal is written "19-09-2026" and never showed: the parser knew
// only US slashes and "Nth of every month". Pulled out of the page source and
// run as-is, so the page and this test cannot drift apart.
test('the page reads renewal dates in every form the sheet uses', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const start = html.indexOf('const MONTH_IDX =');
  const end = html.indexOf("let renewalsTab = 'yearly';");
  assert.ok(start > 0 && end > start, 'parseRenewal must sit where the test expects');
  const parseRenewal = new Function(html.slice(start, end) + '; return parseRenewal;')();
  const ymd = r => r && `${r.date.getFullYear()}-${String(r.date.getMonth()+1).padStart(2,'0')}-${String(r.date.getDate()).padStart(2,'0')}`;
  assert.strictEqual(ymd(parseRenewal('19-09-2026')), '2026-09-19');
  assert.strictEqual(ymd(parseRenewal('1/21/2027')), '2027-01-21');
  assert.strictEqual(ymd(parseRenewal('19/09/2026')), '2026-09-19', 'a slash still reads day-first when the first number cannot be a month');
  assert.strictEqual(ymd(parseRenewal('2026-09-19')), '2026-09-19');
  assert.strictEqual(ymd(parseRenewal('19 Sep 2026')), '2026-09-19');
  assert.strictEqual(ymd(parseRenewal('19-Sep-26')), '2026-09-19');
  assert.strictEqual(ymd(parseRenewal('Sep 19, 2026')), '2026-09-19');
  assert.strictEqual(ymd(parseRenewal('46284')), '2026-09-19', 'an Excel serial number');
  assert.strictEqual(parseRenewal('1st of every month').recurring, true);
  assert.strictEqual(parseRenewal('monthly').recurring, true);
  assert.strictEqual(parseRenewal(''), null);
  assert.strictEqual(parseRenewal('cancelled after june'), null);
  assert.strictEqual(parseRenewal('31-02-2026'), null, 'an impossible date is not a date');
});
