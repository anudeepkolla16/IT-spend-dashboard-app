// Run with: node --test
//
// The 16 September review found three ways the dashboard's numbers could not be
// trusted, and all three were real:
//   A  "Annual cost" $0 and "Total (all-time)" $120 for the same app, because
//      the sheet keeps actual charges and budgeted months in the same columns.
//   B  An Applications search matching nothing turned every recorded charge into
//      an apparent missing charge in Invoices: coverage read "0 of 0 · 100%".
//   C  A part-billed September was reported as a 62% saving against a whole
//      August, in green, on the 16th.
// These pin the fixes against the page source, which is where they live.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
function lift(names, from, to, extra) {
  const a = html.indexOf(from), b = html.indexOf(to, a);
  assert.ok(a > -1 && b > a, `${from} must sit where the test expects`);
  return new Function(`${extra || ''}\n${html.slice(a, b)}\n; return {${names.join(',')}};`)();
}

const TODAY = '2026-09';
const M = lift(
  ['spendIn', 'actualSpend', 'closedSpend', 'mtdSpend', 'plannedSpend', 'billedRows'],
  '/* ---------------- what "spend" means',
  'function latestRowsPerApp(',
  `const realMonthKey = () => '${TODAY}';`
);

const row = (name, month, usd, extra) => ({ name, month, usd, kind: 'Apps', dept: 'IT', cat: 'Monthly', poc: '', renewalDate: '', ...(extra || {}) });
// One app charged through August, one whose only figure is a budgeted future month.
const ROWS = [
  row('AWS', '2026-07', 2260.67), row('AWS', '2026-08', 2374.66), row('AWS', '2026-09', 900),
  row('CANVA', '2026-11', 120, { cat: 'Annual' }),
  { ...row('Laptops', '2026-06', 3493.17), kind: 'Laptops' },
  row('Nothing', null, 0),
];

// --- A. actual, planned and month-to-date are separate measures --------------

test('the four measures split actual charges from budgeted months', () => {
  assert.strictEqual(M.actualSpend(ROWS), 2260.67 + 2374.66 + 900 + 3493.17, 'through this month, this month included');
  assert.strictEqual(M.closedSpend(ROWS), 2260.67 + 2374.66 + 3493.17, 'months that have finished');
  assert.strictEqual(M.mtdSpend(ROWS), 900);
  assert.strictEqual(M.plannedSpend(ROWS), 120, 'a month still ahead is a budget, never spend');
  assert.strictEqual(M.actualSpend(ROWS) + M.plannedSpend(ROWS), ROWS.reduce((s, r) => s + (r.month ? r.usd : 0), 0),
    'actual and planned account for every dated figure exactly once');
  assert.strictEqual(M.closedSpend(ROWS) + M.mtdSpend(ROWS), M.actualSpend(ROWS));
});

test('billedRows drops budgeted months and rows with no month at all', () => {
  assert.deepStrictEqual(M.billedRows(ROWS).map(r => `${r.name} ${r.month}`),
    ['AWS 2026-07', 'AWS 2026-08', 'AWS 2026-09', 'Laptops 2026-06']);
  assert.strictEqual(M.billedRows(ROWS).length, ROWS.length - 2);
  assert.strictEqual(M.actualSpend([]), 0);
  assert.strictEqual(M.actualSpend(null), 0);
});

test('the app card and the detail panel report the same pair of measures', () => {
  // Canva: nothing charged, $120 budgeted ahead. The table's 12-month column and
  // the modal must not disagree about that.
  assert.match(html, /Spend · last 12 mo/, 'the column names its measure');
  assert.match(html, /Actual charges in the twelve months ending with the current one/);
  assert.match(html, /<span class="cap">Actual spend · to date<\/span>/, 'the modal reports actual');
  assert.match(html, /<span class="cap">Budgeted ahead<\/span>/, 'and planned, apart');
  assert.ok(!/cap">Total \(all-time\)/.test(html), 'the label that mixed them is no longer rendered');
  assert.match(html, /if \(m <= modalToday\) modalActual \+= v; else modalPlanned \+= v;/);
  // The KPI strip and the donut say which basis they are on.
  assert.match(html, /Actual spend · to date<\/div><div class="val">\$\{fmtUSD\(toDate\)\}/);
  assert.match(html, /budgeted ahead, not yet billed/);
  assert.match(html, /Last 12<br>months by<br>department/);
  // Charts of "spend" count charges only.
  assert.match(html, /const g=\{\}; billedRows\(apps\)\.forEach/);
  assert.match(html, /groupSum\(billedRows\(all\),'cat'\)/);
});

// --- B. reconciliation is a property of the data, not of a filter ------------

const P = lift(['buildPivot'], 'function buildPivot(all, filt){', '\nfunction draw()', `
  const ALL_MONTHS = ['2026-07','2026-08','2026-09','2026-11'];
  const window = { SHEET_ROWS: [
    { name:'AWS', kind:'Apps', dept:'Engineering', cycle:'Monthly', cur:'USD', poc:'Ajay', renewalDate:'' },
    { name:'CANVA', kind:'Apps', dept:'Marketing', cycle:'Annual', cur:'USD', poc:'Bhavana', renewalDate:'' },
    { name:'Laptops', kind:'Laptops', dept:'IT', cycle:'Monthly', cur:'USD', poc:'', renewalDate:'' },
  ] };`);

test('buildPivot with no filter is the whole sheet; a filter only narrows the copy', () => {
  const all = P.buildPivot(ROWS, {});
  assert.deepStrictEqual(all.map(g => g.name).sort(), ['AWS', 'CANVA', 'Laptops', 'Nothing']);
  assert.strictEqual(all.find(g => g.name === 'AWS')['2026-08'], 2374.66);
  assert.strictEqual(all.find(g => g.name === 'CANVA')['2026-11'], 120);

  assert.deepStrictEqual(P.buildPivot(ROWS, { d: 'Marketing' }).map(g => g.name), ['CANVA']);
  assert.deepStrictEqual(P.buildPivot(ROWS, { t: 'Laptops' }).map(g => g.name), ['Laptops']);
  assert.deepStrictEqual(P.buildPivot(ROWS, { q: 'aws' }).map(g => g.name), ['AWS']);
  // The case that broke reconciliation: a search matching nothing.
  assert.deepStrictEqual(P.buildPivot(ROWS, { q: 'zzzz' }), []);
  // …and the unfiltered pivot is untouched by it.
  assert.strictEqual(P.buildPivot(ROWS, {}).length, 4);
});

test('the checklist and the drill-down read the unfiltered pivot', () => {
  assert.match(html, /function invBuildRows\(\)\{\s*\/\/[^\n]*\n\s*const pivoted = window\.PIVOT_ALL \|\| window\.PIVOTED \|\| \[\];/);
  assert.match(html, /const g = \(window\.PIVOT_ALL\|\|window\.PIVOTED\|\|\[\]\)\.find\(x=>x\.name===name\)/,
    'a renewal or invoice cell still opens its app while an unrelated search is active');
  assert.match(html, /window\.PIVOT_ALL = buildPivot\(all, \{\}\);/);
  assert.match(html, /let pivoted = \(t \|\| d \|\| q\) \? buildPivot\(all, \{ t, d, q \}\) : window\.PIVOT_ALL\.slice\(\);/);
});

test('no charged month reports no coverage, never 100%', () => {
  assert.match(html, /const pct = charged \? Math\.round\(covered\/charged\*100\) : null;/);
  assert.match(html, /'No charged months to reconcile'/);
  assert.ok(!/: 100;/.test(html), 'the empty denominator no longer reads as full coverage');
});

// --- C. an incomplete period is never reported as a saving -------------------

test('the current month is labelled month-to-date and compared against nothing', () => {
  assert.match(html, /Spend · \$\{monthLabel\(thisKey\)\} <span class="mtd">month to date<\/span>/);
  assert.match(html, /const thisMonth = mtdSpend\(all, thisKey\);/);
  // The trend it does report is the last two months that finished.
  assert.match(html, /const lastClosed = totals\[prevKey\] \|\| 0, prevClosed = totals\[prev2Key\] \|\| 0;/);
  assert.match(html, /closed at \$\{fmtUSD\(lastClosed\)\}/);
  assert.ok(!/mom>=0\?'▲':'▼'/.test(html), 'the month-to-date arrow is gone');
});

test('the spend chart is solid to the last closed month and dashed after it', () => {
  assert.match(html, /const closedEnd = months\.reduce\(\(n, m\) => \(m < todayKey \? n \+ 1 : n\), 0\);/);
  assert.match(html, /segment: \{ borderDash: \(ctx\) => \(ctx\.p0DataIndex >= closedEnd - 1 \? \[5, 4\] : undefined\) \}/);
  assert.match(html, /'budgeted, not yet billed'/);
  assert.match(html, /the last month that finished; dashed after it/);
  // The Forecast tab's own actual series is indexed over its own months.
  assert.match(html, /line\('Apps spend \(actual\)'[^\n]*, true\)/);
});

// --- D. status, charges and forecast eligibility are three different things --
//
// The review found windsurf and Hex reported Active with a monthly run-rate
// while the sheet's renewal column read "cancled after june"; Zapier reported
// "not charged lately" with a charge this month; and the headline counted 65
// active subscriptions while the run-rate under it was built from 58 and the
// forecast from 42.

const S = lift(['isCancelled', 'trailingRunRateByApp', 'monthlyRunRate'],
  '// The sheet records a cancellation in the renewal column',
  '\nfunction renderKPIs(',
  `const realMonthKey = () => '${TODAY}';
   function addMonths(key, n){ const [y,m] = key.split('-').map(Number); const d = new Date(y, m-1+n, 1); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
   function monthsBetween(a,b){ const [fy,fm]=a.split('-').map(Number), [ty,tm]=b.split('-').map(Number); return (ty-fy)*12+(tm-fm); }
   function cycleMonths(c){ c=String(c||'Monthly').toLowerCase(); if(c.includes('one'))return 0; const n=c.match(/(\\d+)\\s*year/); if(n)return +n[1]*12; if(c.includes('half'))return 6; if(c.includes('quarter'))return 3; if(c.includes('year')||c.includes('annual'))return 12; return 1; }`);

test('the sheet\'s own words for a cancellation are read', () => {
  for (const text of ['cancled after june', 'Cancled after june', 'cancelled', 'Canceled 1 Aug', 'terminated', 'ended in June']) {
    assert.strictEqual(S.isCancelled(text), true, `"${text}" is a cancellation`);
  }
  for (const text of ['', null, '1st of every month', '9/19/2027', '2026-09-19', 'monthly']) {
    assert.strictEqual(S.isCancelled(text), false, `"${text}" is not`);
  }
});

const app = (name, month, usd, cycle, renewalDate) => ({ name, month, usd, cycle: cycle || 'Monthly', dept: 'IT', renewalDate: renewalDate || '' });

test('a cancelled subscription carries no run-rate, whatever its charges say', () => {
  const rows = [
    app('windsurf', '2026-06', 255, 'Monthly', 'Cancled after june'),
    app('windsurf', '2026-07', 255, 'Monthly', 'Cancled after june'),
    app('windsurf', '2026-08', 255, 'Monthly', 'Cancled after june'),
    app('Hex', '2026-07', 159.38, 'Monthly', 'cancled after june'),
    app('Hex', '2026-08', 159.38, 'Monthly', 'cancled after june'),
    app('Adobe', '2026-07', 37.16), app('Adobe', '2026-08', 37.16),
  ];
  const by = Object.fromEntries(S.trailingRunRateByApp(rows, TODAY).map(x => [x.name, x]));
  assert.ok(!by['windsurf'], 'three months of charges do not make a cancelled app active');
  assert.ok(!by['Hex']);
  assert.ok(by['Adobe'], 'an app with no cancellation note is unaffected');
  // Adobe started in July, so it averages over the two months it has run.
  assert.strictEqual(S.monthlyRunRate(rows), 37.16, 'the run-rate is only the live ones');
});

test('an app first charged this month is active, not lapsed', () => {
  // Zapier: nothing in the three complete months because it had not started.
  const zapier = S.trailingRunRateByApp([app('Zapier', '2026-09', 31.86)], TODAY);
  assert.strictEqual(zapier.length, 1);
  assert.strictEqual(zapier[0].m, 31.86);
  assert.strictEqual(zapier[0].newThisMonth, true);
  assert.deepStrictEqual(zapier[0].basis, [TODAY]);
  // An annual bill landing this month is still spread over its cycle.
  const sprinto = S.trailingRunRateByApp([app('Sprinto', '2026-09', 25000, 'Annual')], TODAY);
  assert.strictEqual(sprinto[0].pattern, 'cycle');
  assert.strictEqual(Math.round(sprinto[0].m), Math.round(25000 / 12));
  assert.ok(!sprinto[0].newThisMonth);
  // An app genuinely not charged for three complete months is still lapsed.
  assert.deepStrictEqual(S.trailingRunRateByApp([app('Gone', '2026-05', 31.86)], TODAY), []);
});

test('active, run-rate and forecast describe one population', () => {
  assert.match(html, /const rates = trailingRunRateByApp\(apps\);/);
  assert.match(html, /const rr = rates\.reduce\(\(s, x\) => s \+ x\.m, 0\);/);
  assert.match(html, /const activeApps = rates\.length;/, 'active is exactly the set carrying a run-rate');
  assert.match(html, /the same set the run-rate and forecast use/);
  assert.match(html, /cancelled in the sheet/);
  assert.match(html, /\['ended', 'Cancelled'\]/, 'the table separates contract status from billing status');
  assert.match(html, /first charged this month/);
});

// --- F. a renewal total says how much of itself is known ---------------------

test('the renewal total says how many renewals still need pricing', () => {
  assert.match(html, /const unpriced = items\.filter\(x => !x\.cost\)\.length;/);
  assert.match(html, /known value\$\{unpriced \? `, \$\{unpriced\} still to price` : ''\}/);
  assert.match(html, /with no charge on record to price/);
  assert.ok(!/\$\{fmtUSD\(total\)\} expected/.test(html), 'an unqualified "expected" total is gone');
});

// --- E. the sidebar is hidden below 1000px; something has to replace it ------

test('a compact nav is cloned from the sidebar, so the two cannot list different places', () => {
  assert.match(html, /<nav class="mobnav" id="mobNav" aria-label="Sections"><\/nav>/);
  assert.match(html, /document\.querySelectorAll\('\.side \.nav-item'\)\.forEach\(src => \{/,
    'built from the sidebar itself, never a second hand-written list');
  assert.match(html, /b\.onclick = \(\) => src\.click\(\);/, 'a clone delegates to its sidebar button, so both route the same way');
  assert.match(html, /@media\(max-width:1000px\)\{[\s\S]*?\.mobnav\{display:flex/);
  // The pending count is mirrored, not recomputed.
  assert.match(html, /mirror\.textContent = badge\.textContent/);
  assert.match(html, /badge\.classList\.toggle\('zero', !items\.length\); syncMobNav\(\);/);
});

test('nothing but the page itself is allowed to be wider than the window', () => {
  // 390px of window against 726px of document, after a resize: a canvas kept
  // its old width because min-width:auto let it hold the card open.
  assert.match(html, /\.card,\.grid>\*,\.kstrip>\*\{min-width:0\}/);
  assert.match(html, /canvas\{max-width:100%\}/);
  assert.match(html, /#invChecklist,#appsOverview,\.loginwrap\{overflow-x:auto\}/);
  assert.match(html, /charts\[id\]\.resize\(\)/, 'charts follow the window down as well as up');
  // One column of KPI tiles on a phone, not two-plus-one.
  assert.match(html, /@media\(max-width:560px\)\{\s*\.kstrip\{grid-template-columns:1fr\}/);
  // The archive-maintenance actions fold away on the widths that need the room.
  assert.match(html, /\.maint-toggle\{display:inline-flex/);
  assert.match(html, /\.toolbar\{order:4;flex-basis:100%;display:none\}/);
});

// --- The decision layer: what changed, and what a click means ---------------
//
// "The current landing page answers 'What is recorded?' better than 'What
// should we do?'" — with no approved budget to compare against, what the sheet
// can still answer is which apps moved the total, and which gaps are worth
// chasing.

const C = lift(['changeDrivers'], '// Which apps moved the total between the two months', '\nfunction renderChange(');

test('what changed ranks apps by the money that moved, not by percentage', () => {
  const pivot = [
    { name: 'Google cloud', dept: 'Engineering', '2026-07': 41225.25, '2026-08': 34633.23 },
    { name: 'Anthropic', dept: 'org', '2026-07': 13895.83, '2026-08': 16376.29 },
    { name: 'Vercel', dept: 'Engineering', '2026-08': 21.25 },                 // new: from nothing
    { name: 'Hex', dept: 'Consulting', '2026-07': 159.38 },                    // gone: to nothing
    { name: 'Adobe', dept: 'Marketing', '2026-07': 37.16, '2026-08': 37.16 },  // flat
    { name: 'Rounding', dept: 'IT', '2026-07': 10, '2026-08': 10.4 },          // below a dollar
  ];
  const { rows, total } = C.changeDrivers(pivot, '2026-07', '2026-08');
  assert.deepStrictEqual(rows.map(r => r.name), ['Google cloud', 'Anthropic', 'Hex', 'Vercel'],
    'biggest dollar move first — a 10,278% rise on $12 is not the story');
  assert.strictEqual(Math.round(rows[0].delta), -6592);
  assert.strictEqual(Math.round(rows[1].delta), 2480);
  assert.strictEqual(rows.find(r => r.name === 'Vercel').from, 0, 'an app that started shows as from nothing');
  assert.strictEqual(rows.find(r => r.name === 'Hex').to, 0);
  assert.ok(!rows.some(r => r.name === 'Adobe'), 'a flat app is not a driver');
  assert.ok(!rows.some(r => r.name === 'Rounding'), 'nor is a move under a dollar');
  assert.strictEqual(Math.round(total), Math.round(rows.reduce((s, r) => s + r.delta, 0)));
  assert.deepStrictEqual(C.changeDrivers([], '2026-07', '2026-08'), { rows: [], total: 0 });
  assert.deepStrictEqual(C.changeDrivers(null, '2026-07', '2026-08').rows, []);
});

test('the card compares the two months that have finished, and is hidden with nothing to say', () => {
  assert.match(html, /const to = realMonthKey\(-1\), from = realMonthKey\(-2\);/);
  assert.match(html, /if \(!rows\.length\)\{ card\.classList\.add\('hidden'\); return; \}/);
  assert.match(html, /the two months that have finished/);
  assert.match(html, /smaller move\$\{rest\.length===1\?'':'s'\}/, 'the tail is accounted for, not dropped');
});

test('a drill-down keeps the month it was opened from', () => {
  assert.match(html, /function openModal\(name, month\)\{/);
  assert.match(html, /modalMonth = month \|\| null;/);
  assert.match(html, /\$\('#modalTitle'\)\.textContent = g\.name \+ \(modalMonth \? ` · \$\{monthLabel\(modalMonth\)\}` : ''\)/);
  assert.match(html, /recorded charge \$\{charge==null\?'none':fmtUSD\(charge\)\} · invoice status/);
  // Every route in carries it: a checklist cell, a renewal, a mover.
  assert.match(html, /const cell = e\.target\.closest\('td\.mark'\);\s*openModal\(decodeURIComponent\(tr\.dataset\.name\), cell \? cell\.dataset\.month : null\);/);
  assert.match(html, /openModal\(decodeURIComponent\(row\.dataset\.name\), row\.dataset\.month\|\|null\)/);
  assert.match(html, /openModal\(decodeURIComponent\(r\.dataset\.name\), r\.dataset\.month\)/);
  assert.match(html, /modalAppName = null;\s*modalMonth = null;/, 'and closing clears it');
});

test('a missing invoice is reported with the charge behind it', () => {
  assert.match(html, /missingValue \+= \(row\.sheet && row\.sheet\[m\]\) \|\| 0;/);
  assert.match(html, /\$\{fmtUSD\(missingValue\)\} unevidenced/);
});

test('every chart carries a label a screen reader can read', () => {
  for (const id of ['trendChart', 'deptChart', 'catChart', 'topChart', 'runrateChart', 'modalChart']) {
    assert.match(html, new RegExp(`<canvas id="${id}" role="img" aria-label="[^"]+"`), `${id} is labelled`);
  }
  // The trend chart's label follows its tab and carries the figures.
  assert.match(html, /\$\('#trendChart'\)\.setAttribute\('aria-label',/);
});

// --- Real routes -------------------------------------------------------------
//
// "The sidebar looks like page navigation, but it jumps between sections of one
// large document… The URL remained unchanged during navigation, limiting
// bookmarking and sharing."

const ROUTER = (() => {
  const a = html.indexOf('const PAGES = {'), b = html.indexOf('function routeTo(');
  return new Function(`${html.slice(a, b)}; return { PAGES, DEFAULT_PAGE, parseRoute, routeQuery };`)();
})();

test('every sidebar destination is a route, and an unknown one falls back', () => {
  assert.deepStrictEqual(Object.keys(ROUTER.PAGES),
    ['overview', 'applications', 'renewals', 'invoices', 'questions', 'rules', 'passwords']);
  // The sidebar lists exactly those, in that order.
  const nav = [...html.matchAll(/<button class="nav-item[^"]*" data-page="([a-z]+)"/g)].map(m => m[1]);
  assert.deepStrictEqual(nav, Object.keys(ROUTER.PAGES), 'the nav and the router cannot list different places');

  assert.strictEqual(ROUTER.parseRoute('#/invoices').page, 'invoices');
  assert.strictEqual(ROUTER.parseRoute('#/nonsense').page, ROUTER.DEFAULT_PAGE);
  assert.strictEqual(ROUTER.parseRoute('').page, ROUTER.DEFAULT_PAGE);
  assert.strictEqual(ROUTER.parseRoute(null).page, ROUTER.DEFAULT_PAGE);
  assert.strictEqual(ROUTER.parseRoute('#/applications').page, 'applications');
  // A route carries the Applications filters.
  const { page, params } = ROUTER.parseRoute('#/applications?q=aws&dept=Marketing&type=Apps');
  assert.strictEqual(page, 'applications');
  assert.deepStrictEqual([params.get('q'), params.get('dept'), params.get('type')], ['aws', 'Marketing', 'Apps']);
  // Every page has a title and a one-line description.
  for (const [name, p] of Object.entries(ROUTER.PAGES)) {
    assert.ok(p.title && p.sub, `${name} names itself`);
  }
});

test('each card declares its page, and the router only hides content', () => {
  for (const [id, page] of [['kpis', 'overview'], ['trendCard', 'overview'], ['changeCard', 'overview'],
                            ['appsCard', 'applications'], ['renewalsCard', 'renewals'],
                            ['invoiceCard', 'invoices'], ['loginsCard', 'passwords']]) {
    assert.match(html, new RegExp(`id="${id}"[^>]*data-page="[^"]*${page}`), `${id} belongs to ${page}`);
  }
  assert.match(html, /id="pendingCard" data-page="questions rules"/, 'one card can serve two routes');
  // Scoped to main: the nav buttons carry data-page too, and hiding those hid
  // the navigation itself.
  assert.match(html, /document\.querySelectorAll\('main \[data-page\]'\)/);
  assert.ok(!/document\.querySelectorAll\('\[data-page\]'\)\.forEach\(el =>/.test(html));
});

test('navigating writes history, and filters stay in the URL', () => {
  assert.match(html, /if \(replace\) history\.replaceState\(null, '', hash\); else location\.hash = hash;/);
  assert.match(html, /window\.addEventListener\('hashchange'/);
  assert.match(html, /if \(routeApplied && \[\.\.\.params\.keys\(\)\]\.length\)\{ applyRouteParams\(params\); draw\(\); \}/,
    'a pasted link sets filters; moving between pages leaves them alone');
  assert.match(html, /if \(location\.hash !== hash\) history\.replaceState\(null, '', hash\);/,
    'typing in a filter must not push a history entry per keystroke');
  assert.match(html, /document\.title = `\$\{PAGES\[name\]\.title\} · Saras IT Spend`/);
  // The initial route runs last: showPage reads `let` state declared further
  // down the script, and calling it earlier died in the temporal dead zone,
  // taking the whole dashboard with it.
  const call = html.lastIndexOf('showPage(parseRoute(location.hash).page);');
  assert.ok(call > html.indexOf('let loginRows'), 'the startup call sits after the state it reads');
  assert.ok(call > html.indexOf('let charts'));
});
