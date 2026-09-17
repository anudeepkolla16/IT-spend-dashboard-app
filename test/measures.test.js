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

// --- Vendor consolidation and the cloud/AI cut -------------------------------
//
// "Consolidate applications under parent vendors — especially where multiple
// products or accounts belong to Google, Anthropic, or another supplier."
// Cloud and AI is the same data cut differently; per-project and utilisation
// figures need the providers' exports, which the sheet does not hold.

const V = (() => {
  const a = html.indexOf('const VENDOR_FAMILIES = {'), b = html.indexOf('let vendorCloudOnly');
  return new Function(`
    const ALL_MONTHS = ['2025-09','2026-07','2026-08','2026-09'];
    function addMonths(key, n){ const [y,m] = key.split('-').map(Number); const d = new Date(y, m-1+n, 1); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
    ${html.slice(a, b)}
    ; return { vendorOf, isCloudAi, buildVendorRows, VENDOR_FAMILIES, CLOUD_AI };`)();
})();

test('several rows of one supplier read as one relationship', () => {
  assert.strictEqual(V.vendorOf('Google cloud'), 'Google');
  assert.strictEqual(V.vendorOf('GOOGLE ADS'), 'Google');
  assert.strictEqual(V.vendorOf('Google Voice'), 'Google');
  // "Claude Ai" does not resemble "Anthropic"; only a curated map joins them.
  assert.strictEqual(V.vendorOf('Claude Ai'), 'Anthropic');
  assert.strictEqual(V.vendorOf('Anthropic(Api Console)'), 'Anthropic');
  assert.strictEqual(V.vendorOf('SLACK'), 'Salesforce');
  // Spelling, case and punctuation in the sheet do not decide it.
  assert.strictEqual(V.vendorOf('google  cloud'), 'Google');
  assert.strictEqual(V.vendorOf('Render '), 'Render', 'an unlisted app is its own vendor, trimmed');
  assert.strictEqual(V.vendorOf(''), '');
});

test('the cloud and AI cut is a named list, not an inference', () => {
  assert.strictEqual(V.isCloudAi('Google cloud'), true);
  assert.strictEqual(V.isCloudAi('Anthropic(Api Console)'), true);
  assert.strictEqual(V.isCloudAi('Cursor pro'), true);
  assert.strictEqual(V.isCloudAi('GOOGLE ADS'), false, 'advertising is not cloud spend');
  assert.strictEqual(V.isCloudAi('Adobe'), false);
  // The view names every app it counted rather than asking anyone to trust it.
  assert.match(html, /Cloud and AI counts \$\{cloudRows\.length\}/);
  assert.match(html, /need the providers' own billing and usage exports, which the sheet does not hold/);
});

test('a vendor row totals its products, over the last 12 months only', () => {
  const pivot = [
    { name: 'Google cloud', kind: 'Apps', '2026-07': 41225.25, '2026-08': 34633.23, '2025-09': 9999 },
    { name: 'GOOGLE ADS', kind: 'Apps', '2026-08': 2196.54 },
    { name: 'Claude Ai', kind: 'Apps', '2026-08': 1000 },
    { name: 'Adobe', kind: 'Apps', '2026-08': 37.16 },
    { name: 'Laptops Procurement', kind: 'Laptops', '2026-08': 6143.59 },
  ];
  const rates = [{ name: 'Google cloud', m: 38182 }, { name: 'GOOGLE ADS', m: 1500 }, { name: 'Adobe', m: 37.16 }];
  const { rows, total } = V.buildVendorRows(pivot, rates, '2026-09');

  const google = rows.find(r => r.vendor === 'Google');
  assert.deepStrictEqual(google.apps, ['Google cloud', 'GOOGLE ADS']);
  assert.strictEqual(google.spend, 41225.25 + 34633.23 + 2196.54, 'the 2025 month is outside the window');
  assert.strictEqual(google.run, 38182 + 1500, 'and the run-rate is the one the rest of the dashboard uses');
  assert.strictEqual(google.cloudAi, true, 'one cloud product makes the relationship a cloud one');

  assert.ok(!rows.some(r => r.vendor === 'Laptops Procurement'), 'hardware is not a subscription vendor');
  assert.strictEqual(rows[0].vendor, 'Google', 'biggest relationship first');
  assert.strictEqual(Math.round(total), Math.round(rows.reduce((s, r) => s + r.spend, 0)));
  assert.ok(Math.abs(rows.reduce((s, r) => s + r.share, 0) - 1) < 1e-9, 'shares account for the whole estate');
  assert.deepStrictEqual(V.buildVendorRows([], [], '2026-09'), { rows: [], total: 0 });
});

test('the vendor view is a tab on Applications and redraws with the data', () => {
  assert.match(html, /<button data-tab="vendor">By vendor<\/button>/);
  assert.match(html, /\$\('#appsVendor'\)\.classList\.toggle\('hidden', tab !== 'vendor'\)/);
  assert.match(html, /if \(tab === 'vendor'\) drawVendors\(\);/);
  assert.match(html, /if \(appsTab === 'vendor'\) drawVendors\(\);/, 'and again when the sheet reloads');
});

// --- Editing the sheet from the dashboard ------------------------------------
//
// "can we also have a feature to edit in the dashboard directly only for my id"
//
// The controls are the owner's alone, and the page showing them is a courtesy,
// not the gate: api/amounts edit settles that from the session on every write
// (test/edit.test.js). What these pin is that the page never offers an edit it
// has not been told to offer, and never loses a figure while saving one.

const E = (() => {
  const a = html.indexOf('function editableField(app, field, label, value){');
  const b = html.indexOf("document.addEventListener('keydown'", a);
  return { src: html.slice(a, b), fn: (canEdit) => new Function(`
    const window = { CAN_EDIT: ${canEdit} };
    ${html.slice(a, b)}
    ; return editableField;`)() };
})();

test('a signed-in user who is not an editor is shown no way to edit', () => {
  // Month cells: the class, the data the handler reads, and the hint all hang
  // off CAN_EDIT, so a non-editor's table is exactly the table it always was.
  assert.match(html, /\$\{window\.CAN_EDIT\?' editable':''\}/);
  assert.match(html, /\$\{window\.CAN_EDIT\?` data-edit-app="\$\{encodeURIComponent\(g\.name\)\}" data-edit-month="\$\{m\}"`:''\}/);
  assert.match(html, /if \(!window\.CAN_EDIT \|\| td\.classList\.contains\('editing'\)\) return;/);
  assert.strictEqual(E.fn(false)('AWS', 'poc', 'POC', 'Ajay'), '<b>Ajay</b>');
  assert.strictEqual(E.fn(false)('AWS', 'poc', 'POC', ''), '<b>—</b>');
  // And the flag comes from the server's answer about this session, not from
  // anything the page decides for itself.
  assert.match(html, /window\.CAN_EDIT = !!me\.canEdit;/);
  assert.match(html, /document\.body\.classList\.toggle\('can-edit'/);
});

test('an editable field carries the app and value it will send', () => {
  const out = E.fn(true)('Google cloud', 'renewalDate', 'Renewal date', '3rd of every month');
  assert.match(out, /data-field="renewalDate"/);
  assert.match(out, /data-app="Google%20cloud"/, 'a name with a space survives the round trip');
  assert.match(out, /value="3rd of every month"/);
  // A quote in the value must not end the attribute and let the rest through.
  assert.match(E.fn(true)('X', 'poc', 'POC', 'a" onfocus="boom'), /value="a&quot; onfocus=&quot;boom"/);
  assert.ok(!/value="a" onfocus="boom"/.test(E.fn(true)('X', 'poc', 'POC', 'a" onfocus="boom')));
});

test('a month cell is an edit, and the rest of the row is still a drill-down', () => {
  // Both listeners see the same click; without this the modal opened on top of
  // the input the click had just created.
  assert.match(html, /if \(e\.target\.closest\('td\.editable, td\.editing'\)\) return;\s*\n\s*const tr = e\.target\.closest\('tr\[data-name\]'\); if \(tr\) openModal/);
});

test('a refused edit puts the figure back and says why', () => {
  const src = html.slice(html.indexOf('function beginCellEdit(td){'), html.indexOf('document.addEventListener(\'click\', (e) => {\n  const td'));
  // One attempt per commit: Enter commits and Escape restores, and both set
  // `done` so the blur that follows cannot send the same edit a second time.
  assert.match(src, /let done = false;/);
  assert.match(src, /const commit = async \(\) => \{\s*\n\s*if \(done\) return;/);
  assert.match(src, /if \(typed === started\)\{ restore\(original\); return; \}/, 'typing nothing new is not a write');
  assert.match(src, /catch\(err\)\{[\s\S]*td\.textContent = original;[\s\S]*Could not save/);
  assert.match(src, /syncFromApi\(true\)/, 'a saved cell is re-read from the sheet, not assumed');
  // Escape leaves the cell as it was found.
  assert.match(src, /else if \(e\.key === 'Escape'\)\{ e\.preventDefault\(\); restore\(original\); \}/);
});

test('payment method reaches the drill-down, the sheet column and all', () => {
  // The sheet records which card paid for a row, and it was the one detail the
  // drill-down could not show: the page dropped the column on the way in, so
  // there was nothing to render or edit. (Currency is not inferred from it —
  // every amount in this workbook is USD.)
  assert.match(html, /paymentMethod:r\.paymentMethod\|\|''/);
  assert.match(html, /pay:a\.paymentMethod\|\|''/);
  assert.match(html, /editableField\(g\.name, 'paymentMethod', 'Payment method', g\.pay\)/);
  const P2 = lift(['buildPivot'], 'function buildPivot(all, filt){', '\nfunction draw()', `
    const ALL_MONTHS = ['2026-07','2026-08','2026-09'];
    const window = {};
    window.SHEET_ROWS = [{ name:'Adobe', kind:'Apps', dept:'IT', cycle:'Monthly', cur:'USD', poc:'', renewalDate:'', paymentMethod:'HDFC card' }];`);
  const rows = P2.buildPivot([{ name:'AWS', month:'2026-08', usd:10, kind:'Apps', dept:'IT', cat:'Monthly', cur:'USD', poc:'', renewalDate:'', paymentMethod:'Amex' }], {});
  assert.strictEqual(rows.find(r => r.name === 'AWS').pay, 'Amex', 'from a charged row');
  assert.strictEqual(rows.find(r => r.name === 'Adobe').pay, 'HDFC card', 'and from a row with no charges yet');
});

// --- The monthly breakdown's Total row ---------------------------------------
//
// "can we also add monthly total here" — 73 rows of a column nobody can add up
// in their head.

const MT = (() => {
  const a = html.indexOf('function breakdownTotals(rows){'), b = html.indexOf('\nfunction draw()', a);
  return new Function(`
    const ALL_MONTHS = ['2026-07','2026-08','2026-09'];
    ${html.slice(a, b)}
    ; return breakdownTotals;`)();
})();

test('a month column adds up only the apps actually charged in it', () => {
  const g = (name, months, total) => ({ name, ...months, total });
  const t = MT([
    g('AWS',   { '2026-07': 2260.67, '2026-08': 2374.66, '2026-09': 900 }, 5535.33),
    g('Adobe', { '2026-08': 37.16 }, 37.16),
    g('Never', {}, 0),
  ]);
  assert.strictEqual(t.byMonth['2026-07'], 2260.67);
  assert.strictEqual(Math.round(t.byMonth['2026-08'] * 100) / 100, 2411.82);
  // A month with no figure is not a zero: it counts nobody, so the row can say
  // "—" instead of "$0", which would read as "we were charged nothing".
  assert.strictEqual(t.countByMonth['2026-07'], 1);
  assert.strictEqual(t.countByMonth['2026-08'], 2);
  assert.strictEqual(t.countByMonth['2026-09'], 1);
  assert.strictEqual(MT([]).grand, 0);
  assert.deepStrictEqual(MT([]).countByMonth, { '2026-07': 0, '2026-08': 0, '2026-09': 0 });
});

test('the grand total is the Total column added up, not the months', () => {
  // They are the same whenever every charge carries a month. When one does not,
  // the figure under the Total column still has to equal that column.
  const t = MT([
    { name: 'AWS', '2026-07': 100, total: 100 },
    { name: 'Odd', total: 40 },                    // a charge with no month
  ]);
  assert.strictEqual(t.grand, 140);
  assert.strictEqual(t.byMonth['2026-07'], 100);
});

test('the Total row follows the filters and flags the part-billed month', () => {
  assert.match(html, /const foot = breakdownTotals\(pivoted\);/, 'the rows on screen, not the whole sheet');
  // A second monthTotals() would have quietly replaced the one the trend chart
  // calls, and the charts are stubbed in the browser harness, so nothing would
  // have looked wrong until the chart was read.
  assert.strictEqual((html.match(/function monthTotals\(/g) || []).length, 1, 'one function, one name');
  assert.match(html, /\$\('#tbl tfoot'\)\.innerHTML = pivoted\.length \?/, 'nothing to total when nothing matches');
  assert.match(html, /const partial = m === realMonthKey\(\);/);
  assert.match(html, /still being billed, so this is the month so far/);
  assert.match(html, /foot\.countByMonth\[m\] \? fmtShort\(foot\.byMonth\[m\]\) : '—'/);
  // Pinned to the bottom of the scroller, or it is the one row you have to
  // scroll to find.
  assert.match(html, /tfoot td\{position:sticky;bottom:0/);
  // And the export carries the same row.
  assert.match(html, /const ft = breakdownTotals\(rows\);/);
});

// --- Pages that fill the window ----------------------------------------------
//
// "and keep these to full page" — a card stopping halfway up the screen with
// the footnote stranded in the grey below it.

test('a page showing one card gives it the window, a stack of cards keeps scrolling', () => {
  assert.match(html, /const onPage = \[\.\.\.document\.querySelectorAll\('main \.card\[data-page\]'\)\]\.filter\(c => !c\.classList\.contains\('offpage'\)\)/);
  assert.match(html, /if \(onPage\.length === 1\) onPage\[0\]\.classList\.add\('fill'\)/);
  assert.match(html, /document\.body\.classList\.toggle\('page-fill', onPage\.length === 1\)/);
  // Cleared first, or yesterday's page keeps stretching.
  assert.match(html, /document\.querySelectorAll\('main \.card\[data-page\]'\)\.forEach\(c => c\.classList\.remove\('fill'\)\)/);
  // The footnote sits at the bottom whether or not the page is framed.
  assert.match(html, /\.foot\{[^}]*margin-top:auto/);
  // Framing only where there is room for it: a phone or a short window scrolls.
  assert.match(html, /@media \(min-width:901px\) and \(min-height:620px\)\{\s*\n\s*body\.page-fill \.content\{height:100vh/);
  assert.match(html, /body\.page-fill main\{flex:1 1 auto;min-height:0;overflow:hidden\}/);
  // A hidden tab must not be revealed by the rule that passes height down.
  for (const id of ['loginTable', 'invChecklist', 'invFiles', 'appsVendor', 'pendingList']) {
    assert.match(html, new RegExp(`#${id}:not\\(\\.hidden\\)`), `${id} keeps its hidden state`);
  }
  // By vendor was the one the hardcoded list was missing, and its table ran out
  // of the card and over the footnote. A :has() rule catches the next one
  // without anybody remembering to add it.
  assert.match(html, /\.card\.fill > \*:not\(\.hidden\):has\(\.tblwrap,\.loginwrap,\.renewals,\.inv-grid,\.inv-files,\.ov\)/);
});

test('the page is as wide as the header above it', () => {
  // main stopped at 1440px while the topbar did not, so a wide screen put the
  // card in a column narrower than its own title with grey either side.
  const main = html.match(/\n  main\{[^}]*\}/)[0];
  assert.ok(!/max-width/.test(main), `main must not cap its width: ${main.trim()}`);
  assert.match(main, /padding:8px 28px 10px/, 'the same gutter the topbar uses');
});

test('Logout cannot fall off the end of the sidebar', () => {
  // The name, the editor chip and Logout shared one nowrap line, and Logout was
  // what the ellipsis ate.
  assert.match(html, /\.side-foot \.who\{[^}]*flex-wrap:wrap/);
  assert.match(html, /\.side-foot \.who-name\{[^}]*text-overflow:ellipsis/, 'only the name is clipped');
  assert.ok(!/\.side-foot \.who\{font-weight:600;color:var\(--txt\);overflow:hidden;text-overflow:ellipsis;white-space:nowrap\}/.test(html));
});
