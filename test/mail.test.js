// Run with: node --test
//
// Covers the judgement calls in mailbox ingestion: telling a real invoice from
// mailbox noise, and working out which app an invoice belongs to when the mail
// has been forwarded by a colleague.

const test = require('node:test');
const assert = require('node:assert');

const mail = require('../lib/mail');
const { buildResolver } = require('../lib/vendor-map');

const APPS = ['Bubble Starter', 'Adobe', 'Github', 'WEBFLOW', 'Bitly', 'Cursor pro', 'Sentry.io', 'Anthropic(Api Console)'];
const resolve = buildResolver({}, APPS);

const pdf = (name) => ({ id: 'a1', name, contentType: 'application/pdf', size: 1024, isInline: false });

// Taken from the real invoices@sarasanalytics.com mailbox.
const BUBBLE_FWD = {
  id: 'm1',
  subject: 'FW: [BULK] [Bubble] Invoice',
  from: { emailAddress: { address: 'anudeep.kolla@sarasanalytics.com', name: 'Anudeep Kolla' } },
  receivedDateTime: '2026-08-25T08:56:07.000Z',
  bodyPreview: 'From: Bubble <noreply@bubble.io>\r\nSent: Tuesday, 25 August, 2026\r\nTo: Audit Agent <audit@sarasanalytics.com>',
};
const BITLY_NOISE = {
  id: 'm2',
  subject: 'Verify your Bitly email address',
  from: { emailAddress: { address: 'support@accounts.bitly.com', name: 'Bitly' } },
  receivedDateTime: '2026-08-25T08:46:08.000Z',
  bodyPreview: 'One last step! Confirm your email address to finish setting up billing.',
};

test('treats a forwarded vendor invoice with a PDF as an invoice', () => {
  assert.strictEqual(mail.looksLikeInvoice(BUBBLE_FWD, [pdf('bubble-invoice-aug.pdf')]), true);
});

test('ignores account-admin mail even when it mentions billing', () => {
  // Has no attachment, and the subject is plainly not an invoice.
  assert.strictEqual(mail.looksLikeInvoice(BITLY_NOISE, []), false);
  assert.strictEqual(mail.looksLikeInvoice(BITLY_NOISE, [pdf('terms.pdf')]), false);
});

test('ignores a message whose only attachment is not a PDF', () => {
  const att = [{ id: 'a', name: 'logo.png', contentType: 'image/png', isInline: false }];
  assert.strictEqual(mail.looksLikeInvoice(BUBBLE_FWD, att), false);
});

test('ignores an inline image masquerading as an attachment', () => {
  const att = [{ id: 'a', name: 'sig.pdf', contentType: 'application/pdf', isInline: true }];
  assert.strictEqual(mail.looksLikeInvoice(BUBBLE_FWD, att), false);
});

test('reads the original sender out of a forwarded message', () => {
  const id = mail.senderIdentity(BUBBLE_FWD);
  assert.strictEqual(id.address, 'anudeep.kolla@sarasanalytics.com');
  assert.strictEqual(id.originalAddress, 'noreply@bubble.io');
});

test('matches a forwarded invoice to the right app, not the forwarder', () => {
  const attachments = [pdf('bubble-invoice-aug.pdf')];
  const identity = mail.senderIdentity(BUBBLE_FWD);
  let app = null;
  for (const text of mail.matchText(BUBBLE_FWD, attachments, identity)) {
    const hit = resolve('', text);
    if (hit.app) { app = hit.app; break; }
  }
  assert.strictEqual(app, 'Bubble Starter');
});

test('falls back to the sender domain when the subject says nothing useful', () => {
  const message = {
    id: 'm3',
    subject: 'Your receipt',
    from: { emailAddress: { address: 'billing@github.com', name: 'GitHub' } },
    receivedDateTime: '2026-08-09T00:00:00.000Z',
    bodyPreview: 'Thanks for your payment.',
  };
  const attachments = [pdf('receipt.pdf')];
  assert.strictEqual(mail.looksLikeInvoice(message, attachments), true);
  const identity = mail.senderIdentity(message);
  let app = null;
  for (const text of mail.matchText(message, attachments, identity)) {
    const hit = resolve('', text);
    if (hit.app) { app = hit.app; break; }
  }
  assert.strictEqual(app, 'Github');
});

test('derives the invoice month from the received date', () => {
  assert.strictEqual(mail.messageMonth(BUBBLE_FWD), '2026-08');
  assert.strictEqual(mail.messageMonth({ receivedDateTime: 'nonsense' }), null);
});

test('defaults to the shared invoices mailbox', () => {
  delete process.env.INVOICE_MAILBOX;
  assert.strictEqual(mail.mailboxAddress(), 'invoices@sarasanalytics.com');
  process.env.INVOICE_MAILBOX = 'other@example.com';
  assert.strictEqual(mail.mailboxAddress(), 'other@example.com');
  delete process.env.INVOICE_MAILBOX;
});

// --- Where invoices get filed -------------------------------------------
//
// Invoices belong in the archive they have always been filed in, under
// {archive}/{vendor}/{month}/ — not in a separate store. The vendor folder comes
// from the folder->app mapping the invoice import already saved, and the month
// subfolder reuses whatever is there. Where the archive itself lives is resolved
// at run time (see test/archive.test.js), never assumed.

const { appToSourceFolder, monthFolderName } = require('../lib/mail-sync');

// The real mapping saved in the archive's _sync-config.json.
const SAVED_MAPPING = {
  Adobe: 'Adobe',
  Bubble: 'Bubble Starter',
  Cursor: 'Cursor pro',
  'Claude Api': 'Anthropic(Api Console)',
  'click up invoices': 'Mango technology(Clickup)',
  'Laptop procurment': 'Laptops Procurement',
  Webflow: 'WEBFLOW',
};

test('routes an app back to the vendor folder it is archived under', () => {
  const byApp = appToSourceFolder(SAVED_MAPPING);
  assert.strictEqual(byApp['Bubble Starter'], 'Bubble');
  assert.strictEqual(byApp['Cursor pro'], 'Cursor');
  assert.strictEqual(byApp['Anthropic(Api Console)'], 'Claude Api');
  assert.strictEqual(byApp['Mango technology(Clickup)'], 'click up invoices');
  assert.strictEqual(byApp['WEBFLOW'], 'Webflow');
  // An app with no procurement folder has no mapping to fall back on.
  assert.strictEqual(byApp['Sentry.io'], undefined);
});

test('a vendor folder that exists beats one the mapping names but is gone', () => {
  // "Bubble" was renamed to "Bubble Starter" in the archive, but the saved
  // mapping still lists both and "Bubble" comes first. Filing there would have
  // recreated the old name and split the vendor across two folders, only one of
  // which is ever totalled — the same shape as the Luzmo / Cumul(Luzmo) split.
  const mapping = { Bubble: 'Bubble Starter', 'Bubble Starter': 'Bubble Starter' };
  assert.strictEqual(appToSourceFolder(mapping)['Bubble Starter'], 'Bubble',
    'with nothing to check against, the old first-wins rule still stands');
  assert.strictEqual(
    appToSourceFolder(mapping, ['Adobe', 'Bubble Starter'])['Bubble Starter'], 'Bubble Starter');
});

test('a folder still present is not swapped for a later one', () => {
  // Only ever trade up. Both exist here, so the first stays chosen rather than
  // the last listed quietly winning.
  const mapping = { 'Cumul(Luzmo)': 'Cumul(Luzmo)', Luzmo: 'Cumul(Luzmo)' };
  assert.strictEqual(
    appToSourceFolder(mapping, ['Cumul(Luzmo)', 'Luzmo'])['Cumul(Luzmo)'], 'Cumul(Luzmo)');
});

test('when no mapped folder exists, the first is still used so filing has a home', () => {
  const mapping = { Bubble: 'Bubble Starter', 'Bubble Starter': 'Bubble Starter' };
  assert.strictEqual(appToSourceFolder(mapping, ['Adobe'])['Bubble Starter'], 'Bubble');
});

test('folder matching ignores case, as SharePoint does', () => {
  // The archive holds "apollo"; the mapping says "Apollo".
  const mapping = { Apollo: 'Apollo', 'apollo old': 'Apollo' };
  assert.strictEqual(appToSourceFolder(mapping, ['apollo'])['Apollo'], 'Apollo');
});

test('a folder renamed to match the sheet beats the stale mapping entry', () => {
  // The archive was renamed folder-by-folder to match the sheet. The saved
  // mapping is a snapshot from before that, so it still points at names that no
  // longer exist: filing by "click up invoices" would recreate that folder
  // beside the real "Mango technology(Clickup)" and total only one of them.
  const folders = ['Anthropic(Api Console)', 'Mango technology(Clickup)', 'Cursor pro', 'Adobe'];
  const appNames = ['Anthropic(Api Console)', 'Mango technology(Clickup)', 'Cursor pro', 'Adobe'];
  const byApp = appToSourceFolder(SAVED_MAPPING, folders, appNames);
  assert.strictEqual(byApp['Anthropic(Api Console)'], 'Anthropic(Api Console)');
  assert.strictEqual(byApp['Mango technology(Clickup)'], 'Mango technology(Clickup)');
  assert.strictEqual(byApp['Cursor pro'], 'Cursor pro');
});

test('the app-named folder is used with the spelling the archive gives it', () => {
  // The sheet says "Apollo" and "Google cloud"; the archive says "apollo" and
  // "Google Cloud". Same folder either way, so match on the normalized form and
  // file under the name that is actually there.
  const byApp = appToSourceFolder({}, ['apollo', 'Google Cloud'], ['Apollo', 'Google cloud']);
  assert.strictEqual(byApp['Apollo'], 'apollo');
  assert.strictEqual(byApp['Google cloud'], 'Google Cloud');
});

test('an app with no folder of its own keeps its mapped folder', () => {
  // "Laptops Procurement" is spelled "Laptop procurment" in the archive and
  // always has been. Nothing carries the sheet's spelling, so the mapping — the
  // only thing that knows the two are the same vendor — has to stand.
  const byApp = appToSourceFolder(SAVED_MAPPING, ['Laptop procurment', 'Adobe'],
    ['Laptops Procurement', 'Adobe']);
  assert.strictEqual(byApp['Laptops Procurement'], 'Laptop procurment');
});

test('an app name matching no folder adds no entry', () => {
  // Every app in the sheet is offered, most of which have no folder. None of
  // them may invent one, or filing would stop reporting a new folder at all.
  const byApp = appToSourceFolder({}, ['Adobe'], ['Adobe', 'Sentry.io', 'Keepa']);
  assert.strictEqual(byApp['Adobe'], 'Adobe');
  assert.strictEqual(byApp['Sentry.io'], undefined);
  assert.strictEqual(byApp['Keepa'], undefined);
});

test('reuses an existing month subfolder rather than adding one beside it', () => {
  // Bubble already has "Aug"; Cursor already has "July". Neither should gain a
  // second folder for the same month.
  assert.strictEqual(monthFolderName(['Aug', 'July'], '2026-08'), 'Aug');
  assert.strictEqual(monthFolderName(['June', 'July'], '2026-07'), 'July');
  assert.strictEqual(monthFolderName(['Aug-26'], '2026-08'), 'Aug-26');
  assert.strictEqual(monthFolderName(['2026-08'], '2026-08'), '2026-08');
  assert.strictEqual(monthFolderName(['August 2026'], '2026-08'), 'August 2026');
  // Case and stray whitespace in the existing name must still match.
  assert.strictEqual(monthFolderName([' aug '], '2026-08'), ' aug ');
});

test('creates a dated month folder only when none exists', () => {
  assert.strictEqual(monthFolderName([], '2026-08'), 'Aug-26');
  assert.strictEqual(monthFolderName(['Quotations'], '2026-01'), 'Jan-26');
  // A folder for a different month must not be reused.
  assert.strictEqual(monthFolderName(['July'], '2026-08'), 'Aug-26');
});

test('returns nothing for a month it cannot parse', () => {
  assert.strictEqual(monthFolderName([], ''), null);
  assert.strictEqual(monthFolderName([], 'not-a-month'), null);
});

// --- The Graph query -----------------------------------------------------
//
// The first live run returned 400 InefficientFilter: "The restriction or sort
// order is too complex for this operation." Exchange rejects a filter on a
// non-indexed property (hasAttachments) combined with a sort. receivedDateTime
// is indexed and sorts fine, so only the date is filtered server-side.

test('does not filter on hasAttachments, which Exchange rejects alongside a sort', () => {
  const url = mail.messagesUrl('invoices@sarasanalytics.com', '2026-06-25T00:00:00Z', 50, true);
  assert.ok(!/hasAttachments\s*eq/i.test(decodeURIComponent(url)), 'must not filter on hasAttachments');
  assert.match(decodeURIComponent(url), /\$filter=receivedDateTime ge 2026-06-25T00:00:00Z/);
  assert.match(url, /\$orderby=receivedDateTime desc/);
});

test('still selects hasAttachments so the screening can happen in code', () => {
  const url = mail.messagesUrl('invoices@sarasanalytics.com', '2026-06-25T00:00:00Z', 50, true);
  assert.match(url, /\$select=[^&]*hasAttachments/);
});

test('drops the filter entirely for the fallback query', () => {
  const url = mail.messagesUrl('invoices@sarasanalytics.com', '2026-06-25T00:00:00Z', 50, false);
  assert.ok(!/\$filter/.test(url), 'fallback must send no filter at all');
  assert.match(url, /\$orderby=receivedDateTime desc/);
});

test('caps the page size Graph will accept', () => {
  assert.match(mail.messagesUrl('a@b.com', null, 500, true), /\$top=100/);
  assert.match(mail.messagesUrl('a@b.com', null, 50, true), /\$top=50/);
  assert.match(mail.messagesUrl('a@b.com', null, undefined, true), /\$top=50/);
});

test('escapes the mailbox address into the path', () => {
  const url = mail.messagesUrl('inv oices+x@b.com', null, 10, true);
  assert.ok(url.includes(encodeURIComponent('inv oices+x@b.com')));
});

// --- Ticking the invoice tracker ----------------------------------------
//
// The first live run reported "Ticked 2 cells" when only one cell existed to
// tick: two Bubble invoices arrived in the same month and each queued its own
// mark, so the same cell was written twice. The sheet was correct, the count
// was not.

const { planTrackerCells } = require('../lib/mail-sync');
const { locateGrid } = require('../lib/excel');

// A tracker shaped like the live one: TRUE/FALSE cells, date-serial headers.
const TRACKER_VALUES = [
  ['APPLICATION / SW / LICENSE', 'Department', 'POC', 'Renewal data', 'Recurring/Onetime', 'FREQUENCY', 46023, 46174, 46235],
  ['Adobe', 'Marketing', 'Bhavana', '', 'Recurring', 'Monthly', true, true, false],
  ['Bubble Starter', 'Product', 'Ganesh', '', 'Recurring', 'Monthly', false, false, false],
  ['Github', 'DE', 'Anudeep', '', 'Recurring', 'Monthly', true, true, false],
];
// Graph renders a boolean cell as "TRUE"/"FALSE" and a date header by its number
// format; every other cell's text matches its value.
const TRACKER_TEXT = TRACKER_VALUES.map((row, i) => row.map((c, j) => {
  if (i === 0 && j >= 6) return ['Jan-26', 'Jun-26', 'Aug-26'][j - 6];
  if (typeof c === 'boolean') return c ? 'TRUE' : 'FALSE';
  return c == null ? '' : String(c);
}));
const TRACKER_USED = { values: TRACKER_VALUES, text: TRACKER_TEXT, start: { col: 0, row: 0 } };
const trackerGrid = () => locateGrid(TRACKER_VALUES, TRACKER_TEXT);

test('ticks one cell however many invoices arrive for that app-month', () => {
  // Two Bubble invoices in August — one cell, written once.
  const cells = planTrackerCells(
    [{ app: 'Bubble Starter', month: '2026-08' }, { app: 'Bubble Starter', month: '2026-08' }],
    trackerGrid(), TRACKER_USED
  );
  assert.strictEqual(cells.length, 1);
  assert.strictEqual(cells[0].address, 'I3'); // Bubble Starter row 3, Aug-26 column I
  assert.strictEqual(cells[0].value, true);
});

test('still ticks separate cells for different apps or months', () => {
  const cells = planTrackerCells(
    [
      { app: 'Bubble Starter', month: '2026-08' },
      { app: 'Github', month: '2026-08' },
      { app: 'Bubble Starter', month: '2026-08' }, // duplicate of the first
    ],
    trackerGrid(), TRACKER_USED
  );
  assert.strictEqual(cells.length, 2);
  assert.deepStrictEqual(cells.map(c => c.address).sort(), ['I3', 'I4']);
});

test('leaves a cell alone when it is already ticked', () => {
  const cells = planTrackerCells([{ app: 'Adobe', month: '2026-01' }], trackerGrid(), TRACKER_USED);
  assert.strictEqual(cells.length, 0, 'Adobe Jan-26 is already TRUE');
});

test('skips a mark that has no row or no column in the tracker', () => {
  const cells = planTrackerCells(
    [
      { app: 'Not In The Tracker', month: '2026-08' },
      { app: 'Bubble Starter', month: '2027-03' }, // no such column
    ],
    trackerGrid(), TRACKER_USED
  );
  assert.strictEqual(cells.length, 0);
});

// --- Reseller-billed charges --------------------------------------------
//
// A dry run against the real June statement routed "Google ChatGPT" to
// GOOGLE ADS. It is an OpenAI charge billed through Google, and it inflated the
// proposed Google Ads figure by exactly its own 20.50.

test('routes a reseller-billed charge to the real vendor, not the reseller', () => {
  const apps = ['GOOGLE ADS', 'Google cloud', 'Google Workspace', 'OPENAI', 'Anthropic(Api Console)'];
  const r = buildResolver({}, apps);
  assert.strictEqual(r('', 'Google ChatGPT 6502530000 CA').app, 'OPENAI');
  // The genuine Google lines must still land where they did.
  assert.strictEqual(r('', 'GOOGLE *ADS710 06/01 PURCHASE Mountain View CA').app, 'GOOGLE ADS');
  assert.strictEqual(r('', 'GOOGLE *CLOUD 9Q4THP 6502530000 CA').app, 'Google cloud');
});

// --- The rescan window ---------------------------------------------------
//
// A live run reported "Scanned 0 messages": the sync only asks Graph for mail
// newer than its last successful run, and all four mailbox messages predated it.
// rescan skipped the seen-list but left that window alone, so it returned zero
// too — the flag did nothing. The window has to widen as well.

function sinceFor(state, opts, lookbackDays) {
  // Mirrors the expression in runMailSync.
  const lookback = new Date(Date.now() - lookbackDays * 86400000).toISOString();
  return opts.rescan ? lookback : (state.lastRunAt || lookback);
}

test('a normal run only looks at mail since the last successful run', () => {
  const since = sinceFor({ lastRunAt: '2026-08-25T11:07:00Z' }, {}, 60);
  assert.strictEqual(since, '2026-08-25T11:07:00Z');
});

test('a rescan widens the window instead of honouring the watermark', () => {
  const state = { lastRunAt: '2026-08-25T11:07:00Z' };
  const since = sinceFor(state, { rescan: true }, 60);
  assert.notStrictEqual(since, state.lastRunAt, 'rescan must not reuse the watermark');
  assert.ok(Date.parse(since) < Date.parse(state.lastRunAt), 'rescan must reach further back');
  // Messages that arrived before the last run are what a rescan exists to find.
  assert.ok(Date.parse('2026-08-25T08:46:00Z') >= Date.parse(since));
});

test('the first ever run falls back to the lookback window', () => {
  const since = sinceFor({}, {}, 60);
  assert.ok(Date.parse(since) < Date.now());
});

test('the source really does branch on rescan, not just the helper above', () => {
  const src = require('node:fs').readFileSync(
    require('node:path').join(__dirname, '..', 'lib', 'mail-sync.js'), 'utf8');
  assert.match(src, /const since = opts\.rescan \?/,
    'runMailSync must pick the window based on opts.rescan');
});
