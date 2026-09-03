// Run with: node --test
//
// Which row an invoice belongs to is decided by rules the owner owns — the
// sender's domain or subject picks the vendor, the invoice's own wording picks
// the row where one vendor bills several — and nothing else. Anything the
// rules do not settle is a question, never a guess.

const test = require('node:test');
const assert = require('node:assert');
const { SEED_RULES, normalizeRules, classify, learn, accountPrefix, phraseIn } = require('../lib/invoices/rules');

const APPS = ['Anthropic(Api Console)', 'Claude Ai', 'Claude Ai Max 6 Accounts', 'Google Voice', 'GOOGLE ADS', 'Google Workspace', 'Google cloud', 'Cumul(Luzmo)', 'Adobe'];
const rules = normalizeRules(SEED_RULES);

// Verbatim from the archive: the Anthropic invoices for the API console and
// for the Claude Team plan are the same sender, the same subject, the same
// layout — only the invoice number's account prefix differs.
const API_CONSOLE = 'Invoice number   Q8MUNTUC   0008  Date of issue   March 1, 2026  One-time credit purchase   1   $250.00';
const TEAM_PLAN = 'Invoice number   2FSKIDHO   0043  Date of issue   June 1, 2026  Auto recharge extra usage, Team plan   1   $95.32';
const stripe = (name) => ({ address: 'invoice+statements@stripe.com', senderName: name, subject: `Your ${name} receipt`, attachmentNames: ['Invoice.pdf'] });

test('Anthropic: the invoice-number prefix picks the row, not the wording', () => {
  const api = classify(rules, stripe('Anthropic, PBC'), API_CONSOLE, APPS);
  assert.strictEqual(api.app, 'Anthropic(Api Console)');
  assert.strictEqual(api.confident, true);
  // "Team plan" and "extra usage" appear on the seats invoice; the prefix is
  // what settles it, so the same words on the other account would not mislead.
  const seats = classify(rules, stripe('Anthropic, PBC'), TEAM_PLAN, APPS);
  assert.strictEqual(seats.app, 'Claude Ai');
});

test('Anthropic: an invoice with neither prefix is a question, with the three rows as options', () => {
  const v = classify(rules, stripe('Anthropic, PBC'), 'Invoice number 12345 Total $5.00', APPS);
  assert.strictEqual(v.app, null);
  assert.strictEqual(v.reason, 'no-line-item-rule');
  assert.deepStrictEqual(v.options, ['Anthropic(Api Console)', 'Claude Ai', 'Claude Ai Max 6 Accounts']);
  assert.match(v.question, /Which is it/);
});

test('Google: the product named in the invoice picks the row', () => {
  const google = { address: 'payments-noreply@google.com', subject: 'Your Google invoice is available', attachmentNames: ['5571764174.pdf'] };
  // Google Voice's invoice also says "Google Workspace Telecom"; Voice is
  // listed first in the rule, so it wins.
  assert.strictEqual(classify(rules, google, 'Google Cloud - Google Workspace Telecom  Google Voice Inc.  Google Voice Starter', APPS).app, 'Google Voice');
  assert.strictEqual(classify(rules, google, 'Statement  Google Ads  Summary for Mar 1, 2026', APPS).app, 'GOOGLE ADS');
  assert.strictEqual(classify(rules, google, 'Google Workspace  Business Starter', APPS).app, 'Google Workspace');
  // The "google → GOOGLE ADS" catch-all is gone: a Google invoice naming no
  // product is a question.
  assert.strictEqual(classify(rules, google, 'Invoice  Total in USD $5.00', APPS).app, null);
});

test('a sender with no rule is a question, not a fuzzy guess', () => {
  const v = classify(rules, { address: 'billing@newvendor.io', subject: 'Invoice #77' }, 'Total $9.00', APPS);
  assert.strictEqual(v.app, null);
  assert.strictEqual(v.reason, 'no-vendor-rule');
  assert.strictEqual(v.domain, 'newvendor.io');
});

test('the original sender of a forwarded mail beats the colleague who forwarded it', () => {
  const v = classify(rules, { address: 'anudeep@sarasanalytics.com', originalAddress: 'billing@luzmo.com', subject: 'FW: invoice' }, '', APPS);
  assert.strictEqual(v.app, 'Cumul(Luzmo)');
  assert.strictEqual(v.via, 'domain');
});

test('a subdomain matches its parent, an unrelated domain does not', () => {
  assert.strictEqual(classify(rules, { address: 'no-reply@mail.adobe.com', subject: '' }, '', APPS).app, 'Adobe');
  assert.strictEqual(classify(rules, { address: 'x@notadobe.com', subject: '' }, '', APPS).app, null);
});

test('a rule naming a row the sheet does not have is a question', () => {
  const v = classify(rules, { address: 'x@luzmo.com', subject: '' }, '', ['Adobe']);
  assert.strictEqual(v.reason, 'app-not-in-sheet');
});

test('phrases match whole words only', () => {
  assert.strictEqual(phraseIn('Hex', 'a hexadecimal number'), false);
  assert.strictEqual(phraseIn('Hex', 'Hex Technologies'), true);
  assert.strictEqual(phraseIn('Q8MUNTUC', 'Invoice number Q8MUNTUC 0180'), true);
});

// --- Remembering an answer ------------------------------------------------

test('answering an unknown sender adds a rule for its domain', () => {
  const r = learn(rules, { reason: 'no-vendor-rule', domain: 'newvendor.io' }, 'Adobe');
  assert.match(r.learned, /newvendor\.io → Adobe/);
  assert.strictEqual(classify(r.rules, { address: 'x@newvendor.io', subject: '' }, '', APPS).app, 'Adobe');
});

test('a Stripe sender is remembered by its name, never by stripe.com', () => {
  const r = learn(rules, { reason: 'no-vendor-rule', domain: 'stripe.com', senderName: 'Acme Widgets' }, 'Adobe');
  assert.match(r.learned, /"Acme Widgets" → Adobe/);
  assert.ok(!r.rules.vendors.some(v => v.domains.includes('stripe.com')), 'stripe.com would match every Stripe invoice');
});

test('answering a several-row vendor remembers the invoice-number prefix', () => {
  const r = learn(rules, { reason: 'no-line-item-rule', vendor: 'Anthropic', ref: 'ZZZZ1234' + '0007' }, 'Claude Ai Max 6 Accounts');
  assert.match(r.learned, /ZZZZ1234-… → Claude Ai Max 6 Accounts/);
  const v = classify(r.rules, stripe('Anthropic, PBC'), 'Invoice number ZZZZ1234 0007', APPS);
  assert.strictEqual(v.app, 'Claude Ai Max 6 Accounts');
});

test('nothing is learned when nothing safe can be inferred', () => {
  assert.strictEqual(learn(rules, { reason: 'no-line-item-rule', vendor: 'Anthropic', ref: 'INV50264' }, 'Claude Ai').learned, null);
  assert.strictEqual(learn(rules, { reason: 'no-vendor-rule', domain: 'gmail.com' }, 'Adobe').learned, null);
});

test('only the Stripe shape of invoice number yields a prefix', () => {
  assert.strictEqual(accountPrefix('Q8MUNTUC0180'), 'Q8MUNTUC');
  assert.strictEqual(accountPrefix('INV50264'), null);
  assert.strictEqual(accountPrefix('12345678' + '0001'), null, 'all digits is not an account prefix');
});

test('the rules file survives a careless edit', () => {
  const r = normalizeRules({ vendors: [null, { name: 'X' }, { domains: 'one.com', app: 'Adobe' }, { name: 'Y', apps: [{ app: 'Adobe', text: 'word' }, {}] }] });
  assert.strictEqual(r.vendors.length, 2);
  assert.deepStrictEqual(r.vendors[0].domains, ['one.com']);
  assert.strictEqual(r.vendors[1].apps.length, 1);
  assert.deepStrictEqual(r.vendors[1].apps[0].text, ['word']);
});


test('locks survive normalisation and match the sheet\'s spelling loosely', () => {
  const { lockFor } = require('../lib/invoices/rules');
  const r = normalizeRules({ vendors: [{ name: 'X', domains: ['x.com'], app: 'Adobe' }], locks: [
    { app: 'Cursor pro', month: '2026-07', value: '1133.53', note: 'invoices missing' },
    { app: 'Claude Ai', month: '2026-08', value: 0 },
    { app: 'Bad', month: 'July' }, { app: '', month: '2026-01', value: 5 },
  ] });
  assert.strictEqual(r.locks.length, 2);
  assert.strictEqual(r.locks[0].value, 1133.53);
  assert.strictEqual(lockFor(r, 'Cursor Pro', '2026-07').value, 1133.53);
  assert.strictEqual(lockFor(r, 'Claude Ai', '2026-08').value, 0);
  assert.strictEqual(lockFor(r, 'Claude Ai', '2026-07'), null);
});
