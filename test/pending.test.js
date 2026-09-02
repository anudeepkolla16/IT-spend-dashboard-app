// Run with: node --test
//
// What the app does with an invoice it is not sure about: hold it, ask, and
// read the answer — from the dashboard or typed into the Slack DM.

const test = require('node:test');
const assert = require('node:assert');
const { parseReply, matchApp, describeQuestions } = require('../lib/invoices/pending');
const { formatRunReport, monthForInvoice } = require('../lib/mail-sync');
const slack = require('../lib/slack');

const APPS = ['Anthropic(Api Console)', 'Claude Ai', 'Google Voice', 'GOOGLE ADS', 'Adobe'];
const ITEMS = [
  { id: 'P3', file: 'Invoice-1.pdf', question: 'Which is it?', options: ['Anthropic(Api Console)', 'Claude Ai'], month: '2026-08' },
  { id: 'P4', file: 'x.pdf', question: 'Which app? Which month?', options: [], month: null },
];
const NOW = new Date('2026-09-02T00:00:00Z');

test('a reply names an app by option number or by name, with a month if given', () => {
  const r = parseReply('P3 = 2\nP4 = Google Voice, Aug-26', ITEMS, APPS, NOW);
  assert.deepStrictEqual(r.map(x => [x.id, x.app, x.month]), [['P3', 'Claude Ai', '2026-08'], ['P4', 'Google Voice', '2026-08']]);
});

test('the month falls back to the one the invoice stated', () => {
  const [r] = parseReply('P3: adobe', ITEMS, APPS, NOW);
  assert.strictEqual(r.app, 'Adobe');
  assert.strictEqual(r.month, '2026-08');
});

test('an ambiguous or unknown app is an error, never a guess', () => {
  const r = parseReply('P4 = goo\nP4 = Nothing Like This\nP9 = Adobe', ITEMS, APPS, NOW);
  assert.match(r[0].error, /could be Google Voice, GOOGLE ADS/);
  assert.match(r[1].error, /not a row in the sheet/);
  assert.match(r[2].error, /not an open question/);
});

test('"ignore" is an answer', () => {
  assert.deepStrictEqual(parseReply('p4 - ignore', ITEMS, APPS, NOW)[0], { id: 'P4', ignore: true, line: 'p4 - ignore' });
});

test('chat that is not an answer is left alone', () => {
  assert.deepStrictEqual(parseReply('thanks, looks good\nwill check tomorrow', ITEMS, APPS, NOW), []);
});

test('an app is matched the way the sheet spells it', () => {
  assert.strictEqual(matchApp('google ads', APPS).app, 'GOOGLE ADS');
  assert.strictEqual(matchApp('api console', APPS).app, 'Anthropic(Api Console)');
});

test('the DM lists each question with its id, options and the reply format', () => {
  const text = describeQuestions(ITEMS);
  assert.match(text, /\*P3\*/);
  assert.match(text, /1\) Anthropic\(Api Console\)   2\) Claude Ai/);
  assert.match(text, /and the month/);
  assert.match(text, /P12 = Google Voice, Aug-26/);
});

// --- The month, by the owner's rules ----------------------------------------

test('the billing period\'s start month wins, then the invoice date, then nothing', () => {
  assert.strictEqual(monthForInvoice('Billing period: 01-08-2026 to 31-08-2026', '2026-09').month, '2026-08');
  const dated = monthForInvoice('Invoice number 5 Date of issue June 1, 2026 Total $5', '2026-07');
  assert.strictEqual(dated.month, '2026-06');
  assert.strictEqual(dated.via, 'invoice-date');
  const none = monthForInvoice('Total $5.00', '2026-07');
  assert.strictEqual(none.month, null, 'the mail\'s month is never assumed');
  assert.match(none.why, /nothing in the invoice says/);
});

test('a period months away from the mail is a question, not a filing', () => {
  const r = monthForInvoice('Subscription period 2027-06-01 to 2027-06-30 Date of issue June 1, 2026', '2026-09');
  assert.strictEqual(r.month, null);
  assert.strictEqual(r.via, 'period-far');
});

// --- The report --------------------------------------------------------------

test('a run that changed nothing and asks nothing sends nothing', () => {
  assert.strictEqual(formatRunReport({ filed: 0, amountsWritten: [], amountsUpdated: [], errors: [] }, []), null);
});

test('the report says what was filed, written, ticked, lowered and asked', () => {
  const text = formatRunReport({
    filed: 2, perApp: { Adobe: 2 },
    amountsWritten: [{ app: 'Adobe', month: '2026-08', amount: 37.16 }],
    amountsUpdated: [{ app: 'Claude Ai', month: '2026-08', previous: 500, amount: 101.28, direction: 'down', wasOurs: false }],
    tracker: { marked: 2 }, errors: [],
    answered: [{ id: 'P1', file: 'a.pdf', app: 'Adobe', month: '2026-08' }],
    learned: ['sender domain x.io → Adobe'],
  }, ITEMS);
  assert.match(text, /Filed 2 invoices/);
  assert.match(text, /Adobe 2026-08: set to 37\.16/);
  assert.match(text, /500\.00 → 101\.28 \(lowered/);
  assert.match(text, /replaced a figure not from invoices/);
  assert.match(text, /Ticked 2 cells/);
  assert.match(text, /P1 a\.pdf → Adobe 2026-08 \(your answer\)/);
  assert.match(text, /remembered: sender domain x\.io → Adobe/);
  assert.match(text, /2 invoices waiting for your answer/);
});

test('without a token the Slack client is a no-op that says so', async () => {
  const saved = [process.env.SLACK_BOT_TOKEN, process.env.SLACK_DM_USER];
  delete process.env.SLACK_BOT_TOKEN;
  delete process.env.SLACK_DM_USER;
  try {
    assert.strictEqual(slack.configured(), false);
    const r = await slack.postDm('hello');
    assert.strictEqual(r.sent, false);
    assert.match(r.why, /SLACK_BOT_TOKEN/);
    assert.deepStrictEqual((await slack.readReplies(null)).messages, []);
  } finally {
    if (saved[0]) process.env.SLACK_BOT_TOKEN = saved[0];
    if (saved[1]) process.env.SLACK_DM_USER = saved[1];
  }
});
