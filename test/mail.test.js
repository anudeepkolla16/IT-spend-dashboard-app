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
