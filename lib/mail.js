// Pulls invoice emails out of the shared invoices mailbox via Graph.
//
// Needs the `Mail.Read` APPLICATION permission on the app registration, with
// admin consent. Scope it to just this mailbox with an application access
// policy so the app cannot read anyone else's mail:
//   New-ApplicationAccessPolicy -AppId <client id> \
//     -PolicyScopeGroupId <mailbox or group> -AccessRight RestrictAccess
// Without that permission every call here comes back 403, which readInvoiceMail
// turns into a clear message rather than a stack trace.

const INVOICE_SUBJECT_RE = /invoice|receipt|billing|bill\b|payment|statement|subscription|renewal|paid|order confirmation/i;
// Mail that merely mentions billing but is plainly not an invoice.
const NOISE_SUBJECT_RE = /verify your|confirm your|password|sign ?in|welcome|newsletter|webinar|survey|unsubscribe|expiring soon|action required to keep/i;

function mailboxAddress() {
  return (process.env.INVOICE_MAILBOX || 'invoices@sarasanalytics.com').trim();
}

async function graphGet(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    if (res.status === 403) {
      throw new Error(
        `Graph refused to read the mailbox (403). The app registration needs the "Mail.Read" application permission with admin consent. Detail: ${text.slice(0, 160)}`
      );
    }
    if (res.status === 404) {
      throw new Error(`Mailbox not found (404). Check the INVOICE_MAILBOX env var. Detail: ${text.slice(0, 160)}`);
    }
    throw new Error(`Graph mail call failed (${res.status}): ${text.slice(0, 200)}`);
  }
  return res.json();
}

// Messages with attachments, newest first, since a given time.
async function listMessages(token, mailbox, sinceIso, top) {
  const filters = ['hasAttachments eq true'];
  if (sinceIso) filters.push(`receivedDateTime ge ${sinceIso}`);
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages`
    + `?$filter=${encodeURIComponent(filters.join(' and '))}`
    + `&$select=id,subject,from,sender,receivedDateTime,bodyPreview,hasAttachments,webLink`
    + `&$orderby=receivedDateTime desc&$top=${Math.min(top || 50, 100)}`;
  const json = await graphGet(token, url);
  return json.value || [];
}

async function listAttachments(token, mailbox, messageId) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}`
    + `/attachments?$select=id,name,contentType,size,isInline`;
  const json = await graphGet(token, url);
  return json.value || [];
}

async function getAttachmentBytes(token, mailbox, messageId, attachmentId) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}`
    + `/attachments/${encodeURIComponent(attachmentId)}`;
  const json = await graphGet(token, url);
  if (!json.contentBytes) throw new Error(`Attachment "${json.name || attachmentId}" has no downloadable content`);
  return Buffer.from(json.contentBytes, 'base64');
}

const isPdf = (att) => !att.isInline && (/pdf/i.test(att.contentType || '') || /\.pdf$/i.test(att.name || ''));

// Is this message actually an invoice, or just mailbox noise? Requires a PDF
// attachment plus invoice-ish wording, and rejects the obvious account-admin mail
// that also mentions billing.
function looksLikeInvoice(message, attachments) {
  if (!attachments.some(isPdf)) return false;
  const subject = String(message.subject || '');
  if (NOISE_SUBJECT_RE.test(subject)) return false;
  const haystack = `${subject} ${String(message.bodyPreview || '').slice(0, 400)} ${attachments.map(a => a.name || '').join(' ')}`;
  return INVOICE_SUBJECT_RE.test(haystack);
}

// Who the invoice is really from. Forwarded mail ("FW: [Bubble] Invoice") arrives
// from a colleague, so the original sender in the quoted header matters more than
// the envelope sender.
function senderIdentity(message) {
  const envelope = (message.from && message.from.emailAddress) || (message.sender && message.sender.emailAddress) || {};
  const body = String(message.bodyPreview || '');
  const forwarded = body.match(/From:\s*[^<\n]*<([^>]+)>/i) || body.match(/From:\s*([^\s<>\n]+@[^\s<>\n]+)/i);
  return {
    address: String(envelope.address || '').toLowerCase(),
    name: String(envelope.name || ''),
    originalAddress: forwarded ? String(forwarded[1]).toLowerCase() : '',
  };
}

// Everything worth matching an app against, richest signal first.
function matchText(message, attachments, identity) {
  const domain = (addr) => {
    const at = String(addr || '').indexOf('@');
    return at === -1 ? '' : addr.slice(at + 1).replace(/\.(com|io|net|org|ai|co|tech|app)$/i, '');
  };
  return [
    String(message.subject || ''),
    attachments.filter(isPdf).map(a => a.name || '').join(' '),
    domain(identity.originalAddress) || domain(identity.address),
    identity.name,
    String(message.bodyPreview || '').slice(0, 300),
  ].filter(Boolean);
}

// The month the invoice belongs to, from the received date.
function messageMonth(message) {
  const d = new Date(message.receivedDateTime);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

module.exports = {
  mailboxAddress, listMessages, listAttachments, getAttachmentBytes,
  isPdf, looksLikeInvoice, senderIdentity, matchText, messageMonth,
  INVOICE_SUBJECT_RE, NOISE_SUBJECT_RE,
};
