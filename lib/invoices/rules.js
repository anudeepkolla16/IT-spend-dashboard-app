// Which app an invoice belongs to, decided by rules the sheet's owner owns.
//
// The owner's process, in their words:
//   1. check for sender domain or subject for application matching name
//   3. check the line items for application names and subsection details
//   4. never guess or mismatch — if you are not clear, ask me
//   7. if you want application mapping, ask me; I will help for a one-time setup
//
// So the rules live in a JSON file beside the invoices, `_vendor-rules.json`,
// that the dashboard shows and lets the owner edit. Nothing here is fuzzy: an
// invoice is filed only when a rule names its app outright, and everything
// else is held with a question. Answers are written back into the file, so
// each vendor is asked about once.
//
// The shape:
//   {
//     "vendors": [
//       { "name": "Luzmo", "domains": ["luzmo.com"], "app": "Cumul(Luzmo)" },
//       { "name": "Anthropic", "domains": ["anthropic.com"], "subject": ["Anthropic"],
//         "apps": [
//           { "app": "Anthropic(Api Console)", "text": ["Q8MUNTUC"] },
//           { "app": "Claude Ai",              "text": ["2FSKIDHO"] }
//         ] }
//     ]
//   }
//
//   domains  the sender's domain, or any parent of it ("mail.anthropic.com"
//            matches "anthropic.com"). The original sender of a forwarded mail
//            counts before the colleague who forwarded it.
//   subject  words looked for in the subject, the attachment names and the
//            sender's display name, for vendors that bill through Stripe and
//            the like, where the domain says nothing.
//   app      the row in the sheet, when the vendor bills exactly one.
//   apps     when one vendor bills several rows, in order: the first whose
//            `text` phrases appear in the invoice wins. Anthropic's invoice
//            numbers carry the billing account's own prefix (Q8MUNTUC-0180 is
//            the API console, 2FSKIDHO-0043 the Claude Team plan), which is
//            the one thing on those otherwise identical invoices that tells
//            them apart. Nothing matching means: ask.
//   currency for a vendor whose invoice states two currencies (PhantomBuster
//            prints EUR and USD), which one the sheet takes.
//   locks    (top level, beside "vendors") cells the owner has set by hand:
//            [{ "app", "month", "value", "note" }]. The sync keeps the cell at
//            that value and never totals invoices into it.
//   period   "usage" for a metered vendor whose line items ("Claude Opus 4.8
//            Usage Jul 2 – Jul 31, 2026") are the billing period; the invoice
//            issued on 1 August is then July's, as the owner's rule says.

const { norm, resolveToSheetApp } = require('../vendor-map');

const RULES_FILE = '_vendor-rules.json';

// The starting rules. Only what has been checked against real invoices in
// the archive or is the vendor's own domain; a vendor not listed here is asked
// about the first time, and the answer is remembered.
const SEED_RULES = {
  version: 1,
  vendors: [
    { name: 'Anthropic', domains: ['anthropic.com'], subject: ['Anthropic'], period: 'usage',
      apps: [
        { app: 'Anthropic(Api Console)', text: ['Q8MUNTUC'] },
        { app: 'Claude Ai', text: ['2FSKIDHO'] },
        { app: 'Claude Ai Max 6 Accounts', text: ['XQRYKLO3'] },
      ] },
    { name: 'Google', domains: ['google.com'],
      apps: [
        { app: 'Google Voice', text: ['Google Voice'] },
        { app: 'GOOGLE ADS', text: ['Google Ads'] },
        { app: 'Google cloud', text: ['Google Cloud Platform', 'Cloud Billing', 'Google Cloud'] },
        { app: 'Google Workspace', text: ['Google Workspace'] },
      ] },
    { name: 'Luzmo', domains: ['luzmo.com', 'cumul.io'], subject: ['Luzmo', 'Cumul'], app: 'Cumul(Luzmo)' },
    { name: 'Adobe', domains: ['adobe.com'], app: 'Adobe' },
    { name: 'Amazon Web Services', domains: ['amazon.com', 'amazonaws.com', 'aws.amazon.com'], subject: ['Amazon Web Services', 'AWS'], app: 'AWS' },
    { name: 'Apify', domains: ['apify.com'], app: 'Apify' },
    { name: 'Apollo', domains: ['apollo.io'], subject: ['Apollo'], app: 'Apollo' },
    { name: 'Bitly', domains: ['bitly.com', 'bit.ly'], app: 'Bitly' },
    { name: 'Bubble', domains: ['bubble.io'], subject: ['Bubble'], app: 'Bubble Starter' },
    { name: 'Canva', domains: ['canva.com'], app: 'CANVA' },
    { name: 'Chargebee', domains: ['chargebee.com'], app: 'Chargebee' },
    { name: 'Cursor', domains: ['cursor.com', 'cursor.sh', 'anysphere.inc'], subject: ['Cursor'], app: 'Cursor pro' },
    { name: 'dbt Labs', domains: ['getdbt.com', 'dbtlabs.com'], subject: ['dbt'], app: 'DBT Cloud' },
    { name: 'Dovetail', domains: ['dovetail.com', 'dovetailapp.com'], app: 'Dovetail' },
    { name: 'ElevenLabs', domains: ['elevenlabs.io'], subject: ['ElevenLabs'], app: 'ElevenLabs' },
    { name: 'Envato', domains: ['envato.com'], app: 'Envato' },
    { name: 'Figma', domains: ['figma.com'], app: 'FIGMA' },
    { name: 'Filestack', domains: ['filestack.com'], app: 'Filestack' },
    { name: 'GitHub', domains: ['github.com'], app: 'Github' },
    { name: 'GoDaddy', domains: ['godaddy.com'], app: 'Godaddy' },
    { name: 'Granola', domains: ['granola.ai', 'granola.so'], subject: ['Granola'], app: ['Granola', 'Granola Business'] },
    { name: 'Helpjuice', domains: ['helpjuice.com'], app: 'Helpjuice' },
    { name: 'Hex', domains: ['hex.tech'], app: 'Hex' },
    { name: 'HubSpot', domains: ['hubspot.com'], app: 'Hubspot' },
    { name: 'JetBrains', domains: ['jetbrains.com'], app: 'JetBrains' },
    { name: 'Keka', domains: ['keka.com'], app: 'Keka' },
    { name: 'Keepa', domains: ['keepa.com'], app: 'Keepa' },
    { name: 'LastPass', domains: ['lastpass.com'], app: 'LastPass' },
    { name: 'LinkedIn', domains: ['linkedin.com'], app: 'Linkedin' },
    { name: 'LottieFiles', domains: ['lottiefiles.com'], app: 'LOTTIEFILES' },
    { name: 'MetalpriceAPI', domains: ['metalpriceapi.com'], app: 'MetalPriceAPI' },
    { name: 'Tata Tele (Microsoft)', domains: ['tatatele.in', 'tatatelebusiness.com'], subject: ['Tata Tele'], app: 'MICROSOFT(Tata Tele)' },
    { name: 'ClickUp', domains: ['clickup.com'], subject: ['ClickUp', 'Mango Technologies'], app: 'Mango technology(Clickup)' },
    { name: 'Naukri', domains: ['naukri.com'], app: 'Naukri' },
    { name: 'OpenAI', domains: ['openai.com'], subject: ['OpenAI', 'ChatGPT'], app: 'OPENAI' },
    { name: 'PagerDuty', domains: ['pagerduty.com'], app: 'Pagerduty' },
    { name: 'PhantomBuster', domains: ['phantombuster.com'], subject: ['PhantomBuster'], app: 'Phantombuster', currency: 'USD' },
    { name: 'PostHog', domains: ['posthog.com'], subject: ['PostHog'], app: 'Posthog' },
    { name: 'Product Fruits', domains: ['productfruits.com'], subject: ['Product Fruits'], app: 'Product Fruits' },
    { name: 'Prosp', domains: ['prosp.ai'], subject: ['Prosp'], app: 'Prosp AI' },
    { name: 'Render', domains: ['render.com'], subject: ['Render'], app: 'Render' },
    { name: 'Sentry', domains: ['sentry.io'], subject: ['Sentry'], app: 'Sentry.io' },
    { name: 'Slack', domains: ['slack.com'], app: 'SLACK' },
    { name: 'Sprinto', domains: ['sprinto.com'], app: 'Sprinto' },
    { name: 'Superhuman', domains: ['superhuman.com'], app: 'Superhuman' },
    { name: 'Shopify', domains: ['shopify.com'], subject: ['Shopify'], app: 'Shopify Inc' },
    { name: 'TAC Security', domains: ['tacsecurity.com'], app: 'TAC Security' },
    { name: 'T-Mobile', domains: ['t-mobile.com'], subject: ['T-Mobile'], app: 'TMobile' },
    { name: 'Twilio SendGrid', domains: ['twilio.com', 'sendgrid.com', 'sendgrid.net'], subject: ['Twilio', 'SendGrid'], app: 'Twilo Sendgrid' },
    { name: 'Typeform', domains: ['typeform.com'], app: 'Typeform' },
    { name: 'Vimcal', domains: ['vimcal.com'], app: 'vimcal' },
    { name: 'Webflow', domains: ['webflow.com'], subject: ['Webflow'], app: 'WEBFLOW' },
    { name: 'Clari (Wingman)', domains: ['clari.com', 'trywingman.com'], subject: ['Clari', 'Wingman'], app: 'WINGMAN(Clari)' },
    { name: 'Windsurf', domains: ['windsurf.com', 'codeium.com'], subject: ['Windsurf', 'Codeium'], app: ['windsurf', 'windsurf pro'] },
    { name: 'X Corp', domains: ['x.com', 'twitter.com'], subject: ['X Corp'], app: 'X corp' },
    { name: 'Zoho', domains: ['zoho.com', 'zohocorp.com', 'zoho.in'], subject: ['Zoho'], app: 'ZOHO Books' },
    { name: 'Zoom', domains: ['zoom.us', 'zoom.com'], subject: ['Zoom'], app: 'ZOOM' },
  ],
};

// Cells the owner named on 3 September 2026 ("never change these amounts in
// next runs, these are correct figures"). Seeded into the live rules file the
// first time it is read without a `locks` list at all; from then on the file
// is the owner's and these are not consulted again.
const SEED_LOCKS = [
  { app: 'Claude Ai', month: '2026-07', value: 0, note: 'Set by the owner on 3 Sep 2026: paid from prepaid credits, not a charge for the month.' },
  { app: 'Claude Ai', month: '2026-08', value: 0, note: 'Set by the owner on 3 Sep 2026: paid from prepaid credits, not a charge for the month.' },
  { app: 'Cursor pro', month: '2026-08', value: 0, note: 'Set by the owner on 3 Sep 2026: paid from prepaid credits, not a charge for the month.' },
  { app: 'Bubble Starter', month: '2026-07', value: 745.89, note: "Set by the owner on 3 Sep 2026: some of the month's invoices are missing from the archive; this is the correct figure." },
  { app: 'Cursor pro', month: '2026-07', value: 1133.53, note: "Set by the owner on 3 Sep 2026: some of the month's invoices are missing from the archive; this is the correct figure." },
];

// What a live rules file is missing that the seed has since gained: a vendor
// setting the seed added after the file was first written (Anthropic's
// `period: "usage"`), and the owner's locks when the file has never had a
// `locks` list. Only ABSENT things are filled in — a value the owner set,
// even to nothing, is theirs. Returns { rules, changed }.
function upgradeRules(raw) {
  const live = raw && typeof raw === 'object' ? raw : {};
  let changed = false;
  const vendors = Array.isArray(live.vendors) ? live.vendors : [];
  for (const seed of SEED_RULES.vendors) {
    const mine = vendors.find(v => v && v.name === seed.name);
    if (!mine) continue;
    for (const key of ['period', 'currency']) {
      if (seed[key] !== undefined && mine[key] === undefined) { mine[key] = seed[key]; changed = true; }
    }
  }
  if (!Object.prototype.hasOwnProperty.call(live, 'locks')) {
    live.locks = SEED_LOCKS.map(l => ({ ...l }));
    changed = true;
  }
  return { rules: live, changed };
}

// A parsed file may be anything a person typed. Keep only the shape above.
function normalizeRules(raw) {
  const src = raw && typeof raw === 'object' ? raw : {};
  const list = Array.isArray(src.vendors) ? src.vendors : [];
  const vendors = [];
  for (const v of list) {
    if (!v || typeof v !== 'object') continue;
    const strings = (x) => (Array.isArray(x) ? x : (x == null ? [] : [x]))
      .map(s => String(s == null ? '' : s).trim()).filter(Boolean);
    const out = {
      name: String(v.name || '').trim(),
      domains: strings(v.domains).map(d => d.toLowerCase().replace(/^@/, '')),
      subject: strings(v.subject),
    };
    if (v.app) out.app = Array.isArray(v.app) ? strings(v.app) : String(v.app).trim();
    if (Array.isArray(v.apps)) {
      out.apps = v.apps
        .filter(a => a && typeof a === 'object' && a.app)
        .map(a => ({ app: Array.isArray(a.app) ? strings(a.app) : String(a.app).trim(), text: strings(a.text) }));
    }
    if (v.currency) out.currency = String(v.currency).trim().toUpperCase();
    // "usage": the vendor's line-item usage ranges are its billing period.
    if (v.period && /^usage$/i.test(String(v.period))) out.period = 'usage';
    // A vendor that names no row can never file anything; drop it rather than
    // let it swallow a domain and turn every invoice from it into a question.
    if (!out.app && !(out.apps && out.apps.length)) continue;
    if (!out.name) out.name = Array.isArray(out.app) ? out.app[0] : (out.app || (out.apps[0] && String(out.apps[0].app)));
    vendors.push(out);
  }
  // Cells the owner has set by hand and the sync must never rewrite:
  //   "locks": [{ "app": "Cursor pro", "month": "2026-07", "value": 1133.53, "note": "…" }]
  // A lock carries the figure the cell must hold. The sync writes that figure
  // if the cell differs and otherwise leaves the cell alone, whatever the
  // month's invoices come to — the owner's reasons (invoices missing from the
  // archive, charges paid from prepaid credits) are theirs to hold.
  const locks = [];
  for (const l of (Array.isArray(src.locks) ? src.locks : [])) {
    if (!l || typeof l !== 'object') continue;
    const app = String(l.app || '').trim();
    const month = String(l.month || '').trim();
    const value = l.value === '' || l.value === null || l.value === undefined ? null : Number(l.value);
    if (!app || !/^\d{4}-\d{2}$/.test(month) || value === null || !Number.isFinite(value)) continue;
    locks.push({ app, month, value, note: String(l.note || '').trim() });
  }
  return { version: 1, vendors, locks };
}

// The lock for one app-month, if any. Matched by the sheet's own normalized
// form, so "Cursor pro" and "Cursor Pro" are the same row.
function lockFor(rules, app, month) {
  return ((rules && rules.locks) || []).find(l => norm(l.app) === norm(app) && l.month === month) || null;
}

const domainOf = (addr) => {
  const s = String(addr || '').toLowerCase().trim();
  const at = s.lastIndexOf('@');
  return at === -1 ? '' : s.slice(at + 1);
};

// "mail.anthropic.com" is under "anthropic.com"; "notanthropic.com" is not.
const underDomain = (domain, rule) => domain === rule || domain.endsWith('.' + rule);

// A phrase counts only as whole words, so "Hex" does not fire on "hexadecimal"
// and "Render" does not fire on "surrender". Case does not matter.
function phraseIn(phrase, haystack) {
  const p = String(phrase || '').trim();
  if (!p) return false;
  const esc = p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`(^|[^A-Za-z0-9])${esc}(?=$|[^A-Za-z0-9])`, 'i').test(String(haystack || ''));
}

// Finds the vendor a message is from. Domain first — the strongest signal —
// then the subject words. Returns null when no rule fits.
function matchVendor(rules, signals) {
  const vendors = (rules && rules.vendors) || [];
  const domains = [signals.originalAddress, signals.address].map(domainOf).filter(Boolean);
  for (const domain of domains) {
    const hit = vendors.find(v => v.domains.some(d => underDomain(domain, d)));
    if (hit) return { vendor: hit, via: 'domain', matched: domain };
  }
  const head = [signals.subject, signals.senderName, ...(signals.attachmentNames || [])].filter(Boolean).join(' \n ');
  for (const v of vendors) {
    const word = v.subject.find(w => phraseIn(w, head));
    if (word) return { vendor: v, via: 'subject', matched: word };
  }
  return null;
}

// The whole decision for one PDF. `signals` is what the mail says about
// itself, `text` what the PDF says. Returns:
//   { app, vendor, via, confident: true }                   — file it
//   { app: null, vendor, reason, question, options, ... }   — hold and ask
function classify(rules, signals, text, appNames) {
  const found = matchVendor(rules, signals);
  if (!found) {
    return {
      app: null, vendor: null, reason: 'no-vendor-rule',
      question: 'Which app is this invoice for? (No rule names this sender yet — your answer will be remembered for it.)',
      options: [],
      domain: domainOf(signals.originalAddress) || domainOf(signals.address) || null,
    };
  }
  const { vendor, via } = found;

  if (vendor.app) {
    const app = resolveToSheetApp(vendor.app, appNames);
    if (app) return { app, vendor: vendor.name, via, confident: true, currency: vendor.currency || null, period: vendor.period || null };
    return {
      app: null, vendor: vendor.name, reason: 'app-not-in-sheet',
      question: `The rule for ${vendor.name} names "${Array.isArray(vendor.app) ? vendor.app.join('" or "') : vendor.app}", which is not a row in the sheet. Which row should it be?`,
      options: [],
    };
  }

  const candidates = vendor.apps || [];
  const body = [text, signals.subject, ...(signals.attachmentNames || [])].filter(Boolean).join('\n');
  for (const c of candidates) {
    const phrase = c.text.find(p => phraseIn(p, body));
    if (!phrase) continue;
    const app = resolveToSheetApp(c.app, appNames);
    if (app) return { app, vendor: vendor.name, via: `${via}+text`, matched: phrase, confident: true, currency: vendor.currency || null, period: vendor.period || null };
  }
  const options = candidates.map(c => resolveToSheetApp(c.app, appNames) || (Array.isArray(c.app) ? c.app[0] : c.app));
  return {
    app: null, vendor: vendor.name, reason: 'no-line-item-rule', period: vendor.period || null,
    question: options.length
      ? `${vendor.name} bills ${options.length} rows and nothing in this invoice matches a rule for any of them. Which is it?`
      : `${vendor.name} is known, but its rule names no app. Which row is this?`,
    options,
  };
}

// Writes an answer back into the rules so the same question is not asked
// twice. Two kinds:
//   · an unknown sender → a new vendor rule for its domain (or, for a sender
//     with no usable domain, its subject words)
//   · a vendor with several rows where nothing matched → a `text` phrase for
//     the chosen row: the invoice number's account prefix when the invoice
//     has one (Stripe's "Q8MUNTUC-0180" style), since that is the one stable
//     mark such invoices carry
// Returns the new rules and a note saying what was learned, or null when
// nothing safe could be inferred — then the file is filed, and the owner
// adds the rule by hand.
function learn(rules, item, app) {
  const out = normalizeRules(rules);
  if (!app) return { rules: out, learned: null };

  if (item.reason === 'no-vendor-rule') {
    const domain = String(item.domain || '').toLowerCase();
    const generic = /^(gmail|outlook|hotmail|yahoo|stripe|paddle|chargebee|quickbooks|intuit|xero|bill|zohomail|sarasanalytics)\./i;
    if (domain && !generic.test(domain) && !/\.(stripe|paddle)\.com$/.test(domain) && domain !== 'stripe.com') {
      out.vendors.push({ name: app, domains: [domain], subject: [], app });
      return { rules: out, learned: `sender domain ${domain} → ${app}` };
    }
    // A Stripe-sent invoice: the vendor's name is in the subject or sender name,
    // so remember that instead. Only a distinctive word, never the whole
    // subject ("Your receipt from …" would then match every receipt).
    const words = String(item.senderName || '').trim();
    if (words && words.length >= 3 && !/^(invoice|receipt|billing|payments?)$/i.test(words)) {
      out.vendors.push({ name: app, domains: [], subject: [words], app });
      return { rules: out, learned: `sender name "${words}" → ${app}` };
    }
    return { rules: out, learned: null };
  }

  if (item.reason === 'no-line-item-rule' || item.reason === 'app-not-in-sheet') {
    const vendor = out.vendors.find(v => v.name === item.vendor);
    if (!vendor) return { rules: out, learned: null };
    const prefix = accountPrefix(item.ref);
    if (item.reason === 'app-not-in-sheet' && vendor.app) {
      vendor.app = app;
      return { rules: out, learned: `${vendor.name} → ${app}` };
    }
    if (!prefix) return { rules: out, learned: null };
    vendor.apps = vendor.apps || [];
    let row = vendor.apps.find(a => resolveKey(a.app) === norm(app));
    if (!row) { row = { app, text: [] }; vendor.apps.push(row); }
    if (!row.text.includes(prefix)) row.text.push(prefix);
    return { rules: out, learned: `${vendor.name} invoices numbered ${prefix}-… → ${app}` };
  }
  return { rules: out, learned: null };
}

const resolveKey = (app) => norm(Array.isArray(app) ? app[0] : app);

// "Q8MUNTUC0180" (as extractInvoiceRef strips it) → "Q8MUNTUC". Only the
// Stripe shape — eight upper-case letters or digits followed by a run of
// digits — is treated as account-plus-sequence.
function accountPrefix(ref) {
  const m = String(ref || '').toUpperCase().match(/^([A-Z0-9]{8})(\d{4,})$/);
  return m && /[A-Z]/.test(m[1]) ? m[1] : null;
}

module.exports = { RULES_FILE, SEED_RULES, SEED_LOCKS, upgradeRules, normalizeRules, lockFor, matchVendor, classify, learn, accountPrefix, phraseIn, domainOf };
