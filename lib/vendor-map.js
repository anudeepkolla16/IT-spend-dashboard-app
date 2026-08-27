// Maps a statement's vendor label onto an app row in the spend sheet.
//
// Three sources, in order of confidence:
//   1. the saved alias map (Invoices/_amount-map.json) — anything confirmed in
//      the review UI is remembered, so each new label is resolved by hand once
//   2. the seed aliases below, taken from labels already present in real
//      statements, so the first run has few exceptions
//   3. the bank descriptor, for the sheet that has no Comments column
// Anything still unresolved becomes an exception for the user to map.

const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
const tokens = (s) => String(s == null ? '' : s).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);

// Vendor label (normalized) -> app name as written in the sheet. Values are
// resolved against the live sheet by normalized comparison, so trailing spaces
// in the sheet's own names ("Render ", "LastPass ") do not matter.
const SEED_ALIASES = {
  adobe: 'Adobe',
  amazonwebservices: 'AWS',
  aws: 'AWS',
  anthropic: 'Anthropic(Api Console)',
  anthropicapiconsole: 'Anthropic(Api Console)',
  claudeapi: 'Anthropic(Api Console)', // the archive folder is named for the vendor, not the row
  apify: 'Apify',
  apollo: 'Apollo',
  bitly: 'Bitly',
  bubblestarter: 'Bubble Starter',
  bubble: 'Bubble Starter',
  canva: 'CANVA',
  chargebee: 'Chargebee',
  cumul: 'Cumul(Luzmo)',
  luzmo: 'Cumul(Luzmo)',
  cursor: 'Cursor pro',
  dbtcloud: 'DBT Cloud',
  dbt: 'DBT Cloud',
  dovetail: 'Dovetail',
  elevenlabs: 'ElevenLabs',
  envato: 'Envato',
  figma: 'FIGMA',
  filestack: 'Filestack',
  github: 'Github',
  godaddy: 'Godaddy',
  googleads: 'GOOGLE ADS',
  google: 'GOOGLE ADS',
  googledes: 'Google Workspace',
  googleworkspace: 'Google Workspace',
  googlecloud: 'Google cloud',
  googlevoice: 'Google Voice',
  granola: 'Granola Business',
  helpjuice: 'Helpjuice',
  hex: 'Hex',
  hextech: 'Hex',
  hubspot: 'Hubspot',
  jetbrains: 'JetBrains',
  keka: 'Keka',
  keepa: 'Keepa',
  laptopprocurment: 'Laptops Procurement', // the folder carries a long-standing typo
  lastpass: 'LastPass',
  linkedin: 'Linkedin',
  likedin: 'Linkedin', // recurring typo in the Comments column
  lottiefiles: 'LOTTIEFILES',
  metalprice: 'MetalPriceAPI',
  metalpriceapi: 'MetalPriceAPI',
  microsoft: 'MICROSOFT(Tata Tele)',
  clickup: 'Mango technology(Clickup)',
  mangotechnology: 'Mango technology(Clickup)',
  naukri: 'Naukri',
  openai: 'OPENAI',
  pagerduty: 'Pagerduty',
  phantombuster: 'Phantombuster',
  productfruits: 'Product Fruits',
  prosp: 'Prosp AI',
  prospai: 'Prosp AI',
  posthog: 'Posthog',
  render: 'Render',
  sentry: 'Sentry.io',
  sentryio: 'Sentry.io',
  slack: 'SLACK',
  sprinto: 'Sprinto',
  superhuman: 'Superhuman',
  shopify: 'Shopify Inc',
  tacsecurity: 'TAC Security',
  tmobile: 'TMobile',
  twilio: 'Twilo Sendgrid',
  twilosendgrid: 'Twilo Sendgrid',
  sendgrid: 'Twilo Sendgrid',
  typeform: 'Typeform',
  vimcal: 'vimcal',
  webflow: 'WEBFLOW',
  wingman: 'WINGMAN(Clari)',
  clari: 'WINGMAN(Clari)',
  windsurf: 'windsurf pro',
  xcorp: 'X corp',
  zohobooks: 'ZOHO Books',
  zoho: 'ZOHO Books',
  zoom: 'ZOOM',
};

// For the statement sheet with no Comments column, the app has to come from the
// raw bank descriptor. Most specific patterns first — "ANTHROPIC* CLAUDE SUB" is
// the Claude seat subscription, which is a different row from the API console.
const DESCRIPTOR_RULES = [
  // Charges billed through a reseller name the reseller first, so the actual
  // vendor has to win: "Google ChatGPT" is an OpenAI charge, not Google Ads.
  [/chat\s*gpt|open\s*ai/i, 'OPENAI'],
  [/anthropic\W*\s*claude\s*sub/i, 'Claude Ai Max 6 Accounts'],
  [/claude\s*(ai|sub)/i, 'Claude Ai'],
  [/google\s*\*?\s*cloud/i, 'Google cloud'],
  [/google\s*\*?\s*ads/i, 'GOOGLE ADS'],
  [/google\s+linkedin/i, 'Linkedin'],
  [/linkedin/i, 'Linkedin'],
  [/webflow/i, 'WEBFLOW'],
  [/cursor/i, 'Cursor pro'],
  [/dbt\s*cloud/i, 'DBT Cloud'],
  [/bubble/i, 'Bubble Starter'],
  [/amazon\s*web\s*servi/i, 'AWS'],
  [/anthropic/i, 'Anthropic(Api Console)'],
  [/windsurf/i, 'windsurf pro'],
  [/t-?mobile/i, 'TMobile'],
  [/twilio|sendgrid/i, 'Twilo Sendgrid'],
  [/zoho/i, 'ZOHO Books'],
  [/sentry/i, 'Sentry.io'],
  [/granola/i, 'Granola Business'],
  [/helpjuice/i, 'Helpjuice'],
  [/pagerduty/i, 'Pagerduty'],
  [/phantombuster/i, 'Phantombuster'],
  [/product\s*fruits/i, 'Product Fruits'],
  [/metalprice/i, 'MetalPriceAPI'],
  [/filestack/i, 'Filestack'],
  [/elevenlabs/i, 'ElevenLabs'],
  [/superhuman/i, 'Superhuman'],
  [/typeform/i, 'Typeform'],
  [/chargebee/i, 'Chargebee'],
  [/hex\.tech/i, 'Hex'],
  [/github/i, 'Github'],
  [/envato/i, 'Envato'],
  [/apify/i, 'Apify'],
  [/adobe/i, 'Adobe'],
  [/render\.com/i, 'Render'],
  [/prosp/i, 'Prosp AI'],
];

// Fuzzy match, used only to pre-select a suggestion in the review UI — the user
// confirms or overrides it before anything is written.
function bestMatch(label, appNames) {
  const na = norm(label);
  if (!na) return { app: null, score: 0 };
  const ta = new Set(tokens(label));
  let best = null, bestScore = 0;
  for (const app of appNames) {
    const nb = norm(app);
    if (!nb) continue;
    let score = 0;
    if (na === nb) score = 1;
    else if (na.includes(nb) || nb.includes(na)) score = 0.8;
    else {
      const tb = new Set(tokens(app));
      const inter = [...ta].filter(x => tb.has(x)).length;
      const union = new Set([...ta, ...tb]).size;
      score = union ? inter / union : 0;
      for (const t of tb) if (t.length >= 4 && na.includes(t)) score = Math.max(score, 0.6);
      for (const t of ta) if (t.length >= 4 && nb.includes(t)) score = Math.max(score, 0.6);
    }
    if (score > bestScore) { bestScore = score; best = app; }
  }
  return { app: best, score: bestScore };
}

// Turn an alias's target into the app name exactly as the sheet spells it.
function resolveToSheetApp(candidate, appNames) {
  if (!candidate) return null;
  const nc = norm(candidate);
  return appNames.find(a => norm(a) === nc) || null;
}

function buildResolver(savedAliases, appNames) {
  const aliases = { ...SEED_ALIASES, ...(savedAliases || {}) };

  return function resolve(vendorLabel, description) {
    const key = norm(vendorLabel);

    // 1. Saved or seeded alias for the hand-typed label.
    if (key && aliases[key]) {
      const app = resolveToSheetApp(aliases[key], appNames);
      if (app) return { app, via: savedAliases && savedAliases[key] ? 'saved-alias' : 'seed-alias', confident: true };
    }

    // 2. The label already spells an app name exactly.
    if (key) {
      const direct = resolveToSheetApp(vendorLabel, appNames);
      if (direct) return { app: direct, via: 'exact-name', confident: true };
    }

    // 3. No label (or an unknown one) — read the bank descriptor.
    if (description) {
      for (const [re, target] of DESCRIPTOR_RULES) {
        if (re.test(description)) {
          const app = resolveToSheetApp(target, appNames);
          if (app) return { app, via: 'descriptor', confident: true };
        }
      }
      const dkey = norm(description);
      for (const alias of Object.keys(aliases)) {
        if (alias.length >= 5 && dkey.includes(alias)) {
          const app = resolveToSheetApp(aliases[alias], appNames);
          if (app) return { app, via: 'descriptor-alias', confident: true };
        }
      }
    }

    // 4. Nothing matched — offer the closest name, but never treat it as settled.
    const guess = bestMatch(vendorLabel || description, appNames);
    return { app: null, suggestion: guess.score >= 0.4 ? guess.app : null, score: guess.score, via: 'unmapped', confident: false };
  };
}

// Some vendors bill one company across several rows of the sheet, and the
// sender, subject and filename are identical for all of them — only the invoice
// body tells them apart. Anthropic is the case in hand: the API console and the
// Claude seat subscriptions all arrive from Anthropic, PBC.
//
// Verified against the two July 2026 invoices in the archive:
//   API console  — "Claude Opus 4.8 Usage Jul 2 - Jul 31", total 13,479.42,
//                  which is exactly what the sheet holds for Jul-26.
//   Claude seats — "Extra usage units, Enterprise plan", "Team plan",
//                  0.00 paid against an applied balance, which is also what the
//                  sheet holds.
//
// The amount threshold the sheet owner described (API console above 10k, Claude
// below) does not hold historically — Claude Ai was 14,160 in Feb-26 and the API
// console 371.88 in Jan-26 — so it is only the last resort here, after the far
// more reliable wording of the line items.
const ANTHROPIC_APPS = { api: 'Anthropic(Api Console)', seats: 'Claude Ai' };

function refineAnthropic(text, appNames, threshold) {
  const t = String(text || '');
  if (!/anthropic/i.test(t)) return null;

  const seatWords = /\b(seat|seats|accounts?|enterprise plan|team plan|max plan|per user|users)\b/i;
  const usageLines = /claude\s+[a-z]+\s*[\d.]*\s+usage\b/i;

  let target = null;
  if (seatWords.test(t)) target = ANTHROPIC_APPS.seats;
  else if (usageLines.test(t)) target = ANTHROPIC_APPS.api;
  else {
    // Nothing in the wording settles it — fall back to the stated threshold.
    const m = t.match(/\bamount\s+due\b[^0-9\-]{0,40}\$?\s*([\d.,]+)/i) || t.match(/\btotal\b[^0-9\-]{0,20}\$?\s*([\d.,]+)/i);
    const amt = m ? parseFloat(String(m[1]).replace(/,/g, '')) : null;
    if (amt === null || Number.isNaN(amt)) return null;
    target = amt > (threshold || 10000) ? ANTHROPIC_APPS.api : ANTHROPIC_APPS.seats;
  }
  return resolveToSheetApp(target, appNames);
}

module.exports = { norm, bestMatch, buildResolver, resolveToSheetApp, refineAnthropic, SEED_ALIASES, DESCRIPTOR_RULES };
