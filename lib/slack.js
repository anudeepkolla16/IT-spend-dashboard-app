// Tells the sheet's owner what a run did, by Slack DM, and reads their answers.
//
// The owner asked for "a slack dm communication on changes made by you on daily
// basis or whenever you made changes", and to be able to answer the app's
// questions from Slack as well as from the dashboard. Both halves live here:
// a run that changed anything — filed, totalled, ticked, or held an invoice to
// ask about — posts one message; the next run reads whatever was typed back.
//
// Needs two env vars in Vercel, never in the repo:
//   SLACK_BOT_TOKEN   xoxb-… token of the workspace's own Slack app, with the
//                     scopes chat:write, im:write and im:history
//   SLACK_DM_USER     the member ID (U…) of the person to DM
// Without them every call here is a no-op that says so, and the run carries on:
// telling somebody what happened is never allowed to stop it happening.

const API = 'https://slack.com/api';

function config() {
  return {
    token: (process.env.SLACK_BOT_TOKEN || '').trim(),
    user: (process.env.SLACK_DM_USER || '').trim(),
  };
}

const configured = () => !!(config().token && config().user);

async function call(method, body) {
  const { token } = config();
  const res = await fetch(`${API}/${method}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body || {}),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok || !json.ok) throw new Error(`Slack ${method} failed: ${json.error || res.status}`);
  return json;
}

// The DM channel with the owner. Opening it is idempotent, so it is looked up
// every time rather than remembered — a remembered id survives the app being
// reinstalled, and then every post fails.
async function dmChannel() {
  const { user } = config();
  const json = await call('conversations.open', { users: user });
  return json.channel && json.channel.id;
}

// Posts one message. Slack caps a message at 40,000 characters and reads
// anything past ~4,000 badly, so a long report is split at line breaks.
const CHUNK = 3500;

async function postDm(text) {
  if (!configured()) return { sent: false, why: 'SLACK_BOT_TOKEN or SLACK_DM_USER is not set' };
  const channel = await dmChannel();
  const chunks = [];
  let buf = '';
  for (const line of String(text || '').split('\n')) {
    if (buf.length + line.length + 1 > CHUNK && buf) { chunks.push(buf); buf = ''; }
    buf += (buf ? '\n' : '') + line;
  }
  if (buf) chunks.push(buf);
  let last = null;
  for (const chunk of chunks) {
    last = await call('chat.postMessage', { channel, text: chunk, mrkdwn: true });
  }
  return { sent: true, ts: last && last.ts, chunks: chunks.length };
}

// Everything the owner typed into the DM since `oldestTs` (a Slack ts string),
// oldest first, without the app's own messages. Returns { messages, latestTs }.
async function readReplies(oldestTs) {
  if (!configured()) return { messages: [], latestTs: oldestTs || null, why: 'not configured' };
  const channel = await dmChannel();
  const params = { channel, limit: 200, inclusive: false };
  if (oldestTs) params.oldest = String(oldestTs);
  const json = await call('conversations.history', params);
  const messages = (json.messages || [])
    .filter(m => m && m.type === 'message' && !m.bot_id && !m.subtype && m.user === config().user)
    .map(m => ({ ts: m.ts, text: String(m.text || '') }))
    .sort((a, b) => Number(a.ts) - Number(b.ts));
  const all = (json.messages || []).map(m => Number(m.ts)).filter(n => !Number.isNaN(n));
  const latestTs = all.length ? String(Math.max(...all, Number(oldestTs || 0))) : (oldestTs || null);
  return { messages, latestTs };
}

module.exports = { postDm, readReplies, configured };
