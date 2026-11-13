const preview = require('../lib/amounts/preview');
const apply = require('../lib/amounts/apply');
const log = require('../lib/amounts/log');
const fixTotals = require('../lib/amounts/fix-totals');
const edit = require('../lib/amounts/edit');
const activity = require('../lib/amounts/activity');
const { requireEditor } = require('../lib/require-editor');

// One function for the whole amount-import flow, dispatching on `action`.
// Vercel's Hobby plan caps a deployment at 12 Serverless Functions and each
// file under api/ counts as one, so these three share a route rather than
// spending three of the budget. The handlers themselves live in lib/amounts/.
const ROUTES = { preview, apply, log, 'fix-totals': fixTotals, edit, activity };

module.exports = async (req, res) => {
  // Every action on this route writes to the workbook or reads the audit
  // trail, so the whole route is for editors. The dashboard's read-only views
  // never call it.
  if (!requireEditor(req, res)) return;

  const action = String(
    (req.body && req.body.action) || (req.query && req.query.action) || ''
  ).trim();

  const handler = ROUTES[action];
  if (!handler) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(400).json({ error: `Unknown action "${action}". Expected one of: ${Object.keys(ROUTES).join(', ')}.` });
    return;
  }
  return handler(req, res);
};
