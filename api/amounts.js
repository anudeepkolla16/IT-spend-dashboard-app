const preview = require('../lib/amounts/preview');
const apply = require('../lib/amounts/apply');
const log = require('../lib/amounts/log');

// One function for the whole amount-import flow, dispatching on `action`.
// Vercel's Hobby plan caps a deployment at 12 Serverless Functions and each
// file under api/ counts as one, so these three share a route rather than
// spending three of the budget. The handlers themselves live in lib/amounts/.
const ROUTES = { preview, apply, log };

module.exports = async (req, res) => {
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
