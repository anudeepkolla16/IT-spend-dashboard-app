const { verify, parseCookies } = require('../../lib/session');

module.exports = async (req, res) => {
  const cookies = parseCookies(req.headers.cookie);
  const session = verify(cookies.session);
  res.setHeader('Cache-Control', 'no-store');
  if (!session) {
    res.status(200).json({ authenticated: false });
    return;
  }
  res.status(200).json({ authenticated: true, email: session.email, name: session.name });
};
