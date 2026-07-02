module.exports = async (req, res) => {
  res.setHeader('Set-Cookie', 'session=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0');
  res.redirect(302, '/api/auth/login');
};
