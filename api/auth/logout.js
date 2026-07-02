module.exports = async (req, res) => {
  res.setHeader('Set-Cookie', 'session=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0');
  res.writeHead(302, { Location: '/api/auth/login' });
  res.end();
};
