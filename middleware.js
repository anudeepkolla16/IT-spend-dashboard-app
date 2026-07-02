export const config = {
  matcher: '/:path*',
};

export default function middleware(request) {
  const expectedUser = process.env.DASHBOARD_USER;
  const expectedPass = process.env.DASHBOARD_PASS;

  // Fail closed: if credentials aren't configured, block everything rather than serve openly.
  if (!expectedUser || !expectedPass) {
    return new Response('Dashboard is not yet configured (missing DASHBOARD_USER/DASHBOARD_PASS).', { status: 503 });
  }

  const authHeader = request.headers.get('authorization') || '';
  if (authHeader.startsWith('Basic ')) {
    try {
      const decoded = atob(authHeader.slice(6));
      const sep = decoded.indexOf(':');
      const user = decoded.slice(0, sep);
      const pass = decoded.slice(sep + 1);
      if (user === expectedUser && pass === expectedPass) {
        return; // authorized, let the request through
      }
    } catch (e) {
      // fall through to 401
    }
  }

  return new Response('Authentication required', {
    status: 401,
    headers: { 'WWW-Authenticate': 'Basic realm="Saras Spend Dashboard"' },
  });
}
