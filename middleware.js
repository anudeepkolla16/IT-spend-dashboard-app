import { createHmac, timingSafeEqual } from 'crypto';
import { next } from '@vercel/functions';

export const config = {
  matcher: '/:path*',
  runtime: 'nodejs', // needed for Node's crypto module used below
};

// Endpoints the report token may read (GET only). Keep this list tiny.
const REPORT_READ_PATHS = new Set(['/api/spend-data', '/api/invoices/list']);

function b64urlDecode(str) {
  return Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
}
function b64urlEncode(buf) {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function verifySession(cookieHeader, secret) {
  if (!cookieHeader || !secret) return null;
  const match = cookieHeader.match(/(?:^|;\s*)session=([^;]+)/);
  if (!match) return null;
  const [payload, sig] = match[1].split('.');
  if (!payload || !sig) return null;
  const expected = b64urlEncode(createHmac('sha256', secret).update(payload).digest());
  const a = Buffer.from(sig), b = Buffer.from(expected);
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
  try {
    const obj = JSON.parse(b64urlDecode(payload).toString('utf8'));
    if (!obj.exp || Date.now() > obj.exp) return null;
    return obj;
  } catch {
    return null;
  }
}

export default function middleware(request) {
  const url = new URL(request.url);

  // Let the sign-in flow itself through unauthenticated (that's the point of it).
  if (url.pathname.startsWith('/api/auth/')) return next();

  // The scheduled invoice job has no browser session; it authenticates itself
  // with CRON_SECRET inside the handler, so let it past the session gate.
  if (url.pathname === '/api/invoices/sync-cron') return next();

  // Read-only report access. The IT report agent (a GitHub Actions job) has no
  // browser session; it presents REPORT_TOKEN as a Bearer token and may read
  // only the two data endpoints — never the auth, import, upload or amount
  // routes, and never the Password page's `?sheet=logins` payload.
  if (request.method === 'GET' && REPORT_READ_PATHS.has(url.pathname)
      && url.searchParams.get('sheet') !== 'logins') {
    const want = process.env.REPORT_TOKEN || '';
    const got = (request.headers.get('authorization') || '').replace(/^Bearer\s+/i, '').trim();
    if (want && got && got.length === want.length
        && timingSafeEqual(Buffer.from(got), Buffer.from(want))) return next();
  }

  const secret = process.env.SESSION_SECRET;
  const session = verifySession(request.headers.get('cookie'), secret);
  if (session) return next(); // valid session, let the request through

  // API calls (e.g. the dashboard's own fetch()) get a JSON 401, not an HTML redirect,
  // so the frontend can show a clear error instead of trying to parse a login page as JSON.
  if (url.pathname.startsWith('/api/')) {
    return new Response(JSON.stringify({ error: 'Not authenticated' }), {
      status: 401,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  return Response.redirect(`${url.origin}/api/auth/login`, 302);
}
