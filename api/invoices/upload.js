const { getGraphToken, resolveDriveId, encodeGraphPath, sanitizeSegment } = require('../../lib/graph');

// Microsoft Graph's simple (single-request) upload endpoint tops out at 4 MiB.
// Capped lower here because base64-encoding the file for the JSON request body
// inflates its size ~33%, and Vercel's request body limit is ~4.5 MB.
const MAX_BYTES = 3 * 1024 * 1024;

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') {
      res.status(405).json({ error: 'Method not allowed' });
      return;
    }
    const { app, filename, contentBase64 } = req.body || {};
    if (!app || !filename || !contentBase64) {
      res.status(400).json({ error: 'Missing app, filename, or contentBase64' });
      return;
    }
    if (!/\.pdf$/i.test(filename)) {
      res.status(400).json({ error: 'Only PDF files are supported' });
      return;
    }

    const buffer = Buffer.from(contentBase64, 'base64');
    if (!buffer.length) {
      res.status(400).json({ error: 'Empty file' });
      return;
    }
    if (buffer.length > MAX_BYTES) {
      res.status(413).json({ error: 'File too large (max 3 MB)' });
      return;
    }

    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) throw new Error('Missing TARGET_USER_UPN env var');

    const baseName = sanitizeSegment(filename.replace(/\.pdf$/i, ''));
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const path = `Invoices/${sanitizeSegment(app)}/${stamp}_${baseName}.pdf`;

    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(path)}:/content`;

    const gres = await fetch(url, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/pdf' },
      body: buffer,
    });
    if (!gres.ok) {
      const text = await gres.text();
      throw new Error(`Graph upload failed (${gres.status}): ${text.slice(0, 300)}`);
    }
    const json = await gres.json();
    res.status(200).json({ ok: true, name: json.name, webUrl: json.webUrl });
  } catch (err) {
    res.status(502).json({ error: err.message || String(err) });
  }
};
