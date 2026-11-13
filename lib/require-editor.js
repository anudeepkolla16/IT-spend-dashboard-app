const { verify, parseCookies, canEdit, editorList } = require('./session');

// The one gate every writing route goes through.
//
// `canEdit` used to be checked in exactly one handler — the cell editor — and
// nowhere else. The dashboard hid the other controls from people who are not
// on the editor list, and that hiding was mistaken for enforcement. It is not:
// the buttons post to ordinary endpoints, and anyone signed in could reach
// them from the console, from a page left open since before they were moved
// off the list, or from curl. Signed in meant allowed to write.
//
// What that opened, for every reader with an account: running the mailbox sync
// and the archive tidy (which move files in SharePoint), applying a statement's
// amounts, repairing totals, and importing invoices from a folder. All of them
// write to the workbook that is the system of record.
//
// So the check lives here, once, and the routes call it. The page's own
// `admin-only` class is a courtesy to people who cannot use a control, never
// the thing that stops them.
//
// Returns the session when the caller may write. Otherwise it has already
// answered 401 or 403, and the handler must return without doing anything.
function requireEditor(req, res) {
  let session = null;
  try {
    session = verify(parseCookies(req.headers.cookie).session);
  } catch (e) {
    session = null;
  }
  if (!session) {
    res.status(401).json({ error: 'Not authenticated' });
    return null;
  }
  if (!canEdit(session.email)) {
    res.status(403).json({
      error: editorList().length
        ? 'Your account may read the dashboard but not change it. Ask the sheet owner to add you to EDITOR_EMAILS.'
        : 'Editing is not configured: set EDITOR_EMAILS in Vercel to the accounts allowed to write to the sheet.',
    });
    return null;
  }
  return session;
}

module.exports = { requireEditor };
