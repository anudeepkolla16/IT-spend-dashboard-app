const { resolveArchiveRoot, archiveFile, readJsonFile, writeJsonFile } = require('./graph');

// Who used the dashboard, and what they did with it.
//
// Two sources, because one of them already existed and duplicating it would
// mean two records of the same write that could disagree:
//
//   _amount-log.json    every cell this app has written, with who, when, and
//                       what the cell held before. Already kept by the sheet
//                       writer; nothing here writes to it.
//   _activity-log.json  everything else a person does that leaves no cell
//                       behind — signing in, running the mailbox sync or the
//                       archive tidy, importing a folder.
//
// The reader merges the two into one feed. The writer below only ever touches
// the second file.
//
// Recording is instrumentation, never a step in the work: a run that cannot
// write its own log entry has still done the thing it was asked to do, and a
// person whose sign-in could not be recorded is still signed in. So every
// function here swallows its errors and returns rather than throwing into a
// caller that was busy doing something that matters.

const ACTIVITY_FILE = '_activity-log.json';

// Enough to answer "who has been in here lately" without making the file
// expensive to read on every open. A month of ordinary use is a few hundred.
const MAX_ENTRIES = 500;

const activityPath = (root) => archiveFile(root, ACTIVITY_FILE);

function clean(value, max) {
  const s = value == null ? '' : String(value);
  return s.length > max ? s.slice(0, max - 1) + '…' : s;
}

// One event. `kind` is what happened, `what` is the one-line summary a person
// reads, `detail` is the longer line under it.
function entryOf(kind, by, what, detail) {
  return {
    at: new Date().toISOString(),
    kind: clean(kind, 40),
    by: clean(by, 200).toLowerCase(),
    what: clean(what, 300),
    detail: clean(detail, 1000),
  };
}

async function readActivity(token, driveId) {
  try {
    const root = await resolveArchiveRoot(token, driveId);
    const file = await readJsonFile(token, driveId, activityPath(root));
    const entries = file && Array.isArray(file.entries) ? file.entries : [];
    return entries;
  } catch (e) {
    return [];
  }
}

// Appends one event. Never throws, and never reports failure: the caller is in
// the middle of signing someone in or finishing a sync, and neither should end
// differently because a log file could not be written.
async function recordActivity(token, driveId, kind, by, what, detail) {
  try {
    const root = await resolveArchiveRoot(token, driveId);
    const existing = await readActivity(token, driveId);
    const entries = existing.concat([entryOf(kind, by, what, detail)]).slice(-MAX_ENTRIES);
    await writeJsonFile(token, driveId, activityPath(root), {
      entries,
      updatedAt: new Date().toISOString(),
    });
    return true;
  } catch (e) {
    return false;
  }
}

// The same, for callers that hold no Graph token — the sign-in handler, which
// runs before anything else has reason to open one. Resolving the drive costs
// two round trips, which is why it is only done where there is no alternative.
async function recordActivityStandalone(kind, by, what, detail) {
  try {
    const { getGraphToken, resolveDriveId } = require('./graph');
    const upn = (process.env.TARGET_USER_UPN || '').trim();
    if (!upn) return false;
    const token = await getGraphToken();
    const driveId = await resolveDriveId(token, upn);
    return await recordActivity(token, driveId, kind, by, what, detail);
  } catch (e) {
    return false;
  }
}

module.exports = {
  ACTIVITY_FILE, MAX_ENTRIES, activityPath, entryOf,
  readActivity, recordActivity, recordActivityStandalone,
};
