const {
  resolveArchiveRoot, graphListAll, moveItem, encodeGraphPath, itemIdByPath,
} = require('../graph');
const { parseMonthFolder } = require('./inventory');

// Tidies the shape of the invoice archive, which drifts as people file by hand.
//
// It fixes exactly one thing: a month folder buried a level deeper than
// {vendor}/{month}. That happens when a whole vendor folder is dragged inside
// another — "Luzmo" ended up inside "Cumul(Luzmo)", leaving its July invoices at
// Cumul(Luzmo)/Luzmo/July/. The sync totals {vendor}/{month}/ and nothing below
// it, so those three invoices counted towards no month at all.
//
// Everything else it only reports. It never deletes, never overwrites, and never
// decides which of two copies of an invoice is the real one.

const children = (token, driveId, itemId) => graphListAll(
  token,
  `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$select=id,name,folder,file,size&$top=200`,
  'Graph folder listing'
);

async function planTidy(token, driveId, opts) {
  const options = opts || {};
  const root = await resolveArchiveRoot(token, driveId, { fresh: true });
  if (!root.resolved) {
    return { root: null, triedPaths: root.candidates, moves: [], duplicates: [], empties: [], notes: [] };
  }

  const moves = [];
  const duplicates = [];
  const empties = [];
  const notes = [];

  const vendors = (await children(token, driveId, root.itemId))
    .filter(c => c.folder && !String(c.name).startsWith('_'));

  for (const vendor of vendors) {
    if (options.deadline && Date.now() > options.deadline) { notes.push('Ran out of time — run it again to finish.'); break; }

    let kids;
    try { kids = await children(token, driveId, vendor.id); }
    catch (e) { notes.push(`${vendor.name}: could not be listed (${e.message})`); continue; }

    const monthFolders = new Map();      // month key -> folder at {vendor}/{month}
    const strays = [];                   // subfolders that are not months
    for (const k of kids) {
      if (!k.folder) continue;
      const month = parseMonthFolder(k.name, null);
      if (month) monthFolders.set(month, k);
      else strays.push(k);
    }
    if (!strays.length) continue;

    // Every file name already somewhere under this vendor, so a name that exists
    // twice is never moved on top of itself or counted into a second month.
    const namesSeen = new Map();         // lowercased name -> where it was first seen
    const noteName = (name, where) => {
      const key = String(name).toLowerCase();
      if (!namesSeen.has(key)) namesSeen.set(key, where);
      return namesSeen.get(key);
    };
    for (const [month, folder] of monthFolders) {
      let inMonth;
      try { inMonth = await children(token, driveId, folder.id); }
      catch (e) { notes.push(`${vendor.name}/${folder.name}: could not be listed (${e.message})`); continue; }
      for (const f of inMonth) if (f.file) noteName(f.name, `${vendor.name}/${folder.name}`);
    }
    for (const k of kids) if (k.file) noteName(k.name, vendor.name);

    for (const stray of strays) {
      let inStray;
      try { inStray = await children(token, driveId, stray.id); }
      catch (e) { notes.push(`${vendor.name}/${stray.name}: could not be listed (${e.message})`); continue; }

      for (const sub of inStray) {
        if (!sub.folder) continue;
        const month = parseMonthFolder(sub.name, null);
        if (!month) continue;            // not a month: not this job's business

        let files;
        try { files = (await children(token, driveId, sub.id)).filter(f => f.file); }
        catch (e) { notes.push(`${vendor.name}/${stray.name}/${sub.name}: could not be listed (${e.message})`); continue; }

        if (!files.length) {
          empties.push({ path: `${vendor.name}/${stray.name}/${sub.name}`, why: 'empty' });
          continue;
        }

        const target = monthFolders.get(month);
        const from = `${vendor.name}/${stray.name}/${sub.name}`;

        // Nothing of that month at the vendor root: the whole folder can go up
        // in one operation, keeping its name and its item ids.
        if (!target) {
          const clash = files.map(f => f.name).filter(n => namesSeen.has(String(n).toLowerCase()));
          if (clash.length) {
            duplicates.push({ vendor: vendor.name, from, month, files: clash, alreadyAt: clash.map(n => namesSeen.get(String(n).toLowerCase())) });
            continue;
          }
          moves.push({
            kind: 'folder', vendor: vendor.name, month, itemId: sub.id,
            from, to: `${vendor.name}/${sub.name}`,
            destPath: `${root.path}/${vendor.name}`, files: files.map(f => f.name),
          });
          files.forEach(f => noteName(f.name, `${vendor.name}/${sub.name}`));
          continue;
        }

        // The month already exists up top, so the files move into it one by one.
        for (const f of files) {
          const seenAt = namesSeen.get(String(f.name).toLowerCase());
          if (seenAt) {
            // The same invoice is already filed under this vendor. Moving this
            // copy up would put the identical charge in a second month, which is
            // exactly the double count the totals are built to avoid.
            duplicates.push({ vendor: vendor.name, from, month, files: [f.name], alreadyAt: [seenAt] });
            continue;
          }
          moves.push({
            kind: 'file', vendor: vendor.name, month, itemId: f.id, name: f.name,
            from: `${from}/${f.name}`, to: `${vendor.name}/${target.name}/${f.name}`,
            destPath: `${root.path}/${vendor.name}/${target.name}`,
          });
          noteName(f.name, `${vendor.name}/${target.name}`);
        }
      }
    }
  }

  return { root: root.path, moves, duplicates, empties, notes };
}

// Executes only the moves the plan proposed and the caller sent back. Every
// destination is re-resolved here rather than trusted from the request, so an
// edited payload cannot move a file somewhere outside the archive.
async function applyTidy(token, driveId, requested) {
  const root = await resolveArchiveRoot(token, driveId, { fresh: true });
  if (!root.resolved) throw new Error('The invoice archive could not be found, so nothing was moved.');

  const done = [];
  const failed = [];
  for (const move of Array.isArray(requested) ? requested : []) {
    if (!move || !move.itemId || !move.destPath) { failed.push({ move, error: 'incomplete' }); continue; }
    const destPath = String(move.destPath);
    // A prefix test alone is not containment: "…/Invoices/../elsewhere" starts
    // with the archive path and still lands outside it. Reject traversal on the
    // segments first, and never lean on Graph to normalise it away.
    const segments = destPath.split('/');
    const traverses = segments.some(seg => seg === '' || seg === '.' || seg === '..');
    const contained = destPath === root.path || destPath.startsWith(`${root.path}/`);
    if (traverses || !contained) {
      failed.push({ move, error: `destination "${destPath}" is outside the archive` });
      continue;
    }
    try {
      const parentId = await itemIdByPath(token, driveId, destPath);
      if (!parentId) throw new Error(`destination "${destPath}" no longer exists`);
      await moveItem(token, driveId, move.itemId, parentId);
      done.push({ from: move.from, to: move.to, kind: move.kind });
    } catch (e) {
      failed.push({ move: { from: move.from, to: move.to }, error: e.message });
    }
  }
  return { root: root.path, moved: done, failed };
}

module.exports = { planTidy, applyTidy };
