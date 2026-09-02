// Run with: node --test
//
// index.html carries the whole dashboard in one inline script. A stray quote
// in it stops the page from rendering at all — nothing loads, no error is
// shown, and no other test notices, because none of them parse the page's
// script. This one does: every <script> block must at least compile.

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

test('every inline script in index.html parses', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
  assert.ok(blocks.length >= 1, 'expected at least one inline script');
  for (const [i, m] of blocks.entries()) {
    const line = html.slice(0, m.index).split('\n').length;
    assert.doesNotThrow(
      () => new vm.Script(m[1], { filename: `index.html (script ${i + 1}, from line ${line})` }),
      `inline script ${i + 1} (starting at line ${line}) does not parse`
    );
  }
});
