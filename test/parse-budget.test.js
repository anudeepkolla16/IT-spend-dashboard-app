// Run with: node --test
//
// The run has one deadline, and two phases share it: parsing the month's PDFs,
// then writing — an Excel session, the tracker, every cell, the audit log, the
// index. The parse loop used to be allowed to run right up to the deadline,
// which went unnoticed for as long as almost every PDF came from the cache and
// parsing took a second or two.
//
// Then PARSE_VERSION went to 3 and invalidated all 107 cached reads at once.
// Parsing ran its full 45 seconds, the writes began from there, and the
// function passed Vercel's 60-second ceiling: HTTP 504, and the run was lost
// along with everything it had just parsed.
//
// So: the parse budget always stops short of the run's deadline.

const test = require('node:test');
const assert = require('node:assert');

const { parseBudget, PARSE_MAX_PER_RUN } = require('../lib/mail-sync');

const NOW = 1_700_000_000_000;

test('parsing stops early enough to leave the writes room', () => {
  const deadline = NOW + 45 * 1000;          // what the mail run is given
  const reserve = deadline - parseBudget(deadline, NOW).deadline;
  assert.ok(reserve >= 15 * 1000,
    `only ${reserve}ms left for the writes — an Excel session, the tracker, the cells, the log and the index`);
});

test('a short run still gets to parse something', () => {
  // The nightly pass is given 25 seconds, not 45. A flat 20-second reserve
  // would leave it 5 and it would total nothing, run after run.
  const deadline = NOW + 25 * 1000;
  const b = parseBudget(deadline, NOW);
  assert.ok(b.deadline - NOW >= 10 * 1000, 'the cron must still be able to read a few invoices');
  assert.ok(deadline - b.deadline >= 8 * 1000, 'and still leave the writes something');
});

test('the parse budget never outlives the run', () => {
  for (const ms of [0, 1000, 5000, 20_000, 45_000, 600_000]) {
    const deadline = NOW + ms;
    const b = parseBudget(deadline, NOW);
    assert.ok(b.deadline <= deadline, `${ms}ms run: budget must not outlast the deadline`);
    assert.ok(b.deadline >= NOW, `${ms}ms run: a budget behind the clock is not a budget`);
  }
});

test('a run with no time left parses nothing rather than overrunning', () => {
  // Better to write nothing this run — the months whose files are unread are
  // held, never written from a partial total — than to die at the hard limit.
  for (const deadline of [NOW, NOW - 5000]) {
    const b = parseBudget(deadline, NOW);
    assert.strictEqual(b.deadline, NOW, 'no parsing time, so the loop stops at once');
    assert.strictEqual(b.parsed, 0);
  }
});

test('a sweep is spread over runs rather than crammed into one', () => {
  // 40 downloads-and-parses in one run is what ran the clock out. Fewer per
  // run, and the next run continues: each one persists what it parsed, stamped
  // with the current PARSE_VERSION.
  assert.ok(PARSE_MAX_PER_RUN <= 20, `${PARSE_MAX_PER_RUN} files in one run is what caused the 504`);
  assert.strictEqual(parseBudget(NOW + 45_000, NOW).maxParse, PARSE_MAX_PER_RUN);
});

test('the budget starts empty and not already exhausted', () => {
  const b = parseBudget(NOW + 45_000, NOW);
  assert.deepStrictEqual(b.fresh, []);
  assert.strictEqual(b.exhausted, false);
});
