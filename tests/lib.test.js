const test = require('node:test');
const assert = require('node:assert');
const {
  splitPools,
  parseActivityTime,
  parseExtractedAt,
  computeStaleness,
  dedupeLatest
} = require('../Lib.js');

test('splitPools', (t) => {
  assert.deepStrictEqual(splitPools(""), []);
  assert.deepStrictEqual(splitPools("101"), ["101"]);
  assert.deepStrictEqual(splitPools("101, 102 , 103"), ["101", "102", "103"]);
  assert.deepStrictEqual(splitPools("101,101,102,, "), ["101", "102"]);
});

test('parseActivityTime', (t) => {
  assert.strictEqual(parseActivityTime("2026-09-20 00:03:38"), "2026-09-20T00:03:38+08:00");
  assert.throws(() => parseActivityTime("2026-09-20"), /Unparseable activity time/);
});

test('parseExtractedAt', (t) => {
  assert.strictEqual(parseExtractedAt("Agent Report 2026-09-21 16:21:28"), "2026-09-21T16:21:28+08:00");
  const fallback = new Date(Date.UTC(2026, 8, 21, 8, 21, 28)); // UTC
  assert.strictEqual(parseExtractedAt("Random File Name", fallback), "2026-09-21T16:21:28+08:00");
});

test('computeStaleness', (t) => {
  const dataEndIso = "2026-09-20T04:00:00+08:00";
  // Same day (Sept 20, 10:00:00 PM PHT)
  let nowMs = new Date("2026-09-20T22:00:00+08:00").getTime();
  let result = computeStaleness(dataEndIso, nowMs);
  assert.strictEqual(result.staleDays, 0);
  assert.strictEqual(result.isStale, false);

  // Next day (Sept 21, 01:00:00 AM PHT)
  nowMs = new Date("2026-09-21T01:00:00+08:00").getTime();
  result = computeStaleness(dataEndIso, nowMs);
  assert.strictEqual(result.staleDays, 1);
  assert.strictEqual(result.isStale, false);

  // Two days later (Sept 22, 01:00:00 AM PHT) - STALE
  nowMs = new Date("2026-09-22T01:00:00+08:00").getTime();
  result = computeStaleness(dataEndIso, nowMs);
  assert.strictEqual(result.staleDays, 2);
  assert.strictEqual(result.isStale, true);
});

test('dedupeLatest', (t) => {
  const { headers, rows } = require('../fixtures/fake_report_data.json');

  const { agents: results, globalDataStartMs } = dedupeLatest(rows, headers);
  assert.strictEqual(results.length, 3);
  assert.strictEqual(globalDataStartMs, new Date("2026-09-20T00:00:00+08:00").getTime());

  const user1 = results.find(r => r.agent === "user1@google.com");
  assert.strictEqual(user1.manager, "mgr1");
  assert.strictEqual(user1.location, "PH");
  assert.strictEqual(user1.latestEndIso, "2026-09-20T08:00:00+08:00");
  assert.strictEqual(user1.rowsInReport, 2);
  assert.strictEqual(user1.changedDuringDay, true); // Email Primary Pool changed
  assert.strictEqual(user1.chatSeenOnline, true);

  const user2 = results.find(r => r.agent === "user2@google.com");
  assert.strictEqual(user2.latestEndIso, "2026-09-20T09:00:00+08:00");
  assert.strictEqual(user2.rowsInReport, 1);
  assert.strictEqual(user2.changedDuringDay, false);
  assert.strictEqual(user2.chatSeenOnline, true);

  const user3 = results.find(r => r.agent === "user3@google.com");
  assert.strictEqual(user3.phoneSeenOnline, false);
  assert.strictEqual(user3.chatSeenOnline, false);
  assert.strictEqual(user3.emailSeenOnline, false);
});
