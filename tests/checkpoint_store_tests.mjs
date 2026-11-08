import assert from 'assert';
import {
  formatAutoCheckpointLabel,
  makeCheckpointRecord,
  idsToEvict,
  migrateLegacyCheckpoints,
  idsToPrune,
  MAX_CHECKPOINTS,
  OTHER_SESSION_MAX_AGE_MS
} from '../src/taskpane/modules/storage/checkpoint-store.js';

function testFormatAutoCheckpointLabel() {
  const d = new Date('2026-06-12T15:04:05.000Z');
  assert.strictEqual(formatAutoCheckpointLabel('apply_redlines', d), 'auto:apply_redlines:2026-06-12T15:04:05.000Z');
  // Accepts epoch millis / string dates too.
  assert.strictEqual(formatAutoCheckpointLabel('edit_table', d.getTime()), 'auto:edit_table:2026-06-12T15:04:05.000Z');
  // Missing tool name falls back to "unknown".
  assert.ok(formatAutoCheckpointLabel(null, d).startsWith('auto:unknown:'));
}

function testMakeCheckpointRecord() {
  const rec = makeCheckpointRecord('manual', 'PK…', 1234);
  assert.deepStrictEqual(rec, { timestamp: 1234, label: 'manual', ooxml: 'PK…' });
  // Defaults: empty label -> "manual", null ooxml -> "".
  const rec2 = makeCheckpointRecord('', null, 1);
  assert.strictEqual(rec2.label, 'manual');
  assert.strictEqual(rec2.ooxml, '');
}

function testIdsToEvictUnderCap() {
  const records = [{ id: 1 }, { id: 2 }, { id: 3 }];
  assert.deepStrictEqual(idsToEvict(records, 10), []);
  assert.deepStrictEqual(idsToEvict([], 10), []);
  assert.deepStrictEqual(idsToEvict(null, 10), []);
}

function testIdsToEvictOverCap() {
  // 12 records, cap 10 -> evict the 2 oldest by id.
  const records = [];
  for (let i = 1; i <= 12; i++) records.push({ id: i });
  assert.deepStrictEqual(idsToEvict(records, 10), [1, 2]);
  // Unsorted input is handled (sorts by id first).
  const shuffled = [{ id: 5 }, { id: 1 }, { id: 3 }, { id: 2 }, { id: 4 }];
  assert.deepStrictEqual(idsToEvict(shuffled, 3), [1, 2]);
}

function testDefaultCapIsTen() {
  assert.strictEqual(MAX_CHECKPOINTS, 10);
  const records = [];
  for (let i = 1; i <= 11; i++) records.push({ id: i });
  assert.deepStrictEqual(idsToEvict(records), [1]); // uses default MAX_CHECKPOINTS
}

function testMigrateLegacyCheckpoints() {
  const legacy = ['<ooxml1>', '', '<ooxml2>', null, 42, '<ooxml3>'];
  const records = migrateLegacyCheckpoints(legacy, 1000);
  // Only the 3 non-empty strings survive, labeled "migrated", with increasing timestamps.
  assert.strictEqual(records.length, 3);
  assert.deepStrictEqual(records.map((r) => r.ooxml), ['<ooxml1>', '<ooxml2>', '<ooxml3>']);
  assert.ok(records.every((r) => r.label === 'migrated'));
  assert.deepStrictEqual(records.map((r) => r.timestamp), [1000, 1001, 1002]);
  // Non-array input is safe.
  assert.deepStrictEqual(migrateLegacyCheckpoints(null), []);
}

function testMakeCheckpointRecordSession() {
  assert.deepStrictEqual(makeCheckpointRecord('auto', '<x/>', 5, 'pane-a'), { timestamp: 5, label: 'auto', ooxml: '<x/>', sessionId: 'pane-a' });
}

function testIdsToPrune() {
  const now = 10 * OTHER_SESSION_MAX_AGE_MS;
  const fresh = now - 1000;
  const stale = now - OTHER_SESSION_MAX_AGE_MS - 1;
  const records = [];
  // 12 own checkpoints: only the 2 oldest go, however old the others' are.
  for (let i = 1; i <= 12; i++) records.push({ id: i, sessionId: 'mine', timestamp: stale });
  records.push({ id: 20, sessionId: 'other', timestamp: fresh }); // another open pane: kept
  records.push({ id: 21, sessionId: 'other', timestamp: stale }); // abandoned pane: pruned
  records.push({ id: 22, timestamp: fresh });                      // untagged legacy, still fresh: kept
  records.push({ id: 23 });                                        // no timestamp: pruned
  assert.deepStrictEqual(idsToPrune(records, { sessionId: 'mine', now }).sort((a, b) => a - b), [1, 2, 21, 23]);
  // Another document's activity never evicts this session's checkpoints.
  const busyOther = Array.from({ length: 15 }, (_, i) => ({ id: 100 + i, sessionId: 'other', timestamp: fresh }));
  const mine = [{ id: 1, sessionId: 'mine', timestamp: fresh }];
  assert.deepStrictEqual(idsToPrune([...mine, ...busyOther], { sessionId: 'other', now }), [100, 101, 102, 103, 104]);
  assert.deepStrictEqual(idsToPrune(null), []);
}

testFormatAutoCheckpointLabel();
testMakeCheckpointRecordSession();
testIdsToPrune();
testMakeCheckpointRecord();
testIdsToEvictUnderCap();
testIdsToEvictOverCap();
testDefaultCapIsTen();
testMigrateLegacyCheckpoints();

console.log('checkpoint_store_tests passed');
