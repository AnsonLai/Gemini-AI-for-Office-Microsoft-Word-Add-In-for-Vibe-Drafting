import assert from 'node:assert/strict';
import { updateCheckpoint } from '../src/taskpane/modules/storage/checkpoint-store.js';

// Minimal IndexedDB fake: get/put requests settle, then the transaction completes.
function installStore(records) {
  const rows = new Map(records.map(record => [record.id, { ...record }]));
  let closed = 0;
  const db = {
    close() { closed++; },
    transaction(name, mode) {
      assert.equal(name, 'checkpoints');
      assert.equal(mode, 'readwrite');
      const t = {};
      t.objectStore = () => ({
        get(id) {
          const request = { result: rows.get(id) };
          queueMicrotask(() => { request.onsuccess?.(); queueMicrotask(() => t.oncomplete?.()); });
          return request;
        },
        put(record) { rows.set(record.id, record); },
      });
      return t;
    },
  };
  globalThis.indexedDB = { open() {
    const request = { result: db };
    queueMicrotask(() => request.onsuccess());
    return request;
  } };
  return { rows, get closed() { return closed; } };
}

let store = installStore([{ id: 3, label: 'auto', ooxml: '<body/>', sessionId: 'pane' }]);
await updateCheckpoint(3, { afterToken: 'token-after-edit' });
assert.deepEqual(store.rows.get(3), { id: 3, label: 'auto', ooxml: '<body/>', sessionId: 'pane', afterToken: 'token-after-edit' });
assert.equal(store.closed, 1);

store = installStore([]);
await updateCheckpoint(9, { afterToken: 'ignored' });
assert.equal(store.rows.has(9), false, 'a pruned checkpoint is not recreated by a late token report');
assert.equal(store.closed, 1);

console.log('PASS: checkpoint updates merge fields and never recreate a missing checkpoint');
