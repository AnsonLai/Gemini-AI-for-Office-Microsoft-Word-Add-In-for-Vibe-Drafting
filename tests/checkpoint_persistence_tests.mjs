import assert from 'node:assert/strict';
import { saveCheckpoint } from '../src/taskpane/modules/storage/checkpoint-store.js';

function database() {
  let transaction, addRequest, closed = false;
  const db = {
    close() { closed = true; },
    transaction(name, mode) {
      assert.equal(name, 'checkpoints');
      if (mode === 'readwrite') {
        transaction = { objectStore() { return { add(record) { addRequest = { result: 7 }; assert.equal(record.ooxml, '<saved-body/>'); return addRequest; } }; } };
        return transaction;
      }
      // Cap enforcement is queued only after the checkpoint transaction commits.
      return { objectStore() { return { getAll() {
        const request = { result: [{ id: 7 }] };
        queueMicrotask(() => request.onsuccess());
        return request;
      } }; } };
    },
  };
  globalThis.indexedDB = { open() {
    const request = { result: db };
    queueMicrotask(() => request.onsuccess());
    return request;
  } };
  return { get transaction() { return transaction; }, get request() { return addRequest; }, get closed() { return closed; } };
}

let host = database(), settled = false;
let pending = saveCheckpoint('auto:apply_document_edits', '<saved-body/>');
pending.then(() => { settled = true; });
await new Promise(resolve => setImmediate(resolve));
host.request.onsuccess();
await new Promise(resolve => setImmediate(resolve));
assert.equal(settled, false, 'a successful IndexedDB add is not a persisted checkpoint until transaction completion');
host.transaction.oncomplete();
assert.equal(await pending, 7);
assert.equal(host.closed, true);

for (const event of ['onabort', 'onerror']) {
  host = database();
  pending = saveCheckpoint('auto:apply_document_edits', '<saved-body/>');
  const rejected = assert.rejects(pending, /storage failure/);
  await new Promise(resolve => setImmediate(resolve));
  host.request.onsuccess();
  host.transaction.error = new Error('storage failure');
  host.transaction[event]();
  await rejected;
  assert.equal(host.closed, true);
}
console.log('PASS: exact-source checkpoint waits for committed storage and rejects post-request abort/error');
