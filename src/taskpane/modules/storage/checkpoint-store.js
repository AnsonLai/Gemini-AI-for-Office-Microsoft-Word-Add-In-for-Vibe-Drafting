/* global indexedDB */
/**
 * checkpoint-store.js
 *
 * IndexedDB-backed document checkpoint storage. Replaces the previous
 * localStorage array (which capped out at the ~5MB localStorage quota and could
 * not hold large documents).
 *
 * The pure helpers (label formatting, cap enforcement, legacy migration
 * transform) are exported separately and contain no IndexedDB access, so they
 * can be unit tested in plain Node where `indexedDB` is unavailable. The
 * IndexedDB functions reference `indexedDB` only inside function bodies, so the
 * module itself is still importable in Node.
 */

const DB_NAME = "aiwordplugin-checkpoints";
const STORE_NAME = "checkpoints";
const DB_VERSION = 1;
export const MAX_CHECKPOINTS = 10;
// Revert buttons only exist in the pane session that made a checkpoint, so other
// sessions' checkpoints are kept just long enough for a pane left open alongside.
export const OTHER_SESSION_MAX_AGE_MS = 24 * 60 * 60 * 1000;

// ---------------------------------------------------------------------------
// Pure helpers (no IndexedDB) — unit tested directly.
// ---------------------------------------------------------------------------

/**
 * Auto-checkpoint label format: `auto:<toolName>:<ISO timestamp>`.
 * @param {string} toolName
 * @param {Date|number|string} [date]
 * @returns {string}
 */
export function formatAutoCheckpointLabel(toolName, date = new Date()) {
  const d = date instanceof Date ? date : new Date(date);
  return `auto:${toolName || "unknown"}:${d.toISOString()}`;
}

/**
 * Build a checkpoint record (the `id` is assigned by IndexedDB on insert).
 * @param {string} label
 * @param {string} ooxml
 * @param {number} [timestamp]
 * @param {string} [sessionId] the pane session that can offer this checkpoint
 * @returns {{ timestamp: number, label: string, ooxml: string, sessionId?: string }}
 */
export function makeCheckpointRecord(label, ooxml, timestamp = Date.now(), sessionId) {
  return {
    timestamp,
    label: label || "manual",
    ooxml: ooxml == null ? "" : String(ooxml),
    ...(sessionId ? { sessionId } : {}),
  };
}

/**
 * Ids to delete after `sessionId` saves: its own checkpoints beyond the newest
 * `maxPerSession`, plus any other session's (or untagged) checkpoints older
 * than `maxAgeMs`. One busy document no longer evicts another's checkpoints.
 * @returns {number[]}
 */
export function idsToPrune(records, {
  sessionId = null, maxPerSession = MAX_CHECKPOINTS, maxAgeMs = OTHER_SESSION_MAX_AGE_MS, now = Date.now(),
} = {}) {
  if (!Array.isArray(records)) return [];
  const own = records.filter((r) => sessionId && r.sessionId === sessionId);
  const ownEvicted = idsToEvict(own, maxPerSession);
  const staleOthers = records
    .filter((r) => !(sessionId && r.sessionId === sessionId))
    .filter((r) => !Number.isFinite(r.timestamp) || now - r.timestamp > maxAgeMs)
    .map((r) => r.id);
  return [...ownEvicted, ...staleOthers];
}

/**
 * Given checkpoint records, return the ids that must be evicted (oldest first)
 * to keep at most `maxCount`.
 * @param {Array<{id:number}>} records
 * @param {number} [maxCount]
 * @returns {number[]}
 */
export function idsToEvict(records, maxCount = MAX_CHECKPOINTS) {
  if (!Array.isArray(records) || records.length <= maxCount) {
    return [];
  }
  const sorted = [...records].sort((a, b) => (a.id || 0) - (b.id || 0));
  return sorted.slice(0, sorted.length - maxCount).map((r) => r.id);
}

/**
 * Transform legacy localStorage checkpoints (an array of OOXML strings) into
 * checkpoint records for one-time migration into IndexedDB.
 * @param {string[]} legacyArray
 * @param {number} [now]
 * @returns {Array<{timestamp:number,label:string,ooxml:string}>}
 */
export function migrateLegacyCheckpoints(legacyArray, now = Date.now()) {
  if (!Array.isArray(legacyArray)) {
    return [];
  }
  return legacyArray
    .filter((ooxml) => typeof ooxml === "string" && ooxml.length > 0)
    .map((ooxml, i) => makeCheckpointRecord("migrated", ooxml, now + i));
}

// ---------------------------------------------------------------------------
// IndexedDB plumbing.
// ---------------------------------------------------------------------------

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id", autoIncrement: true });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function listFromDb(db) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE_NAME, "readonly");
    const req = t.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

function deleteIdsFromDb(db, ids) {
  if (!ids || ids.length === 0) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE_NAME, "readwrite");
    const store = t.objectStore(STORE_NAME);
    ids.forEach((id) => store.delete(id));
    t.oncomplete = () => resolve();
    t.onerror = () => reject(t.error);
  });
}

async function enforceCapInDb(db) {
  const all = await listFromDb(db);
  await deleteIdsFromDb(db, idsToEvict(all, MAX_CHECKPOINTS));
}

async function pruneInDb(db, sessionId) {
  const all = await listFromDb(db);
  await deleteIdsFromDb(db, idsToPrune(all, { sessionId }));
}

/**
 * Persist a checkpoint and return its new id. Keeps at most MAX_CHECKPOINTS
 * for the saving session and drops other sessions' stale checkpoints.
 * @param {string} label
 * @param {string} ooxml
 * @param {{ sessionId?: string }} [options]
 * @returns {Promise<number>} the new checkpoint id
 */
export async function saveCheckpoint(label, ooxml, { sessionId } = {}) {
  const db = await openDb();
  try {
    const record = makeCheckpointRecord(label, ooxml, Date.now(), sessionId);
    const id = await new Promise((resolve, reject) => {
      const t = db.transaction(STORE_NAME, "readwrite");
      const req = t.objectStore(STORE_NAME).add(record);
      let checkpointId;
      req.onsuccess = () => { checkpointId = req.result; };
      req.onerror = () => reject(req.error);
      // A successful add request can still be rolled back by a transaction abort.
      t.oncomplete = () => resolve(checkpointId);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error || new Error('Checkpoint persistence was aborted.'));
    });
    if (sessionId) await pruneInDb(db, sessionId);
    else await enforceCapInDb(db);
    return id;
  } finally {
    db.close();
  }
}

/** Delete other sessions' checkpoints past their age limit (run at pane startup). */
export async function pruneCheckpoints(sessionId) {
  const db = await openDb();
  try {
    await pruneInDb(db, sessionId);
  } finally {
    db.close();
  }
}

/**
 * Merge fields into an existing checkpoint (for example the document token
 * observed right after its edit was verified). Missing ids are ignored.
 */
export async function updateCheckpoint(id, fields) {
  const db = await openDb();
  try {
    await new Promise((resolve, reject) => {
      const t = db.transaction(STORE_NAME, "readwrite");
      const store = t.objectStore(STORE_NAME);
      const req = store.get(id);
      req.onsuccess = () => {
        if (req.result) store.put({ ...req.result, ...fields, id });
      };
      req.onerror = () => reject(req.error);
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error || new Error('Checkpoint update was aborted.'));
    });
  } finally {
    db.close();
  }
}

/**
 * Insert pre-built records (used for one-time legacy migration). Enforces cap.
 * @param {Array<{timestamp:number,label:string,ooxml:string}>} records
 */
export async function importCheckpoints(records) {
  if (!Array.isArray(records) || records.length === 0) return;
  const db = await openDb();
  try {
    await new Promise((resolve, reject) => {
      const t = db.transaction(STORE_NAME, "readwrite");
      const store = t.objectStore(STORE_NAME);
      records.forEach((r) => store.add(r));
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
    });
    await enforceCapInDb(db);
  } finally {
    db.close();
  }
}

/**
 * @param {number} id
 * @returns {Promise<object|null>} the checkpoint record, or null if missing
 */
export async function getCheckpoint(id) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const t = db.transaction(STORE_NAME, "readonly");
      const req = t.objectStore(STORE_NAME).get(id);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}
