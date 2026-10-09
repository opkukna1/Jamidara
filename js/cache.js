// Tiny IndexedDB key-value cache (big arrays don't fit in localStorage).
// Used so the app does NOT re-read the whole Firestore on every page open.
const DB = 'jamidara-cache', STORE = 'kv';

const open = () => new Promise((resolve, reject) => {
  const r = indexedDB.open(DB, 1);
  r.onupgradeneeded = () => r.result.createObjectStore(STORE);
  r.onsuccess = () => resolve(r.result);
  r.onerror = () => reject(r.error);
});

export async function cacheGet(key) {
  try {
    const db = await open();
    return await new Promise(resolve => {
      const q = db.transaction(STORE).objectStore(STORE).get(key);
      q.onsuccess = () => resolve(q.result ?? null);
      q.onerror = () => resolve(null);
    });
  } catch (e) { return null; }
}

export async function cacheSet(key, value) {
  try {
    const db = await open();
    await new Promise(resolve => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  } catch (e) { /* cache is best-effort */ }
}

export async function cacheClear(key) {
  try {
    const db = await open();
    await new Promise(resolve => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  } catch (e) { /* ignore */ }
}

export async function cacheClearPrefix(prefix) {
  try {
    const db = await open();
    await new Promise(resolve => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      const req = store.openKeyCursor();
      req.onsuccess = () => {
        const c = req.result;
        if (!c) return;
        if (String(c.key).startsWith(prefix)) store.delete(c.key);
        c.continue();
      };
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  } catch (e) { /* ignore */ }
}
