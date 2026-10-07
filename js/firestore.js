// All Firebase / Firestore code lives here (no UI code).
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getFirestore, collection, getDocs, doc, writeBatch } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, signOut, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { firebaseConfig } from "../firebase-config.js";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth(app);

export const loadMappings = async () => (await getDocs(collection(db, 'mappings'))).docs.map(d => ({ _id: d.id, ...d.data() }));
export const loadDispatchData = async () => (await getDocs(collection(db, 'dispatches'))).docs.map(d => ({ _id: d.id, ...d.data() }));
export async function googleLogin() { // popup first, redirect fallback for browsers that block popups
  const p = new GoogleAuthProvider();
  try { await signInWithPopup(auth, p); }
  catch (e) { if (/popup-blocked|operation-not-supported/.test(e.code)) await signInWithRedirect(auth, p); else throw e; }
}
export const logout = () => signOut(auth);
export const onUser = cb => onAuthStateChanged(auth, cb);

// Stable doc id: Year + Unit ID (no duplicates; re-upload updates in place).
export const dispatchId = (year, unitId) => `${year}_${String(unitId).trim()}`.replace(/\//g, '-');

/** Upsert rows with merge (never deletes). existingIds = Set of ids already in Firestore. */
export async function uploadDispatchRows(rows, year, existingIds, onProgress) {
  const res = { inserted: 0, updated: 0, skipped: 0, errors: 0 };
  const valid = [];
  for (const r of rows) {
    if (!String(r['Unit ID'] ?? '').trim()) { res.skipped++; continue; }
    valid.push(r);
  }
  for (let i = 0; i < valid.length; i += 400) {
    const slice = valid.slice(i, i + 400), batch = writeBatch(db);
    slice.forEach(r => batch.set(doc(db, 'dispatches', dispatchId(year, r['Unit ID'])), { ...r, Year: year }, { merge: true }));
    try {
      await batch.commit();
      slice.forEach(r => existingIds.has(dispatchId(year, r['Unit ID'])) ? res.updated++ : res.inserted++);
    } catch (e) { res.errors += slice.length; if (/permission/.test(e.code)) throw e; }
    onProgress && onProgress(Math.min(i + 400, valid.length), valid.length);
  }
  return res;
}

// Stable mapping id from English names (so re-upload updates instead of duplicating).
const slug = s => String(s ?? '').toLowerCase().replace(/gram\s+panchayat|panchayat\s+samiti/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
export const mappingId = r => [r.DIST_EN, r.PS_EN, r.GP_EN].map(slug).join('__');

/** Upsert mapping rows (merge, never deletes). Duplicate rows inside the file are skipped. */
export async function uploadMappingRows(rows, existingIds, onProgress) {
  const res = { inserted: 0, updated: 0, skipped: 0, errors: 0 };
  const seen = new Set(), valid = [];
  for (const r of rows) {
    const id = mappingId(r);
    if (!r.GP_EN || !r.PS_EN || !r.DIST_EN || seen.has(id)) { res.skipped++; continue; }
    seen.add(id); valid.push([id, r]);
  }
  for (let i = 0; i < valid.length; i += 400) {
    const slice = valid.slice(i, i + 400), batch = writeBatch(db);
    slice.forEach(([id, r]) => batch.set(doc(db, 'mappings', id), r, { merge: true }));
    try { await batch.commit(); slice.forEach(([id]) => existingIds.has(id) ? res.updated++ : res.inserted++); }
    catch (e) { res.errors += slice.length; if (/permission/.test(e.code)) throw e; }
    onProgress && onProgress(Math.min(i + 400, valid.length), valid.length);
  }
  return res;
}
