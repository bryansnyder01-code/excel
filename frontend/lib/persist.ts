// Saved workbooks live in this browser's localStorage. Every call is guarded:
// storage can be unavailable (private mode) or full, and the app must keep working.

import type { SavedDoc } from "./types";

const DOCS_KEY = "sheetcraft:docs";
const LAST_KEY = "sheetcraft:last";

function read(): SavedDoc[] {
  try {
    const raw = localStorage.getItem(DOCS_KEY);
    return raw ? (JSON.parse(raw) as SavedDoc[]) : [];
  } catch {
    return [];
  }
}

function write(docs: SavedDoc[]): boolean {
  try {
    localStorage.setItem(DOCS_KEY, JSON.stringify(docs));
    return true;
  } catch {
    return false;
  }
}

export function listDocs(): SavedDoc[] {
  return read().sort((a, b) => b.updated - a.updated);
}

export function getDoc(id: string): SavedDoc | undefined {
  return read().find((d) => d.id === id);
}

/** Save or update a document. If storage is full, retry without its version history, then drop the oldest docs. */
export function saveDoc(doc: SavedDoc): boolean {
  const others = read().filter((d) => d.id !== doc.id);
  if (write([doc, ...others])) return remember(doc.id);
  const lean = { ...doc, checkpoints: doc.checkpoints.slice(-3) };
  if (write([lean, ...others])) return remember(doc.id);
  const sorted = others.sort((a, b) => b.updated - a.updated);
  while (sorted.length) {
    sorted.pop();
    if (write([lean, ...sorted])) return remember(doc.id);
  }
  return false;
}

/** Store a document (e.g. one synced from the account) without making it the one reopened next time. */
export function putDoc(doc: SavedDoc): boolean {
  return write([doc, ...read().filter((d) => d.id !== doc.id)]);
}

export function deleteDoc(id: string) {
  write(read().filter((d) => d.id !== id));
}

function remember(id: string): boolean {
  try {
    localStorage.setItem(LAST_KEY, id);
  } catch {}
  return true;
}

export function lastDocId(): string | null {
  try {
    return localStorage.getItem(LAST_KEY);
  } catch {
    return null;
  }
}
