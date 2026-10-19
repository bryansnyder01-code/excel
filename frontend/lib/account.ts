"use client";

// Signing in, and keeping this browser's workbooks in step with the account.
// Without an account everything still works from this browser's storage.

import { create } from "zustand";
import { ApiError, auth, cloud, setAuthToken, type User } from "./api";
import { getDoc, listDocs, putDoc } from "./persist";
import type { SavedDoc } from "./types";

const TOKEN_KEY = "sheetcraft:token";

interface AccountState {
  user: User | null;
  /** Set while the first check of a saved sign-in is running. */
  checking: boolean;
  syncing: boolean;
  lastSync: number | null;
  syncError: string | null;
}

export const useAccount = create<AccountState>(() => ({
  user: null,
  checking: true,
  syncing: false,
  lastSync: null,
  syncError: null,
}));

function saveToken(token: string | null) {
  setAuthToken(token);
  try {
    if (token) localStorage.setItem(TOKEN_KEY, token);
    else localStorage.removeItem(TOKEN_KEY);
  } catch {}
}

/** On startup: pick up a saved sign-in, if it's still valid. */
export async function restoreSession() {
  let token: string | null = null;
  try {
    token = localStorage.getItem(TOKEN_KEY);
  } catch {}
  if (!token) return useAccount.setState({ checking: false });
  setAuthToken(token);
  try {
    const user = await auth.me();
    useAccount.setState({ user, checking: false });
    syncAll();
  } catch (e) {
    if (e instanceof ApiError && e.status === 401) saveToken(null);
    useAccount.setState({ checking: false });
  }
}

export async function signIn(kind: "login" | "signup", email: string, password: string) {
  const res = await auth[kind](email, password);
  saveToken(res.token);
  useAccount.setState({ user: res.user });
  await syncAll();
}

/** Finish any sign-in that ended with a session token: password reset, Google. */
export async function startSession(res: { token: string; user: User }) {
  saveToken(res.token);
  useAccount.setState({ user: res.user });
  await syncAll();
}

/** A password-reset link's token, waiting for the account dialog to ask for a new password. */
export let pendingReset: string | null = null;
export function setPendingReset(token: string | null) {
  pendingReset = token;
}

/** The signed-in token, for the live editing connection. */
export function currentToken(): string | null {
  try {
    return localStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

export async function signOut() {
  try {
    await auth.logout();
  } catch {}
  saveToken(null);
  useAccount.setState({ user: null, lastSync: null, syncError: null });
}

const signedIn = () => !!useAccount.getState().user;

function handle(e: unknown) {
  if (e instanceof ApiError && e.status === 401) {
    saveToken(null);
    useAccount.setState({ user: null, syncError: null });
    return;
  }
  useAccount.setState({ syncError: e instanceof Error ? e.message : "Sync failed." });
}

/**
 * Two-way sync: workbooks only in this browser go up, workbooks only in the account come down,
 * and where both have one, the newer copy wins.
 */
let onPulled: (ids: string[]) => void = () => {};
/** Called with the workbooks a sync brought down, so the open one can be refreshed. */
export function whenPulled(fn: (ids: string[]) => void) {
  onPulled = fn;
}

let onPushed: () => void = () => {};
/** Called after workbooks were uploaded, so a live connection waiting for its first upload can start. */
export function whenPushed(fn: () => void) {
  onPushed = fn;
}

// Workbooks shared with us view-only: the server refuses their uploads, so don't send them.
const viewOnly = new Set<string>();

export function setViewOnly(id: string, yes: boolean) {
  if (yes) viewOnly.add(id);
  else viewOnly.delete(id);
}

export async function syncAll(): Promise<void> {
  if (!signedIn() || useAccount.getState().syncing) return;
  useAccount.setState({ syncing: true, syncError: null });
  try {
    const remote = new Map((await cloud.list()).map((d) => [d.id, d]));
    for (const local of listDocs()) {
      const r = remote.get(local.id);
      remote.delete(local.id);
      setViewOnly(local.id, r?.role === "viewer");
      if (r?.role === "viewer") {
        if (r.updated > local.updated) remote.set(local.id, r);
        continue;
      }
      if (!r || local.updated > r.updated) await cloud.put(local.id, local.title, local.updated, local);
      else if (r.updated > local.updated) remote.set(local.id, r);
    }
    // Whatever is left is newer in the account, or not in this browser yet.
    for (const r of remote.values()) {
      const doc = (await cloud.get(r.id)).data as SavedDoc;
      putDoc({ ...doc, id: r.id, updated: r.updated });
    }
    if (remote.size) onPulled([...remote.keys()]);
    useAccount.setState({ lastSync: Date.now() });
    onPushed();
  } catch (e) {
    handle(e);
  } finally {
    useAccount.setState({ syncing: false });
  }
}

let pushTimer: ReturnType<typeof setTimeout> | undefined;

/** After a local save: send that workbook to the account a moment later. */
export function schedulePush(id: string) {
  if (!signedIn()) return;
  clearTimeout(pushTimer);
  pushTimer = setTimeout(async () => {
    const doc = getDoc(id);
    if (!doc || viewOnly.has(id)) return;
    try {
      await cloud.put(doc.id, doc.title, doc.updated, doc);
      useAccount.setState({ lastSync: Date.now(), syncError: null });
      onPushed();
    } catch (e) {
      handle(e);
    }
  }, 1500);
}

export async function removeFromAccount(id: string) {
  if (!signedIn()) return;
  try {
    await cloud.remove(id);
  } catch (e) {
    handle(e);
  }
}
