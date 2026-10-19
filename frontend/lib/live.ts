"use client";

// Live co-editing: while a signed-in workbook is open, a WebSocket to /api/live/{id} carries
// every cell change to the others editing it, and theirs to us, plus where each person's
// selection is. Cell edits travel as small patches; anything structural (rows, sheets, charts,
// formats) goes as a whole-workbook snapshot. Remote edits are written into the undo history
// too, so undoing your own change never throws away someone else's.

import { create } from "zustand";
import { API_URL } from "./api";
import { currentToken, setViewOnly, useAccount, whenPushed } from "./account";
import { editRefusal } from "./access";
import { loadComments } from "./comments";
import { applyOp, onStructuralOps, setEditGuard, useStore, type StructuralOp } from "./store";
import { useUi } from "./ui";
import type { CellSpec, Selection, SheetSpec, WorkbookSpec } from "./types";

export interface Peer {
  id: string;
  email: string;
  color: string;
  sel: (Selection & { sheet: string }) | null;
}

interface LiveState {
  status: "off" | "connecting" | "live" | "offline";
  docId: string | null;
  you: { id: string; email: string; color: string } | null;
  peers: Record<string, Peer>;
  /** Your role in the open workbook, from the server: viewers can look but not change. */
  role: "owner" | "editor" | "viewer" | null;
  /** The last change the server refused (a viewer typing, a locked cell). */
  refused: string | null;
}

export const useLive = create<LiveState>(() => ({ status: "off", docId: null, you: null, peers: {}, role: null, refused: null }));

interface CellPatch {
  sheet: string;
  r: number;
  c: number;
  cell: CellSpec | null;
}
interface HeaderPatch {
  sheet: string;
  c: number;
  header: string;
}

let ws: WebSocket | null = null;
let retry: ReturnType<typeof setTimeout> | undefined;
let attempts = 0;
/** The workbook as the room last saw it (after our sends and their patches). */
let synced: WorkbookSpec | null = null;
/** Set while we write a remote change into the store, so it isn't sent back. */
let applyingRemote = false;
let flushTimer: ReturnType<typeof setTimeout> | undefined;
let presenceTimer: ReturnType<typeof setTimeout> | undefined;

const send = (msg: unknown) => {
  if (ws?.readyState === WebSocket.OPEN) ws.send(JSON.stringify(msg));
};

/** Open (or move) the live connection to `docId`; null closes it. */
export function connectLive(docId: string | null) {
  if (useLive.getState().docId === docId && ws) return;
  disconnect();
  loadComments(null);
  useLive.setState({ docId, peers: {}, you: null, role: null, status: docId ? "connecting" : "off" });
  attempts = 0;
  if (docId) open(docId);
}

function disconnect() {
  clearTimeout(retry);
  clearTimeout(flushTimer);
  if (ws) {
    ws.onclose = null;
    ws.close();
  }
  ws = null;
  synced = null;
  pendingOps = [];
}

function open(docId: string) {
  const token = currentToken();
  if (!token) return useLive.setState({ status: "off" });
  // An empty API_URL means the backend serves this page too (the desktop app).
  const base = API_URL || window.location.origin;
  const url = `${base.replace(/^http/, "ws")}/api/live/${encodeURIComponent(docId)}?token=${encodeURIComponent(token)}`;
  const sock = new WebSocket(url);
  ws = sock;
  sock.onmessage = (e) => {
    if (ws !== sock) return;
    try {
      receive(JSON.parse(e.data));
    } catch {}
  };
  sock.onclose = (e) => {
    if (ws !== sock) return;
    ws = null;
    synced = null;
    // 4401: signed out; 4404: not in the account (yet) or no access. Neither gets better by retrying.
    if (e.code === 4401 || e.code === 4404) return useLive.setState({ status: "off", peers: {} });
    useLive.setState({ status: "offline", peers: {} });
    attempts++;
    retry = setTimeout(() => useLive.getState().docId === docId && open(docId), Math.min(30000, 1000 * 2 ** attempts));
  };
}

type Message =
  | { type: "state"; version: number; workbook: WorkbookSpec; you: LiveState["you"]; peers: Peer[]; role?: LiveState["role"] }
  | { type: "patch"; from: string; cells: CellPatch[]; headers: HeaderPatch[] }
  | { type: "snapshot"; from: string; workbook: WorkbookSpec }
  | { type: "ops"; from: string; ops: StructuralOp[] }
  | { type: "presence"; user: Omit<Peer, "sel">; sel: Peer["sel"] }
  | { type: "joined"; user: Omit<Peer, "sel"> }
  | { type: "left"; user_id: string }
  | { type: "ack" }
  | { type: "error"; message: string }
  | { type: "role"; role: NonNullable<LiveState["role"]> }
  | { type: "comments"; thread_id?: string };

function receive(msg: Message) {
  const live = useLive.getState();
  switch (msg.type) {
    case "state": {
      attempts = 0;
      loadComments(useLive.getState().docId);
      const others = msg.peers.filter((p) => p.id !== msg.you?.id);
      if (useLive.getState().docId) setViewOnly(useLive.getState().docId!, msg.role === "viewer");
      useLive.setState({ status: "live", you: msg.you, role: msg.role ?? null, peers: Object.fromEntries(others.map((p) => [p.id, p])) });
      const local = useStore.getState().workbook;
      if (others.length || msg.version > 0 || adoptServer || useLive.getState().role === "viewer") {
        adoptServer = false;
        // Others are already editing: their copy is the live one.
        synced = msg.workbook;
        if (JSON.stringify(msg.workbook) !== JSON.stringify(local)) remote(() => msg.workbook, true);
      } else {
        // First one in: this browser's copy (perhaps with edits not yet uploaded) becomes the room's.
        synced = msg.workbook;
        flush();
      }
      sendPresence();
      break;
    }
    case "patch":
      synced = synced && applyPatch(synced, msg.cells, msg.headers);
      remote((wb) => applyPatch(wb, msg.cells, msg.headers), false);
      break;
    case "snapshot":
      synced = msg.workbook;
      remote(() => msg.workbook, true);
      break;
    case "ops": {
      // Rows or columns moved: apply the same move to our copy and to every undo step, so undo keeps working.
      const withOps = (wb: WorkbookSpec) => {
        const next = structuredClone(wb);
        msg.ops.forEach((o) => applyOp(next, o));
        return next;
      };
      synced = synced && withOps(synced);
      remote(withOps, false);
      break;
    }
    case "presence":
    case "joined":
      if (msg.user.id === live.you?.id) break;
      useLive.setState({
        peers: { ...live.peers, [msg.user.id]: { ...msg.user, sel: msg.type === "presence" ? msg.sel : (live.peers[msg.user.id]?.sel ?? null) } },
      });
      break;
    case "role":
      // The owner changed what we can do; a new viewer also drops any unsent edits.
      useLive.setState({ role: msg.role });
      if (useLive.getState().docId) setViewOnly(useLive.getState().docId!, msg.role === "viewer");
      if (msg.role === "viewer") {
        adoptServer = true;
        reconnect();
      }
      break;
    case "comments":
      loadComments(useLive.getState().docId);
      break;
    case "error":
      // The server didn't apply our change (no edit access, or a locked cell): take its copy back.
      useLive.setState({ refused: msg.message });
      adoptServer = true;
      reconnect();
      break;
    case "left": {
      const peers = { ...live.peers };
      delete peers[msg.user_id];
      useLive.setState({ peers });
      break;
    }
  }
}

/** Write a remote change into the open workbook (and, for patches, its undo history). */
function remote(fn: (wb: WorkbookSpec) => WorkbookSpec, structural: boolean) {
  const st = useStore.getState();
  applyingRemote = true;
  try {
    const workbook = fn(st.workbook);
    useStore.setState({
      workbook,
      // A whole new workbook can't be merged into old undo steps, so those go.
      past: structural ? [] : st.past.map(fn),
      future: structural ? [] : st.future.map(fn),
      active: Math.min(st.active, Math.max(0, workbook.sheets.length - 1)),
    });
  } finally {
    applyingRemote = false;
  }
}

/** Copy-on-write: only the sheets and rows a patch touches are copied. */
export function applyPatch(wb: WorkbookSpec, cells: CellPatch[], headers: HeaderPatch[]): WorkbookSpec {
  const sheets = [...wb.sheets];
  const touched = new Map<number, SheetSpec>();
  const sheetFor = (name: string) => {
    const i = sheets.findIndex((s) => s.name === name);
    if (i < 0) return null;
    let s = touched.get(i);
    if (!s) {
      s = { ...sheets[i], rows: [...sheets[i].rows], columns: [...sheets[i].columns] };
      touched.set(i, s);
      sheets[i] = s;
    }
    return s;
  };
  for (const h of headers) {
    const s = sheetFor(h.sheet);
    if (!s) continue;
    while (s.columns.length <= h.c) s.columns.push({ header: "", format: "text" });
    s.columns[h.c] = { ...s.columns[h.c], header: h.header };
  }
  const copiedRows = new Set<string>();
  for (const p of cells) {
    const s = sheetFor(p.sheet);
    if (!s || p.r < 1) continue;
    while (s.columns.length <= p.c) s.columns.push({ header: "", format: "text" });
    while (s.rows.length < p.r) s.rows.push({ kind: "data", cells: [] });
    const key = `${p.sheet}:${p.r}`;
    if (!copiedRows.has(key)) {
      s.rows[p.r - 1] = { ...s.rows[p.r - 1], cells: [...(s.rows[p.r - 1].cells ?? [])] };
      copiedRows.add(key);
    }
    const row = s.rows[p.r - 1];
    while (row.cells.length <= p.c) row.cells.push({ value: null });
    row.cells[p.c] = p.cell ?? { value: null };
  }
  return { ...wb, sheets };
}

const json = (v: unknown) => JSON.stringify(v ?? null);

/** What changed between two workbooks, as cell patches, or null when only a snapshot will do. */
export function diffForLive(a: WorkbookSpec, b: WorkbookSpec): { cells: CellPatch[]; headers: HeaderPatch[] } | null {
  if (a.title !== b.title || a.sheets.length !== b.sheets.length || json(a.scenarios) !== json(b.scenarios)) return null;
  const cells: CellPatch[] = [];
  const headers: HeaderPatch[] = [];
  for (let i = 0; i < a.sheets.length; i++) {
    const x = a.sheets[i];
    const y = b.sheets[i];
    if (x === y) continue;
    const { rows: xr, columns: xc, ...xs } = x;
    const { rows: yr, columns: yc, ...ys } = y;
    if (json(xs) !== json(ys) || xr.length > yr.length || xc.length > yc.length) return null;
    for (let c = 0; c < yc.length; c++) {
      const { header: xh, ...xo } = xc[c] ?? { header: "" };
      const { header: yh, ...yo } = yc[c];
      if (c < xc.length && json(xo) !== json(yo)) return null;
      if (xh !== yh) headers.push({ sheet: y.name, c, header: yh });
    }
    for (let r = 0; r < yr.length; r++) {
      const p = xr[r];
      const q = yr[r];
      if (p === q) continue;
      if (p && ((p.kind ?? "data") !== (q.kind ?? "data") || p.height !== q.height)) return null;
      if (!p && ((q.kind ?? "data") !== "data" || q.height)) return null;
      const n = Math.max(p?.cells.length ?? 0, q.cells.length);
      for (let c = 0; c < n; c++) {
        const u = p?.cells[c] ?? null;
        const v = q.cells[c] ?? null;
        if (json(u) !== json(v)) cells.push({ sheet: y.name, r: r + 1, c, cell: v });
      }
    }
  }
  return cells.length + headers.length > 2000 ? null : { cells, headers };
}

/** Structural changes made here since the last flush, sent ahead of the cell changes. */
let pendingOps: StructuralOp[] = [];
/** After a refused change, take the server's copy on reconnect instead of sending ours again. */
let adoptServer = false;

function reconnect() {
  const { docId } = useLive.getState();
  disconnect();
  if (docId) {
    useLive.setState({ status: "connecting" });
    open(docId);
  }
}

/** Send what changed since the room last saw the workbook. */
function flush() {
  clearTimeout(flushTimer);
  if (!synced || ws?.readyState !== WebSocket.OPEN) return;
  if (pendingOps.length) {
    send({ type: "ops", ops: pendingOps });
    const next = structuredClone(synced);
    pendingOps.forEach((o) => applyOp(next, o));
    synced = next;
    pendingOps = [];
  }
  const now = useStore.getState().workbook;
  if (now === synced) return;
  const d = diffForLive(synced, now);
  if (!d) send({ type: "snapshot", workbook: now });
  else if (d.cells.length || d.headers.length) send({ type: "patch", ...d });
  synced = now;
}

function sendPresence() {
  clearTimeout(presenceTimer);
  presenceTimer = setTimeout(() => {
    const { selection, active, workbook } = useStore.getState();
    const sheet = workbook.sheets[active];
    if (sheet) send({ type: "presence", sel: { sheet: sheet.name, ...selection } });
  }, 80);
}

let started = false;
/** Follow the open workbook and the signed-in account; call once on startup. */
export function startLive() {
  if (started) return;
  started = true;
  const follow = () => {
    const signedIn = !!useAccount.getState().user;
    connectLive(signedIn ? useStore.getState().docId : null);
  };
  useAccount.subscribe((s, prev) => s.user !== prev.user && follow());
  useStore.subscribe((s, prev) => {
    if (s.docId !== prev.docId) return follow();
    if (s.workbook !== prev.workbook && !applyingRemote && synced) {
      clearTimeout(flushTimer);
      flushTimer = setTimeout(flush, 120);
    }
    if (s.selection !== prev.selection || s.active !== prev.active) sendPresence();
  });
  whenPushed(reconnectLive);
  setEditGuard((before, after) => {
    const why = useLive.getState().status === "live" ? editRefusal(useLive.getState().role, before, after) : null;
    if (why) useUi.getState().notify(why, true);
    return !why;
  });
  onStructuralOps((ops) => {
    if (synced) pendingOps.push(...ops);
  });
  follow();
}

/** Try again now, e.g. once a workbook has been uploaded to the account for the first time. */
export function reconnectLive() {
  const { docId, status } = useLive.getState();
  if (docId && status !== "live" && status !== "connecting") {
    disconnect();
    useLive.setState({ status: "connecting" });
    open(docId);
  }
}
