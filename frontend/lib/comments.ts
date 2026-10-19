"use client";

// Comment threads on cells of a workbook in the account: anyone with access, viewers included,
// can comment; @someone@example.com emails that person a link to the cell.

import { create } from "zustand";
import { send } from "./api";

export interface CommentAuthor {
  id: string;
  email: string;
}

export interface CommentItem {
  id: string;
  author: CommentAuthor;
  body: string;
  created: number;
}

export interface Thread {
  id: string;
  sheet: string;
  /** Excel address, e.g. "B5". */
  cell: string;
  resolved: boolean;
  comments: CommentItem[];
}

interface CommentsState {
  docId: string | null;
  threads: Thread[];
  loading: boolean;
  error: string | null;
}

export const useComments = create<CommentsState>(() => ({ docId: null, threads: [], loading: false, error: null }));

const base = (docId: string) => `/api/docs/${encodeURIComponent(docId)}/comments`;

export async function loadComments(docId: string | null) {
  if (!docId) return useComments.setState({ docId: null, threads: [], error: null });
  if (useComments.getState().docId !== docId) useComments.setState({ docId, threads: [] });
  useComments.setState({ loading: true });
  try {
    const threads = await send<Thread[]>("GET", base(docId));
    if (useComments.getState().docId === docId) useComments.setState({ threads, error: null });
  } catch (e) {
    if (useComments.getState().docId === docId) useComments.setState({ threads: [], error: e instanceof Error ? e.message : "Comments didn't load." });
  } finally {
    useComments.setState({ loading: false });
  }
}

function replace(thread: Thread) {
  const { threads } = useComments.getState();
  const i = threads.findIndex((t) => t.id === thread.id);
  useComments.setState({ threads: i < 0 ? [...threads, thread] : threads.map((t) => (t.id === thread.id ? thread : t)) });
}

export async function addComment(docId: string, sheet: string, cell: string, body: string, threadId?: string) {
  replace(await send<Thread>("POST", base(docId), { sheet, cell, body, thread_id: threadId ?? null }));
}

export async function setResolved(docId: string, thread: Thread, resolved: boolean) {
  await send<unknown>("PATCH", `${base(docId)}/${encodeURIComponent(thread.id)}`, { resolved });
  replace({ ...thread, resolved });
}

export async function deleteComment(docId: string, thread: Thread, commentId: string) {
  await send<void>("DELETE", `${base(docId)}/${encodeURIComponent(thread.id)}/${encodeURIComponent(commentId)}`);
  const rest = thread.comments.filter((c) => c.id !== commentId);
  if (rest.length) replace({ ...thread, comments: rest });
  else useComments.setState({ threads: useComments.getState().threads.filter((t) => t.id !== thread.id) });
}

/** Grid cells (r, c) of `sheet` that have an open thread, as "r:c". */
export function openThreadCells(threads: Thread[], sheet: string): Set<string> {
  const out = new Set<string>();
  for (const t of threads) {
    if (t.resolved || t.sheet !== sheet) continue;
    const m = /^([A-Z]+)(\d+)$/.exec(t.cell.toUpperCase());
    if (!m) continue;
    const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    out.add(`${Number(m[2]) - 1}:${c}`);
  }
  return out;
}
