// The =AI("instruction", A2, ...) spreadsheet function.
//
// Evaluation never waits: lookupAi() answers from an in-memory cache, or says "loading" and queues
// the question. A debounced batcher sends up to 100 queued questions per /ai-eval call, fills the cache
// and nudges the store with a new workbook object, so evaluateCached computes again with the answers.
// The answer each AI cell shows is kept in the cell's ai_value (without an undo step) so downloads
// and the next page load have it; on load a cell's ai_value is trusted until its inputs change.

import { aiEval, AI_EVAL_MAX_ITEMS, type AiEvalItem } from "./aiApi";
import { useStore } from "./store";
import type { WorkbookSpec } from "./types";
import { useUi } from "./ui";

export const AI_LOADING = "Loading...";
const RETRY_AFTER_MS = 60_000;
const DEBOUNCE_MS = 250;

export interface AiValueUpdate {
  /** Sheet index, 0-based row index into sheet.rows, 0-based column. */
  sheet: number;
  row: number;
  col: number;
  value: string;
}

export type AiLookup = { status: "done"; value: string } | { status: "loading" } | { status: "error"; message: string };

/** Where an AI formula lives (Excel row/col, 1-based) and the answer it last showed. */
export interface AiCellRef {
  sheet: string;
  row: number;
  col: number;
  aiValue: string | null;
}

const answers = new Map<string, string>();
const failures = new Map<string, { at: number; message: string }>();
const inflight = new Set<string>();
const queue = new Map<string, AiEvalItem>();
/** The key each AI cell of the open document last asked for, to tell when its inputs changed. */
const cellKeys = new Map<string, string>();
/** Why an AI cell shows #AI!, for the formula bar. */
const cellErrors = new Map<string, string>();

let timer: ReturnType<typeof setTimeout> | undefined;
let running = false;
let lastToast = 0;
let now = () => Date.now();

type Sender = (items: AiEvalItem[]) => Promise<Record<string, string>>;
let sender: Sender = async (items) => (await aiEval(items)).results;

/** Every answer seen, so an ai_value copied along with a formula (fill, paste) isn't mistaken for this cell's answer. */
const seenValues = new Set<string>();
/** The first workbook object evaluated for each document: its saved answers are all trusted. */
const firstBook = new Map<string, WorkbookSpec>();

function remember(key: string, value: string) {
  answers.set(key, value);
  seenValues.add(value);
}

/**
 * Whether a cell's saved ai_value can stand for inputs we haven't seen at that position before.
 * Yes when the document has just been opened; later (a new cell from fill, paste or an edit) only
 * if that answer hasn't already been given to some other question, which would mean it was copied.
 */
function trusted(wb: WorkbookSpec, value: string): boolean {
  const doc = docScope();
  if (!firstBook.has(doc) && isCurrent(wb)) firstBook.set(doc, wb);
  return firstBook.get(doc) === wb || !seenValues.has(value);
}

/** The cache key: the instruction and every input as displayed. */
export const aiKey = (prompt: string, inputs: string[]) => JSON.stringify([prompt, ...inputs]);

const docScope = () => {
  try {
    return useStore.getState().docId;
  } catch {
    return "";
  }
};
const cellId = (sheet: string, row: number, col: number) => `${docScope()}|${sheet.toLowerCase()}|${row}|${col}`;

const isCurrent = (wb: WorkbookSpec) => {
  try {
    return useStore.getState().workbook === wb;
  } catch {
    return false;
  }
};

/**
 * The answer for one AI() call. Only the open workbook queues requests: copies evaluated for
 * previews or goal seek get cached answers (or the cell's saved one) and otherwise "loading".
 */
export function lookupAi(wb: WorkbookSpec, prompt: string, inputs: string[], at: AiCellRef | null): AiLookup {
  const key = aiKey(prompt, inputs);
  const live = isCurrent(wb);
  const id = at ? cellId(at.sheet, at.row, at.col) : null;
  const before = id ? cellKeys.get(id) : undefined;
  if (id && live) cellKeys.set(id, key);
  const known = answers.get(key);
  if (known !== undefined) {
    if (id) cellErrors.delete(id);
    return { status: "done", value: known };
  }
  // A saved answer stands until this cell's inputs change.
  if (at?.aiValue != null && before === undefined && trusted(wb, at.aiValue)) {
    if (live) remember(key, at.aiValue);
    return { status: "done", value: at.aiValue };
  }
  const failed = failures.get(key);
  if (failed && now() - failed.at < RETRY_AFTER_MS) {
    if (id) cellErrors.set(id, failed.message);
    return { status: "error", message: failed.message };
  }
  if (id) cellErrors.delete(id);
  if (!live || inflight.has(key) || queue.has(key)) return { status: "loading" };
  failures.delete(key);
  queue.set(key, { key, prompt, inputs });
  schedule();
  return { status: "loading" };
}

/** Why the AI formula at grid (r, c) of sheet `si` in the open workbook shows #AI!, if it does. */
export function aiErrorAt(si: number, r: number, c: number): string | null {
  const sheet = useStore.getState().workbook.sheets[si];
  if (!sheet || r < 1) return null;
  return cellErrors.get(cellId(sheet.name, r + 1, c + 1)) ?? null;
}

function schedule(delay = DEBOUNCE_MS) {
  if (running || timer !== undefined) return;
  timer = setTimeout(() => {
    timer = undefined;
    void flush();
  }, delay);
}

/** Send the next batch of queued questions. Exported for tests; normally runs on its own. */
export async function flush(): Promise<void> {
  if (running || !queue.size) return;
  running = true;
  const batch = [...queue.values()].slice(0, AI_EVAL_MAX_ITEMS);
  for (const item of batch) {
    queue.delete(item.key);
    inflight.add(item.key);
  }
  try {
    const results = await sender(batch);
    for (const item of batch) {
      const v = results?.[item.key];
      if (typeof v === "string") remember(item.key, v);
      else if (typeof v === "number") remember(item.key, String(v));
      else failures.set(item.key, { at: now(), message: "The AI didn't return an answer for this cell." });
    }
  } catch (e) {
    const message = e instanceof Error && e.message ? e.message : "The AI request failed.";
    for (const item of batch) failures.set(item.key, { at: now(), message });
    if (now() - lastToast > RETRY_AFTER_MS) {
      lastToast = now();
      try {
        useUi.getState().notify(`AI formulas: ${message}`, true);
      } catch {}
    }
  } finally {
    for (const item of batch) inflight.delete(item.key);
    running = false;
    refresh();
    if (queue.size) schedule(0);
  }
}

/** A new workbook object (same content) so every view evaluates again with the new answers. */
function refresh() {
  try {
    const { workbook } = useStore.getState();
    useStore.setState({ workbook: { ...workbook } });
  } catch {}
}

/**
 * Called by the formula engine when AI cells show something other than their saved ai_value.
 * Writes the new values a moment later, only if that workbook is still the open one, and without
 * touching undo history.
 */
export function aiValuesStale(wb: WorkbookSpec, apply: () => WorkbookSpec) {
  if (!isCurrent(wb)) return;
  setTimeout(() => {
    if (isCurrent(wb)) useStore.setState({ workbook: apply() });
  }, 0);
}

/** For scripts/check-ai.ts: swap the server call and clock, and inspect or reset the state. */
export const __aiTest = {
  setSender: (fn: Sender) => (sender = fn),
  setClock: (fn: () => number) => (now = fn),
  reset: () => {
    answers.clear();
    failures.clear();
    inflight.clear();
    queue.clear();
    cellKeys.clear();
    cellErrors.clear();
    seenValues.clear();
    firstBook.clear();
    clearTimeout(timer);
    timer = undefined;
    running = false;
  },
  state: () => ({ answers: new Map(answers), queued: [...queue.keys()], inflight: [...inflight], failed: [...failures.keys()] }),
};
