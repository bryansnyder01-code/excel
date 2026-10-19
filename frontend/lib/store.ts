"use client";

import { create } from "zustand";
import { fillSeries } from "./fill";
import { offsetFormula, renameSheetInFormula, shiftFormula, type Axis } from "./refs";
import type {
  CellSpec,
  ConditionalFormatSpec,
  CellValue,
  ChartSpec,
  ChatMessage,
  Checkpoint,
  Issue,
  NumberFormat,
  RowKind,
  SavedDoc,
  Selection,
  SheetSpec,
  WorkbookSpec,
} from "./types";
import { evaluateCached } from "./formula";
import { columnLetter, defaultDecimals, numberStyleAt, typedFormat } from "./format";
import { parseRange } from "./cellStyle";

const HISTORY_LIMIT = 50;
export const CHECKPOINT_LIMIT = 15;

export const newId = () =>
  typeof crypto !== "undefined" && "randomUUID" in crypto ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
export const INVALID_SHEET_CHARS = /[[\]:*?/\\]/;

export function blankSheet(name = "Sheet1"): SheetSpec {
  return {
    name,
    columns: ["A", "B", "C", "D", "E"].map((h) => ({ header: `Column ${h}`, format: "text" as NumberFormat })),
    rows: [],
  };
}

export function blankWorkbook(): WorkbookSpec {
  return { title: "Untitled workbook", sheets: [blankSheet()] };
}

/** True when the workbook has no content yet, so the next prompt should build from scratch. */
export function isBlank(wb: WorkbookSpec): boolean {
  return wb.sheets.every((s) => s.rows.every((r) => r.cells.every((c) => c.value === null || c.value === undefined || c.value === "")));
}

const clone = <T,>(v: T): T => structuredClone(v);
const isFormula = (v: unknown): v is string => typeof v === "string" && v.startsWith("=");

function normalize(sel: Selection): Selection {
  return {
    r1: Math.min(sel.r1, sel.r2),
    r2: Math.max(sel.r1, sel.r2),
    c1: Math.min(sel.c1, sel.c2),
    c2: Math.max(sel.c1, sel.c2),
  };
}

/** Make sure grid cell (r, c) exists in the sheet (r >= 1 is a data row). */
function ensureCell(sheet: SheetSpec, r: number, c: number): CellSpec {
  while (sheet.columns.length <= c) sheet.columns.push({ header: `Column ${sheet.columns.length + 1}`, format: "text" });
  while (sheet.rows.length < r) sheet.rows.push({ kind: "data", cells: [] });
  const row = sheet.rows[r - 1];
  while (row.cells.length <= c) row.cells.push({ value: null });
  return row.cells[c];
}

function forEachFormula(wb: WorkbookSpec, fn: (f: string, sheet: SheetSpec) => string) {
  for (const sheet of wb.sheets)
    for (const row of sheet.rows)
      for (const cell of row.cells) if (isFormula(cell.value)) cell.value = fn(cell.value, sheet);
}

/** Excel A1 range for a grid selection, data rows only (grid row r is Excel row r + 1). */
export function selectionRange(sel: Selection): string {
  const s = normalize(sel);
  const r1 = Math.max(1, s.r1) + 1;
  const r2 = Math.max(1, s.r2) + 1;
  return `${columnLetter(s.c1)}${r1}:${columnLetter(s.c2)}${r2}`;
}

interface State {
  docId: string;
  checkpoints: Checkpoint[];
  workbook: WorkbookSpec;
  active: number;
  selection: Selection;
  past: WorkbookSpec[];
  future: WorkbookSpec[];
  messages: ChatMessage[];
  busy: boolean;
  /** What the AI is doing right now, streamed from the server. */
  progress: string | null;
  /** An AI change waiting for Accept or Discard, and the workbook from before it. */
  review: Review | null;

  /** Apply a change to a copy of the workbook and record it for undo. */
  mutate: (fn: (wb: WorkbookSpec) => void) => void;
  replaceWorkbook: (wb: WorkbookSpec) => void;
  newWorkbook: () => void;
  undo: () => void;
  redo: () => void;

  setActive: (i: number) => void;
  select: (sel: Selection) => void;

  /** `typed` is the text as entered, so "15%" or "€9" also sets the cell's number format. */
  setCellValue: (r: number, c: number, value: CellValue, typed?: string) => void;
  clearSelection: () => void;
  toggleStyle: (key: "bold" | "italic" | "underline" | "wrap") => void;
  styleSelection: (patch: Partial<CellSpec>) => void;
  setNumberFormat: (fmt: NumberFormat) => void;
  setRowKind: (kind: RowKind) => void;

  insertRow: (below: boolean) => void;
  deleteRows: () => void;
  insertColumn: (right: boolean) => void;
  deleteColumns: () => void;

  addSheet: () => void;
  renameSheet: (i: number, name: string) => string | null;
  deleteSheet: (i: number) => void;
  setTitle: (title: string) => void;

  addMessage: (m: Omit<ChatMessage, "id">) => void;
  setBusy: (b: boolean) => void;

  /** Take a new workbook from the AI, an import, a template or a summary, as a restorable version. */
  applyResult: (wb: WorkbookSpec, label: string, reply?: { text: string; issues?: Issue[] }, focusLast?: boolean) => void;
  saveCheckpoint: (label: string) => string;
  restoreCheckpoint: (id: string) => void;
  loadDocument: (doc: SavedDoc) => void;

  setDropdown: (options: string[] | null) => void;
  addChart: (chart: ChartSpec) => void;
  removeChart: (index: number) => void;

  /** Write a block of cells with its top-left corner at grid (r, c). */
  pasteCells: (r: number, c: number, cells: CellSpec[][]) => void;
  /** Extend the selection's contents into `target` (which contains the selection) like the fill handle. */
  fill: (target: Selection) => void;
  /** Ctrl+D / Ctrl+R: copy the first row (or column) of the selection across the rest. */
  fillFromEdge: (axis: "row" | "col") => void;
  setColumnWidth: (c: number, width: number | null) => void;
  setFreeze: (header: boolean, columns: number) => void;
  changeDecimals: (delta: 1 | -1) => void;
  setCurrency: (code: string) => void;
  setNote: (r: number, c: number, note: string | null) => void;
  setConditionalFormats: (sheetIndex: number, rules: ConditionalFormatSpec[]) => void;
  /** Replace every match of `find` in the given cells, all in one undo step. */
  replaceIn: (targets: { sheet: number; r: number; c: number }[], find: RegExp, replacement: string) => number;

  /** Sort each block of data rows by column `c` (total and section rows stay put). */
  sortBy: (c: number, descending: boolean) => void;
  /** Show only rows whose column `c` displays one of `values`; null clears the filter. */
  setFilter: (c: number, values: string[] | null) => void;
  /** Merge the selection, or unmerge if it touches merged cells. */
  toggleMerge: () => void;
  /** Row height in points for grid row r (0 = header); null for automatic. */
  setRowHeight: (r: number, pt: number | null) => void;

  setProgress: (text: string | null) => void;
  /** Show an AI result in the grid, highlighted, until it's accepted or discarded. */
  startReview: (wb: WorkbookSpec, label: string, reply: { text: string; issues?: Issue[] }) => void;
  acceptReview: () => void;
  discardReview: () => void;
}

export interface Review {
  before: WorkbookSpec;
  label: string;
  reply: { text: string; issues?: Issue[] };
}

export const useStore = create<State>((set, get) => ({
  docId: newId(),
  checkpoints: [],
  workbook: blankWorkbook(),
  active: 0,
  selection: { r1: 1, c1: 0, r2: 1, c2: 0 },
  past: [],
  future: [],
  messages: [],
  busy: false,
  progress: null,
  review: null,

  mutate: (fn) => {
    // Editing by hand while an AI change is on screen keeps that change.
    if (get().review) get().acceptReview();
    const { workbook, past } = get();
    const next = clone(workbook);
    fn(next);
    if (!editGuard(workbook, next)) return;
    set({ workbook: next, past: [...past, workbook].slice(-HISTORY_LIMIT), future: [] });
  },

  replaceWorkbook: (wb) => {
    const { workbook, past, active } = get();
    set({
      workbook: wb,
      past: [...past, workbook].slice(-HISTORY_LIMIT),
      future: [],
      active: Math.min(active, wb.sheets.length - 1),
    });
  },

  newWorkbook: () =>
    set({
      docId: newId(),
      checkpoints: [],
      workbook: blankWorkbook(),
      active: 0,
      past: [],
      future: [],
      messages: [],
      review: null,
      selection: { r1: 1, c1: 0, r2: 1, c2: 0 },
    }),

  undo: () => {
    if (get().review) return get().discardReview();
    const { past, workbook, future, active } = get();
    if (!past.length) return;
    const prev = past[past.length - 1];
    set({ workbook: prev, past: past.slice(0, -1), future: [workbook, ...future], active: Math.min(active, prev.sheets.length - 1) });
  },

  redo: () => {
    const { past, workbook, future, active } = get();
    if (!future.length) return;
    const next = future[0];
    set({ workbook: next, past: [...past, workbook], future: future.slice(1), active: Math.min(active, next.sheets.length - 1) });
  },

  setActive: (i) => set({ active: i, selection: { r1: 1, c1: 0, r2: 1, c2: 0 } }),
  select: (sel) => set({ selection: sel }),

  setCellValue: (r, c, value, typed) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      if (r === 0) {
        ensureCell(sheet, 1, c); // grows columns
        sheet.columns[c].header = value === null ? "" : String(value);
        return;
      }
      const implied = typed !== undefined && typeof value === "number" ? typedFormat(typed) : null;
      const now = implied ? numberStyleAt(sheet, r, c) : null;
      const cell = ensureCell(sheet, r, c);
      cell.value = value;
      if (!isAiFormula(value)) delete cell.ai_value;
      if (implied && now && (now.fmt !== implied.format || (implied.format === "currency" && (now.currency ?? "USD") !== (implied.currency ?? "USD")))) {
        cell.format = implied.format;
        if (implied.format === "currency") cell.currency = implied.currency ?? null;
      }
    }),

  clearSelection: () => {
    const s = normalize(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      for (let r = Math.max(1, s.r1); r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) {
          const cell = sheet.rows[r - 1]?.cells[c];
          if (cell) cell.value = null;
        }
    });
  },

  toggleStyle: (key) => {
    const s = normalize(get().selection);
    const sheet = get().workbook.sheets[get().active];
    let all = true;
    for (let r = Math.max(1, s.r1); r <= s.r2; r++)
      for (let c = s.c1; c <= s.c2; c++) if (!sheet.rows[r - 1]?.cells[c]?.[key]) all = false;
    get().styleSelection({ [key]: !all });
  },

  styleSelection: (patch) => {
    const s = normalize(get().selection);
    if (s.r2 < 1) return; // header row has fixed styling
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      for (let r = Math.max(1, s.r1); r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) Object.assign(ensureCell(sheet, r, c), patch);
    });
  },

  setNumberFormat: (fmt) => {
    const s = normalize(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      if (s.r1 === 0) {
        // Selection starts at the header: format the whole column(s).
        for (let c = s.c1; c <= s.c2; c++) {
          ensureCell(sheet, 1, c);
          sheet.columns[c].format = fmt;
          for (const row of sheet.rows) if (row.cells[c]) row.cells[c].format = null;
        }
        return;
      }
      for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) ensureCell(sheet, r, c).format = fmt;
    });
  },

  setRowKind: (kind) => {
    const s = normalize(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      for (let r = Math.max(1, s.r1); r <= s.r2; r++) {
        ensureCell(sheet, r, 0);
        sheet.rows[r - 1].kind = kind;
      }
    });
  },

  insertRow: (below) => {
    const s = normalize(get().selection);
    const at = below ? Math.max(1, s.r2) + 1 : Math.max(1, s.r1); // grid row of the new row
    const sheet = get().workbook.sheets[get().active].name;
    structural([{ op: "insert_rows", sheet, at, count: 1 }]);
    set({ selection: { r1: at, r2: at, c1: s.c1, c2: s.c1 } });
  },

  deleteRows: () => {
    const s = normalize(get().selection);
    const first = Math.max(1, s.r1);
    const sheet = get().workbook.sheets[get().active];
    const last = Math.min(s.r2, sheet.rows.length);
    if (last >= first) structural([{ op: "delete_rows", sheet: sheet.name, at: first, count: last - first + 1 }]);
    set({ selection: { r1: first, r2: first, c1: s.c1, c2: s.c1 } });
  },

  insertColumn: (right) => {
    const s = normalize(get().selection);
    const at = right ? s.c2 + 1 : s.c1; // 0-based column index of the new column
    const sheet = get().workbook.sheets[get().active].name;
    structural([{ op: "insert_cols", sheet, at, count: 1 }]);
    set({ selection: { r1: 0, r2: 0, c1: at, c2: at } });
  },

  deleteColumns: () => {
    const s = normalize(get().selection);
    const sheet = get().workbook.sheets[get().active];
    const last = Math.min(s.c2, sheet.columns.length - 1);
    const count = Math.min(last - s.c1 + 1, sheet.columns.length - 1);
    if (count > 0) structural([{ op: "delete_cols", sheet: sheet.name, at: s.c1, count }]);
    set({ selection: { r1: s.r1, r2: s.r1, c1: s.c1, c2: s.c1 } });
  },

  addSheet: () => {
    const names = new Set(get().workbook.sheets.map((s) => s.name.toLowerCase()));
    let n = get().workbook.sheets.length + 1;
    while (names.has(`sheet${n}`)) n++;
    get().mutate((wb) => {
      wb.sheets.push(blankSheet(`Sheet${n}`));
    });
    get().setActive(get().workbook.sheets.length - 1);
  },

  /** Returns an error message, or null on success. */
  renameSheet: (i, raw) => {
    const name = raw.trim();
    const { workbook } = get();
    const old = workbook.sheets[i].name;
    if (name === old) return null;
    if (!name) return "Sheet names can't be empty.";
    if (name.length > 31) return "Sheet names can be at most 31 characters.";
    if (INVALID_SHEET_CHARS.test(name)) return "Sheet names can't contain [ ] : * ? / \\";
    if (workbook.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase()))
      return "Another sheet already has that name.";
    get().mutate((wb) => {
      wb.sheets[i].name = name;
      forEachFormula(wb, (f) => renameSheetInFormula(f, old, name));
    });
    return null;
  },

  deleteSheet: (i) => {
    if (get().workbook.sheets.length <= 1) return;
    get().mutate((wb) => {
      wb.sheets.splice(i, 1);
    });
    get().setActive(Math.max(0, Math.min(get().active, get().workbook.sheets.length - 1)));
  },

  setTitle: (title) =>
    get().mutate((wb) => {
      wb.title = title || "Untitled workbook";
    }),

  addMessage: (m) => set({ messages: [...get().messages, { ...m, id: newId() }] }),
  setBusy: (b) => set({ busy: b }),

  applyResult: (wb, label, reply, focusLast) => {
    get().replaceWorkbook(wb);
    const id = get().saveCheckpoint(label);
    if (reply) get().addMessage({ role: "assistant", text: reply.text, issues: reply.issues, checkpointId: id });
    if (focusLast) get().setActive(wb.sheets.length - 1);
  },

  saveCheckpoint: (label) => {
    const cp: Checkpoint = { id: newId(), label, at: Date.now(), workbook: get().workbook };
    set({ checkpoints: [...get().checkpoints, cp].slice(-CHECKPOINT_LIMIT) });
    return cp.id;
  },

  restoreCheckpoint: (id) => {
    const cp = get().checkpoints.find((c) => c.id === id);
    if (!cp) return;
    get().replaceWorkbook(clone(cp.workbook));
    get().saveCheckpoint(`Restored "${cp.label}"`);
  },

  loadDocument: (doc) =>
    set({
      docId: doc.id,
      workbook: doc.workbook,
      messages: doc.messages ?? [],
      checkpoints: doc.checkpoints ?? [],
      active: 0,
      past: [],
      future: [],
      review: null,
      selection: { r1: 1, c1: 0, r2: 1, c2: 0 },
    }),

  setDropdown: (options) => {
    const range = selectionRange(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      const others = (sheet.validations ?? []).filter((v) => v.range !== range);
      sheet.validations = options && options.length ? [...others, { range, type: "list", options }] : others;
    });
  },

  addChart: (chart) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      sheet.charts = [...(sheet.charts ?? []), chart];
    }),

  removeChart: (index) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      sheet.charts = (sheet.charts ?? []).filter((_, i) => i !== index);
    }),

  pasteCells: (r0, c0, cells) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      cells.forEach((line, dr) =>
        line.forEach((src, dc) => {
          const r = r0 + dr;
          const c = c0 + dc;
          if (r === 0) {
            ensureCell(sheet, 1, c);
            sheet.columns[c].header = src.value === null || src.value === undefined ? "" : String(src.value);
            return;
          }
          const cell = ensureCell(sheet, r, c);
          for (const k of Object.keys(cell)) delete (cell as Record<string, unknown>)[k];
          Object.assign(cell, structuredClone(src));
        }),
      );
    }),

  fill: (target) => {
    const s = normalize(get().selection);
    const t = normalize(target);
    const down = t.r2 > s.r2;
    const up = t.r1 < s.r1;
    const right = t.c2 > s.c2;
    const left = t.c1 < s.c1;
    if (!(down || up || right || left)) return;
    const vertical = down || up;
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      const at = (r: number, c: number): CellSpec => structuredClone(sheet.rows[r - 1]?.cells[c] ?? { value: null });
      if (vertical) {
        const first = Math.max(1, s.r1);
        for (let c = s.c1; c <= s.c2; c++) {
          const src: CellSpec[] = [];
          for (let r = first; r <= s.r2; r++) src.push(at(r, c));
          if (up) src.reverse();
          const n = down ? t.r2 - s.r2 : s.r1 - Math.max(1, t.r1);
          fillSeries(src, n, "row", down ? 1 : -1).forEach((cell, i) => {
            const r = down ? s.r2 + 1 + i : s.r1 - 1 - i;
            if (r >= 1) Object.assign(ensureCell(sheet, r, c), cell);
          });
        }
      } else {
        for (let r = Math.max(1, s.r1); r <= s.r2; r++) {
          const src: CellSpec[] = [];
          for (let c = s.c1; c <= s.c2; c++) src.push(at(r, c));
          if (left) src.reverse();
          const n = right ? t.c2 - s.c2 : s.c1 - t.c1;
          fillSeries(src, n, "col", right ? 1 : -1).forEach((cell, i) => {
            const c = right ? s.c2 + 1 + i : s.c1 - 1 - i;
            if (c >= 0) Object.assign(ensureCell(sheet, r, c), cell);
          });
        }
      }
    });
    set({ selection: { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 } });
  },

  fillFromEdge: (axis) => {
    const s = normalize(get().selection);
    const r1 = Math.max(1, s.r1);
    if (axis === "row" ? s.r2 <= r1 : s.c2 <= s.c1) return;
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      for (let r = r1; r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) {
          if (axis === "row" ? r === r1 : c === s.c1) continue;
          const src = structuredClone(sheet.rows[(axis === "row" ? r1 : r) - 1]?.cells[axis === "row" ? c : s.c1] ?? { value: null });
          if (isFormula(src.value)) src.value = offsetFormula(src.value, axis === "row" ? r - r1 : 0, axis === "col" ? c - s.c1 : 0);
          Object.assign(ensureCell(sheet, r, c), src);
        }
    });
  },

  setColumnWidth: (c, width) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      ensureCell(sheet, 1, c);
      sheet.columns[c].width = width === null ? null : Math.max(3, Math.min(100, Math.round(width * 10) / 10));
    }),

  setFreeze: (header, columns) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      sheet.freeze_header = header;
      sheet.freeze_columns = Math.max(0, columns);
    }),

  changeDecimals: (delta) => {
    const s = normalize(get().selection);
    const sheet0 = get().workbook.sheets[get().active];
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      if (s.r1 === 0) {
        for (let c = s.c1; c <= s.c2; c++) {
          ensureCell(sheet, 1, c);
          const col = sheet.columns[c];
          const fmt = col.format === "integer" || col.format === "text" || !col.format ? "decimal" : col.format;
          const cur = col.decimals ?? (col.format === "integer" ? 0 : defaultDecimals(fmt, col.currency));
          col.format = fmt === "date" ? "date" : fmt;
          col.decimals = Math.max(0, Math.min(10, cur + delta));
        }
        return;
      }
      for (let r = s.r1; r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) {
          const ns = numberStyleAt(sheet0, r, c);
          if (ns.fmt === "date") continue;
          const cell = ensureCell(sheet, r, c);
          const fmt = ns.fmt === "integer" || ns.fmt === "text" ? "decimal" : ns.fmt;
          const cur = ns.decimals ?? (ns.fmt === "integer" || ns.fmt === "text" ? 0 : defaultDecimals(fmt, ns.currency));
          if (fmt !== ns.fmt) cell.format = fmt;
          cell.decimals = Math.max(0, Math.min(10, cur + delta));
        }
    });
  },

  setCurrency: (code) => {
    const s = normalize(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      if (s.r1 === 0) {
        for (let c = s.c1; c <= s.c2; c++) {
          ensureCell(sheet, 1, c);
          Object.assign(sheet.columns[c], { format: "currency", currency: code, decimals: null });
          for (const row of sheet.rows) if (row.cells[c]) Object.assign(row.cells[c], { format: null, currency: null, decimals: null });
        }
        return;
      }
      for (let r = s.r1; r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) Object.assign(ensureCell(sheet, r, c), { format: "currency", currency: code, decimals: null });
    });
  },

  setNote: (r, c, note) => {
    if (r === 0) return;
    get().mutate((wb) => {
      ensureCell(wb.sheets[get().active], r, c).note = note?.trim() ? note.trim() : null;
    });
  },

  setConditionalFormats: (si, rules) =>
    get().mutate((wb) => {
      wb.sheets[si].conditional_formats = rules;
    }),

  replaceIn: (targets, find, replacement) => {
    let n = 0;
    get().mutate((wb) => {
      for (const t of targets) {
        const sheet = wb.sheets[t.sheet];
        if (t.r === 0) {
          const col = sheet.columns[t.c];
          if (!col) continue;
          const next = col.header.replace(find, replacement);
          if (next !== col.header) (col.header = next), n++;
          continue;
        }
        const cell = sheet.rows[t.r - 1]?.cells[t.c];
        if (!cell || cell.value === null || cell.value === undefined) continue;
        const before = String(cell.value);
        const after = before.replace(find, replacement);
        if (after === before) continue;
        n++;
        // Typed-in text that becomes a number stays a number, like Excel.
        cell.value = typeof cell.value === "number" && /^-?\d+(\.\d+)?$/.test(after) ? parseFloat(after) : after;
        if (!isAiFormula(cell.value)) delete cell.ai_value;
      }
    });
    return n;
  },


  sortBy: (c, descending) => {
    const { workbook, active } = get();
    const computed = evaluateCached(workbook);
    const src = workbook.sheets[active];
    const key = (rowIndex: number) => {
      const cell = src.rows[rowIndex]?.cells[c];
      const v = typeof cell?.value === "string" && cell.value.startsWith("=") ? computed.get(`${active}:${rowIndex + 2}:${c + 1}`) : cell?.value;
      return v === null || v === undefined || v === "" || typeof v === "object" ? null : v;
    };
    const cmp = (a: number, b: number) => {
      const x = key(a);
      const y = key(b);
      if (x === null || y === null) return x === y ? 0 : x === null ? 1 : -1; // blanks last either way
      const d = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: "base" });
      return descending ? -d : d;
    };
    get().mutate((wb) => {
      const sheet = wb.sheets[active];
      const original = sheet.rows;
      const out = [...original];
      // Sort each run of data rows on its own, so totals stay under the rows they add up.
      let start = 0;
      while (start < original.length) {
        if ((original[start].kind ?? "data") !== "data") {
          start++;
          continue;
        }
        let end = start;
        while (end < original.length && (original[end].kind ?? "data") === "data") end++;
        const order = Array.from({ length: end - start }, (_, i) => start + i).sort((a, b) => cmp(a, b) || a - b);
        order.forEach((from, i) => {
          const row = structuredClone(original[from]);
          const shift = start + i - from;
          // Formulas move with their row, as in Excel.
          for (const cell of row.cells) if (isFormula(cell.value) && shift) cell.value = offsetFormula(cell.value, shift, 0);
          out[start + i] = row;
        });
        start = end;
      }
      sheet.rows = out;
    });
  },

  setFilter: (c, values) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      const rest = (sheet.filters ?? []).filter((f) => f.column !== c);
      sheet.filters = values ? [...rest, { column: c, values }] : rest;
    }),

  toggleMerge: () => {
    const s = normalize(get().selection);
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      const merges = sheet.merges ?? [];
      const touching = merges.filter((m) => {
        const g = parseRange(m);
        return g && g.r1 - 1 <= s.r2 && g.r2 - 1 >= s.r1 && g.c1 - 1 <= s.c2 && g.c2 - 1 >= s.c1;
      });
      if (touching.length) {
        sheet.merges = merges.filter((m) => !touching.includes(m));
        return;
      }
      if (s.r1 === s.r2 && s.c1 === s.c2) return;
      sheet.merges = [...merges, `${columnLetter(s.c1)}${s.r1 + 1}:${columnLetter(s.c2)}${s.r2 + 1}`];
      // Like Excel, only the top-left value survives.
      for (let r = Math.max(1, s.r1); r <= s.r2; r++)
        for (let c = s.c1; c <= s.c2; c++) {
          if (r === s.r1 && c === s.c1) continue;
          const cell = sheet.rows[r - 1]?.cells[c];
          if (cell) cell.value = null;
        }
    });
  },

  setRowHeight: (r, pt) =>
    get().mutate((wb) => {
      const sheet = wb.sheets[get().active];
      const h = pt === null ? null : Math.max(9, Math.min(400, Math.round(pt * 4) / 4));
      if (r === 0) sheet.header_height = h;
      else ensureCell(sheet, r, 0) && (sheet.rows[r - 1].height = h);
    }),

  setProgress: (text) => set({ progress: text }),

  startReview: (wb, label, reply) => {
    const { workbook, past, active } = get();
    set({
      review: { before: workbook, label, reply },
      workbook: wb,
      past: [...past, workbook].slice(-HISTORY_LIMIT),
      future: [],
      active: Math.min(active, wb.sheets.length - 1),
    });
  },

  acceptReview: () => {
    const review = get().review;
    if (!review) return;
    set({ review: null });
    const id = get().saveCheckpoint(review.label);
    get().addMessage({ role: "assistant", text: review.reply.text, issues: review.reply.issues, checkpointId: id });
  },

  discardReview: () => {
    const review = get().review;
    if (!review) return;
    const { past, active } = get();
    set({
      review: null,
      workbook: review.before,
      past: past.at(-1) === review.before ? past.slice(0, -1) : past,
      active: Math.min(active, review.before.sheets.length - 1),
    });
    get().addMessage({ role: "assistant", text: `Discarded: ${review.reply.text}` });
  },
}));

const isAiFormula = (v: unknown) => typeof v === "string" && v.startsWith("=") && /\bAI\s*\(/i.test(v);

let editGuard: (before: WorkbookSpec, after: WorkbookSpec) => boolean = () => true;
/** Live editing installs a check that refuses changes the signed-in role may not make. */
export function setEditGuard(fn: typeof editGuard) {
  editGuard = fn;
}

/** Inserting or deleting rows or columns: the one kind of change that moves other cells. */
export interface StructuralOp {
  op: "insert_rows" | "delete_rows" | "insert_cols" | "delete_cols";
  sheet: string;
  /** Rows: the grid row (1 = first data row) where rows go in or come out. Columns: 0-based index. */
  at: number;
  count: number;
}

let opListener: (ops: StructuralOp[]) => void = () => {};
/** Live editing hears about each local structural change, so it can send the operation, not a whole workbook. */
export function onStructuralOps(fn: (ops: StructuralOp[]) => void) {
  opListener = fn;
}

function structural(ops: StructuralOp[]) {
  const before = useStore.getState().workbook;
  useStore.getState().mutate((wb) => ops.forEach((o) => applyOp(wb, o)));
  if (useStore.getState().workbook !== before) opListener(ops);
}

/** Apply one structural change in place, moving every reference to the cells it shifts, like Excel. */
export function applyOp(wb: WorkbookSpec, o: StructuralOp) {
  const sheet = wb.sheets.find((s) => s.name === o.sheet);
  if (!sheet || o.count < 1) return;
  if (o.op === "insert_rows") {
    while (sheet.rows.length < o.at - 1) sheet.rows.push({ kind: "data", cells: [] });
    sheet.rows.splice(o.at - 1, 0, ...Array.from({ length: o.count }, () => ({ kind: "data" as const, cells: [] })));
    for (let i = 0; i < o.count; i++) shiftAll(wb, sheet.name, "row", o.at + 1, 1);
  } else if (o.op === "delete_rows") {
    for (let r = o.at + o.count - 1; r >= Math.max(1, o.at); r--) {
      if (r - 1 >= sheet.rows.length) continue;
      sheet.rows.splice(r - 1, 1);
      shiftAll(wb, sheet.name, "row", r + 1, -1);
    }
  } else if (o.op === "insert_cols") {
    while (sheet.columns.length < o.at) sheet.columns.push({ header: `Column ${sheet.columns.length + 1}`, format: "text" });
    for (let i = 0; i < o.count; i++) {
      sheet.columns.splice(o.at, 0, { header: "New column", format: "text" });
      for (const row of sheet.rows) if (row.cells.length > o.at) row.cells.splice(o.at, 0, { value: null });
      shiftAll(wb, sheet.name, "col", o.at + 1, 1);
    }
  } else {
    for (let c = o.at + o.count - 1; c >= o.at; c--) {
      if (c >= sheet.columns.length || sheet.columns.length <= 1) continue;
      sheet.columns.splice(c, 1);
      for (const row of sheet.rows) if (row.cells.length > c) row.cells.splice(c, 1);
      shiftAll(wb, sheet.name, "col", c + 1, -1);
    }
  }
}

function shiftAll(wb: WorkbookSpec, target: string, axis: Axis, at: number, delta: 1 | -1) {
  forEachFormula(wb, (f, sheet) => shiftFormula(f, sheet.name, target, axis, at, delta));
  // Chart and dropdown ranges live on their own sheet and move with its rows/columns.
  const sheet = wb.sheets.find((s) => s.name === target);
  if (!sheet) return;
  const move = (range: string) => shiftFormula(`=${range}`, target, target, axis, at, delta).slice(1);
  sheet.validations = (sheet.validations ?? [])
    .map((v) => ({ ...v, range: move(v.range) }))
    .filter((v) => !v.range.includes("#REF!"));
  sheet.conditional_formats = (sheet.conditional_formats ?? [])
    .map((v) => ({ ...v, range: move(v.range) }))
    .filter((v) => !v.range.includes("#REF!"));
  sheet.merges = (sheet.merges ?? []).map(move).filter((m) => !m.includes("#REF!"));
  if (sheet.locked_ranges?.length) sheet.locked_ranges = sheet.locked_ranges.map(move).filter((m) => !m.includes("#REF!"));
  // Inserting or deleting inside the frozen columns grows or shrinks them, like Excel.
  if (axis === "col" && at - 1 < (sheet.freeze_columns ?? 0)) sheet.freeze_columns = Math.max(0, (sheet.freeze_columns ?? 0) + delta);
  // Addresses written as "Sheet!B2" elsewhere in the workbook: scenario inputs and dashboard tiles.
  const moveRef = (ref: string) => {
    const bang = ref.lastIndexOf("!");
    const name = bang < 0 ? "" : ref.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'");
    if (name !== target) return ref;
    const out = move(ref.slice(bang + 1));
    return out.includes("#REF!") ? null : `${ref.slice(0, bang)}!${out}`;
  };
  if (wb.scenarios?.length)
    wb.scenarios = wb.scenarios.map((sc) => ({
      ...sc,
      values: Object.fromEntries(
        Object.entries(sc.values).flatMap(([k, v]) => {
          const m = moveRef(k);
          return m ? [[m, v]] : [];
        }),
      ),
    }));
  if (wb.dashboard?.tiles.length)
    wb.dashboard = {
      ...wb.dashboard,
      tiles: wb.dashboard.tiles.map((t) => ({
        ...t,
        ref: t.ref ? (moveRef(t.ref) ?? "#REF!") : t.ref,
        compare_ref: t.compare_ref ? moveRef(t.compare_ref) : t.compare_ref,
      })),
    };
  if (axis === "col" && sheet.filters?.length) {
    // `at` is 1-based; filters use 0-based columns.
    sheet.filters = sheet.filters
      .filter((f) => !(delta === -1 && f.column === at - 1))
      .map((f) => (f.column >= at - 1 ? { ...f, column: f.column + delta } : f));
  }
  sheet.charts = (sheet.charts ?? [])
    .map((c) => ({
      ...c,
      categories: move(c.categories),
      series: c.series.map((x) => ({ ...x, range: move(x.range) })).filter((x) => !x.range.includes("#REF!")),
    }))
    .filter((c) => c.series.length && !c.categories.includes("#REF!"));
}

export { normalize };
