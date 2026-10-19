// "Clean up data" tools: pure functions over the workbook model, used by CleanupDialog.
// Every operation acts on data rows only (subtotal / total / section rows are left alone)
// and never changes a formula cell (a string starting with "=").

import { CURRENCIES, columnLetter, parseInput } from "./format";
import { offsetFormula, shiftFormula, type Axis } from "./refs";
import type { CellSpec, CellValue, NumberFormat, Selection, SheetSpec, WorkbookSpec } from "./types";

export type CaseMode = "upper" | "lower" | "title";
export type FillMode = "above" | "value" | "zero";

export type CleanupOp =
  | { kind: "dedupe" }
  | { kind: "trim" }
  | { kind: "case"; mode: CaseMode }
  /** Split one column (0-based index) by `delimiter`; " " splits on any run of whitespace. */
  | { kind: "split"; column: number; delimiter: string }
  | { kind: "dates"; dayFirst: boolean }
  | { kind: "fill"; mode: FillMode; value?: string }
  | { kind: "numbers" };

export interface CleanupResult {
  /** Cells changed, or rows removed for "dedupe". */
  changed: number;
  /** Cells that looked like candidates but could not be converted (dates / numbers). */
  skipped: number;
  /** Columns added by "split". */
  addedColumns: number;
}

export interface CleanupScope {
  /** 0-based column indices to act on. */
  cols: number[];
  /** True when the selection was a single cell, so the whole sheet is used. */
  whole: boolean;
}

export const isFormula = (v: unknown): boolean => typeof v === "string" && v.startsWith("=");
const isBlank = (v: CellValue | undefined) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

/** Which columns a selection covers; a single cell means every column of the sheet. */
export function cleanupScope(sheet: SheetSpec, sel: Selection): CleanupScope {
  const c1 = Math.min(sel.c1, sel.c2);
  const c2 = Math.max(sel.c1, sel.c2);
  const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
  const last = Math.max(sheet.columns.length - 1, 0);
  if (single) return { cols: Array.from({ length: sheet.columns.length }, (_, i) => i), whole: true };
  const cols: number[] = [];
  for (let c = c1; c <= Math.min(c2, last); c++) cols.push(c);
  return { cols, whole: false };
}

/** "B", "B–D" or "A, C" style label for a list of columns. */
export function columnsLabel(cols: number[]): string {
  if (!cols.length) return "none";
  const contiguous = cols.every((c, i) => i === 0 || c === cols[i - 1] + 1);
  if (contiguous && cols.length > 1) return `${columnLetter(cols[0])}–${columnLetter(cols[cols.length - 1])}`;
  return cols.map(columnLetter).join(", ");
}

/** The cell at (row, c), created if the row is shorter. */
function cellOf(row: { cells: CellSpec[] }, c: number): CellSpec {
  while (row.cells.length <= c) row.cells.push({ value: null });
  return row.cells[c];
}

function dataRows(sheet: SheetSpec) {
  return sheet.rows.map((row, i) => ({ row, i })).filter(({ row }) => (row.kind ?? "data") === "data");
}

/** Apply `fn` to every non-formula text value in the target cells; returns how many changed. */
function mapText(sheet: SheetSpec, cols: number[], fn: (s: string) => string): number {
  let n = 0;
  for (const { row } of dataRows(sheet)) {
    for (const c of cols) {
      const cell = row.cells[c];
      if (!cell || typeof cell.value !== "string" || isFormula(cell.value)) continue;
      const next = fn(cell.value);
      if (next !== cell.value) {
        cell.value = next === "" ? null : next;
        n++;
      }
    }
  }
  return n;
}

/**
 * Keep formulas, scenarios and sheet-level ranges pointing at the same cells after a row or column
 * is inserted / deleted on `target` (the same bookkeeping the store does for Insert / Delete).
 * `at` is Excel 1-based; delta +1 inserts before `at`, -1 deletes `at`.
 */
export function shiftReferences(wb: WorkbookSpec, target: string, axis: Axis, at: number, delta: 1 | -1) {
  for (const sheet of wb.sheets)
    for (const row of sheet.rows)
      for (const cell of row.cells)
        if (typeof cell.value === "string" && isFormula(cell.value)) cell.value = shiftFormula(cell.value, sheet.name, target, axis, at, delta);
  if (wb.scenarios?.length) {
    wb.scenarios = wb.scenarios.map((s) => {
      const values: Record<string, number> = {};
      for (const [k, v] of Object.entries(s.values)) {
        const moved = shiftFormula(`=${k}`, "", target, axis, at, delta).slice(1);
        if (!moved.includes("#REF!")) values[moved] = v;
      }
      return { ...s, values };
    });
  }
  const sheet = wb.sheets.find((s) => s.name.toLowerCase() === target.toLowerCase());
  if (!sheet) return;
  const move = (range: string) => shiftFormula(`=${range}`, sheet.name, sheet.name, axis, at, delta).slice(1);
  if (sheet.validations) sheet.validations = sheet.validations.map((v) => ({ ...v, range: move(v.range) })).filter((v) => !v.range.includes("#REF!"));
  if (sheet.conditional_formats)
    sheet.conditional_formats = sheet.conditional_formats.map((v) => ({ ...v, range: move(v.range) })).filter((v) => !v.range.includes("#REF!"));
  if (sheet.merges) sheet.merges = sheet.merges.map(move).filter((m) => !m.includes("#REF!"));
  if (axis === "col" && sheet.filters?.length) {
    sheet.filters = sheet.filters
      .filter((f) => !(delta === -1 && f.column === at - 1))
      .map((f) => (f.column >= at - 1 ? { ...f, column: f.column + delta } : f));
  }
  if (sheet.charts)
    sheet.charts = sheet.charts
      .map((c) => ({
        ...c,
        categories: move(c.categories),
        series: c.series.map((x) => ({ ...x, range: move(x.range) })).filter((x) => !x.range.includes("#REF!")),
      }))
      .filter((c) => c.series.length && !c.categories.includes("#REF!"));
}

// ---------------------------------------------------------------- duplicates

const keyPart = (v: CellValue | undefined) => {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return `s:${v.trim().replace(/\s+/g, " ").toLowerCase()}`;
  return `${typeof v}:${v}`;
};

/** Grid-row indices (into sheet.rows) of data rows that repeat an earlier row in `cols`. Rows blank in every column are never duplicates. */
export function duplicateRows(sheet: SheetSpec, cols: number[]): number[] {
  const seen = new Set<string>();
  const out: number[] = [];
  for (const { row, i } of dataRows(sheet)) {
    const parts = cols.map((c) => keyPart(row.cells[c]?.value));
    if (parts.every((p) => p === "" || p === "s:")) continue;
    const key = parts.join("\u0001");
    if (seen.has(key)) out.push(i);
    else seen.add(key);
  }
  return out;
}

function removeDuplicates(wb: WorkbookSpec, si: number, cols: number[]): number {
  const sheet = wb.sheets[si];
  const dupes = duplicateRows(sheet, cols);
  for (const i of [...dupes].reverse()) {
    sheet.rows.splice(i, 1);
    shiftReferences(wb, sheet.name, "row", i + 2, -1);
  }
  return dupes.length;
}

// ---------------------------------------------------------------- text

export const trimText = (s: string) => s.replace(/[ \t ]+/g, " ").replace(/^\s+|\s+$/g, "");

export function changeCaseText(s: string, mode: CaseMode): string {
  if (mode === "upper") return s.toUpperCase();
  if (mode === "lower") return s.toLowerCase();
  return s.toLowerCase().replace(/(^|[\s\-(/"])(\p{L})/gu, (_, pre: string, ch: string) => pre + ch.toUpperCase());
}

// ---------------------------------------------------------------- split

export function splitText(s: string, delimiter: string): string[] {
  if (delimiter === " ") return s.trim().split(/\s+/).filter(Boolean);
  if (!delimiter) return [s];
  return s.split(delimiter).map((p) => p.trim());
}

const splitValue = (p: string): CellValue => (p === "" ? null : /^-?\d+(\.\d+)?$/.test(p) && !/^-?0\d/.test(p) ? parseFloat(p) : p);

function splitColumn(wb: WorkbookSpec, si: number, col: number, delimiter: string): { changed: number; added: number } {
  const sheet = wb.sheets[si];
  const pieces = new Map<number, string[]>();
  let parts = 1;
  for (const { row, i } of dataRows(sheet)) {
    const v = row.cells[col]?.value;
    if (typeof v !== "string" || isFormula(v)) continue;
    const p = splitText(v, delimiter);
    if (p.length > 1) {
      pieces.set(i, p);
      parts = Math.max(parts, p.length);
    }
  }
  const added = parts - 1;
  if (!added) return { changed: 0, added: 0 };
  const base = sheet.columns[col];
  const header = base?.header || `Column ${columnLetter(col)}`;
  const newCols = Array.from({ length: added }, (_, k) => ({ header: `${header} ${k + 2}`, format: "text" as NumberFormat, width: base?.width ?? null }));
  sheet.columns.splice(col + 1, 0, ...newCols);
  for (const row of sheet.rows) if (row.cells.length > col + 1) row.cells.splice(col + 1, 0, ...newCols.map(() => ({ value: null })));
  for (let k = 0; k < added; k++) shiftReferences(wb, sheet.name, "col", col + 2, 1);
  if (sheet.freeze_columns && sheet.freeze_columns > col + 1) sheet.freeze_columns += added;
  for (const [i, p] of pieces) {
    const row = sheet.rows[i];
    cellOf(row, col).value = splitValue(p[0]);
    for (let k = 1; k < p.length; k++) {
      const src = row.cells[col];
      const cell = cellOf(row, col + k);
      cell.value = splitValue(p[k]);
      // Carry the look of the original cell, but not a number/date format meant for the whole text.
      for (const key of ["bold", "italic", "underline", "font_color", "fill_color", "align", "border"] as const)
        if (src[key] !== undefined) (cell as Record<string, unknown>)[key] = src[key];
    }
  }
  return { changed: pieces.size, added };
}

// ---------------------------------------------------------------- dates

const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
function monthOf(word: string): number {
  const w = word.toLowerCase().replace(/\.$/, "");
  if (w.length < 3) return 0;
  const i = MONTHS.findIndex((m) => m.startsWith(w));
  return i + 1;
}
const pad = (n: number) => String(n).padStart(2, "0");
const fullYear = (y: string) => (y.length === 2 ? (+y < 50 ? 2000 + +y : 1900 + +y) : +y);

function iso(y: number, m: number, d: number): string | null {
  if (!(y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1)) return null;
  const days = new Date(Date.UTC(y, m, 0)).getUTCDate();
  if (d > days) return null;
  return `${y}-${pad(m)}-${pad(d)}`;
}

/**
 * Read a text date in a common format and return it the way the app stores dates: an ISO
 * "YYYY-MM-DD" string in a date-formatted cell. Numeric dates like 03/05/2024 are read day-first
 * or month-first per `dayFirst`, unless one part is over 12. Returns null when it isn't a date.
 */
export function parseDateText(input: string, dayFirst: boolean): string | null {
  const s = input
    .trim()
    .replace(/^(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/i, "")
    .replace(/(\d)(st|nd|rd|th)\b/gi, "$1")
    .replace(/\s+/g, " ");
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);
  if (m) return iso(+m[1], +m[2], +m[3]);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})$/.exec(s);
  if (m) {
    const a = +m[1];
    const b = +m[2];
    const [d, mo] = a > 12 && b <= 12 ? [a, b] : b > 12 && a <= 12 ? [b, a] : dayFirst ? [a, b] : [b, a];
    return iso(fullYear(m[3]), mo, d);
  }
  m = /^(\d{1,2})[ \-/.]?([a-z]{3,}\.?),?[ \-/.](\d{4}|\d{2})$/i.exec(s);
  if (m && monthOf(m[2])) return iso(fullYear(m[3]), monthOf(m[2]), +m[1]);
  m = /^([a-z]{3,}\.?)[ \-/.](\d{1,2}),?[ \-/.](\d{4}|\d{2})$/i.exec(s);
  if (m && monthOf(m[1])) return iso(fullYear(m[3]), monthOf(m[1]), +m[2]);
  m = /^(\d{4})[ \-/.]([a-z]{3,}\.?)[ \-/.](\d{1,2})$/i.exec(s);
  if (m && monthOf(m[2])) return iso(+m[1], monthOf(m[2]), +m[3]);
  return null;
}

function fixDates(sheet: SheetSpec, cols: number[], dayFirst: boolean): { changed: number; skipped: number } {
  let changed = 0;
  let skipped = 0;
  for (const c of cols) {
    let colChanged = 0;
    const col = sheet.columns[c];
    for (const { row } of dataRows(sheet)) {
      const cell = row.cells[c];
      if (!cell || typeof cell.value !== "string" || isFormula(cell.value) || cell.value.trim() === "") continue;
      const date = parseDateText(cell.value, dayFirst);
      if (!date) {
        if (/\d/.test(cell.value)) skipped++;
        continue;
      }
      const fmt = cell.format ?? col?.format ?? "text";
      if (date === cell.value && fmt === "date") continue;
      cell.value = date;
      if (cell.format && cell.format !== "date") cell.format = null;
      colChanged++;
    }
    if (colChanged && col) col.format = "date";
    changed += colChanged;
  }
  return { changed, skipped };
}

// ---------------------------------------------------------------- numbers

const SYMBOLS = [...CURRENCIES].map((c) => ({ code: c.code, symbol: c.symbol.trim() })).sort((a, b) => b.symbol.length - a.symbol.length);

/**
 * Read text that looks like a number: "1,200", "$5", "-€3.50", "(500)", "12%", "1 200". Returns
 * null for anything else, including codes with leading zeros like "00123" (zip codes, IDs).
 */
export function parseNumberText(input: string): { value: number; format?: NumberFormat; currency?: string } | null {
  let t = input.trim().replace(/[\s ]/g, "");
  if (!t) return null;
  let negative = false;
  if (/^\(.+\)$/.test(t)) {
    negative = true;
    t = t.slice(1, -1);
  }
  const takeSign = () => {
    if (t[0] === "-" || t[0] === "+") {
      if (t[0] === "-") negative = !negative;
      t = t.slice(1);
    }
  };
  takeSign();
  let currency: string | undefined;
  for (const { code, symbol } of SYMBOLS) {
    if (symbol === "R") continue; // a bare "R" prefix is too ambiguous (R5 is more likely a code)
    if (t.startsWith(symbol)) {
      currency = code;
      t = t.slice(symbol.length);
      break;
    }
    if (t.endsWith(symbol) && t.length > symbol.length) {
      currency = code;
      t = t.slice(0, -symbol.length);
      break;
    }
  }
  if (currency) takeSign();
  const percent = t.endsWith("%");
  if (percent) t = t.slice(0, -1);
  if (!/^(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$/.test(t) || !/\d/.test(t)) return null;
  if (/^0\d/.test(t)) return null;
  let value = parseFloat(t.replace(/,/g, ""));
  if (percent) value /= 100;
  if (negative) value = -value;
  if (percent) return { value, format: "percent" };
  if (currency) return { value, format: "currency", currency };
  return { value };
}

function convertNumbers(sheet: SheetSpec, cols: number[]): number {
  let n = 0;
  for (const { row } of dataRows(sheet)) {
    for (const c of cols) {
      const cell = row.cells[c];
      if (!cell || typeof cell.value !== "string" || isFormula(cell.value)) continue;
      const num = parseNumberText(cell.value);
      if (!num) continue;
      cell.value = num.value;
      const colFmt = sheet.columns[c]?.format ?? "text";
      if (num.format && (cell.format ?? colFmt) !== num.format) cell.format = num.format;
      if (num.currency && num.currency !== "USD" && !cell.currency) cell.currency = num.currency;
      if (!num.format && cell.format === "text") cell.format = null;
      n++;
    }
  }
  return n;
}

// ---------------------------------------------------------------- blanks

function fillBlanks(sheet: SheetSpec, cols: number[], mode: FillMode, text?: string): number {
  let n = 0;
  const fixed: CellValue = mode === "zero" ? 0 : parseInput(text ?? "");
  if (mode === "value" && (fixed === null || fixed === undefined)) return 0;
  for (const c of cols) {
    let above: { cell: CellSpec; i: number } | null = null;
    for (let i = 0; i < sheet.rows.length; i++) {
      const row = sheet.rows[i];
      if ((row.kind ?? "data") !== "data") {
        above = null; // don't carry values across a subtotal / section break
        continue;
      }
      const existing = row.cells[c];
      if (existing && !isBlank(existing.value)) {
        above = { cell: existing, i };
        continue;
      }
      if (mode === "above") {
        if (!above) continue;
        const src = above.cell;
        const cell = cellOf(row, c);
        cell.value = typeof src.value === "string" && isFormula(src.value) ? offsetFormula(src.value, i - above.i, 0) : src.value ?? null;
        if (src.format) cell.format = src.format;
        if (src.decimals != null) cell.decimals = src.decimals;
        if (src.currency) cell.currency = src.currency;
        above = { cell, i };
      } else {
        cellOf(row, c).value = fixed;
      }
      n++;
    }
  }
  return n;
}

// ---------------------------------------------------------------- entry point

/** Run one clean-up operation on sheet `si`, in place. Run it on a structuredClone for a preview. */
export function applyCleanup(wb: WorkbookSpec, si: number, op: CleanupOp, cols: number[]): CleanupResult {
  const sheet = wb.sheets[si];
  const none = { changed: 0, skipped: 0, addedColumns: 0 };
  if (!sheet) return none;
  const valid = cols.filter((c) => c >= 0 && c < sheet.columns.length);
  switch (op.kind) {
    case "dedupe":
      return { ...none, changed: valid.length ? removeDuplicates(wb, si, valid) : 0 };
    case "trim":
      return { ...none, changed: mapText(sheet, valid, trimText) };
    case "case":
      return { ...none, changed: mapText(sheet, valid, (s) => changeCaseText(s, op.mode)) };
    case "split": {
      if (op.column < 0 || op.column >= sheet.columns.length) return none;
      const r = splitColumn(wb, si, op.column, op.delimiter);
      return { ...none, changed: r.changed, addedColumns: r.added };
    }
    case "dates": {
      const r = fixDates(sheet, valid, op.dayFirst);
      return { ...none, changed: r.changed, skipped: r.skipped };
    }
    case "fill":
      return { ...none, changed: fillBlanks(sheet, valid, op.mode, op.value) };
    case "numbers":
      return { ...none, changed: convertNumbers(sheet, valid) };
  }
}

const plural = (n: number, one: string, many = `${one}s`) => `${n.toLocaleString("en-US")} ${n === 1 ? one : many}`;

/** "Will remove 12 duplicate rows" (or the past tense once applied). */
export function describeCleanup(op: CleanupOp, r: CleanupResult, done = false): string {
  const w = (future: string, past: string) => (done ? past : future);
  if (!r.changed) {
    const nothing: Record<CleanupOp["kind"], string> = {
      dedupe: "No duplicate rows found.",
      trim: "No extra spaces found.",
      case: "Nothing to change.",
      split: "No cells contain that delimiter.",
      dates: r.skipped ? `No text dates found (${plural(r.skipped, "cell")} not recognised).` : "No text dates found.",
      fill: "No blank cells found.",
      numbers: "No numbers stored as text found.",
    };
    return nothing[op.kind];
  }
  switch (op.kind) {
    case "dedupe":
      return `${w("Will remove", "Removed")} ${plural(r.changed, "duplicate row")} (the first of each is kept).`;
    case "trim":
      return `${w("Will trim", "Trimmed")} ${plural(r.changed, "cell")}.`;
    case "case":
      return `${w("Will change", "Changed")} ${plural(r.changed, "cell")} to ${op.mode === "upper" ? "UPPER" : op.mode === "lower" ? "lower" : "Title"} case.`;
    case "split":
      return `${w("Will split", "Split")} ${plural(r.changed, "cell")} into ${r.addedColumns + 1} columns (${plural(r.addedColumns, "new column")} ${w("inserted", "inserted")} to the right).`;
    case "dates":
      return `${w("Will convert", "Converted")} ${plural(r.changed, "cell")} to dates${r.skipped ? `; ${plural(r.skipped, "cell")} not recognised` : ""}.`;
    case "fill":
      return `${w("Will fill", "Filled")} ${plural(r.changed, "blank cell")}.`;
    case "numbers":
      return `${w("Will convert", "Converted")} ${plural(r.changed, "cell")} to numbers.`;
  }
}
