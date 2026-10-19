// Live formula evaluation for the grid preview.
// The .xlsx keeps real formulas; this only computes what the preview shows.

import FormulaParser from "fast-formula-parser";
import * as formulajs from "@formulajs/formulajs";
import type { CellValue, SheetSpec, WorkbookSpec } from "./types";
import { formatValue, numberStyleAt } from "./format";
import { AI_LOADING, aiValuesStale, lookupAi, type AiValueUpdate } from "./aiFormula";

export type Computed = CellValue | { error: string };

const CYCLE = { error: "#CYCLE!" };

function errorOf(v: unknown): { error: string } | null {
  if (v && typeof v === "object" && "_error" in v) return { error: String((v as { _error: string })._error) };
  return null;
}

// fast-formula-parser leaves many functions as empty stubs (MAX, SUMIFS, AVERAGEIFS, ...) and lacks
// others (COUNTIFS, XLOOKUP, PMT, ...). Fill those gaps from formula.js (MIT).
type Arg = { value: unknown };
type PArg = Arg & {
  omitted?: boolean;
  ref?: { sheet?: string; row?: number; col?: number; from?: { row: number; col: number }; to?: { row: number; col: number } };
};
const fj = formulajs as unknown as Record<string, unknown>;

function lookupFj(name: string): ((...a: unknown[]) => unknown) | undefined {
  const [base, sub] = name.split(".");
  const top = fj[base] as unknown;
  const fn = sub ? (top as Record<string, unknown> | undefined)?.[sub] : top;
  return typeof fn === "function" ? (fn as (...a: unknown[]) => unknown) : undefined;
}

const clean = (v: unknown): unknown => (Array.isArray(v) ? v.map(clean) : v === undefined ? null : v);

const EXTRA_FUNCTIONS: Record<string, (...args: Arg[]) => unknown> = (() => {
  const builtin = (new FormulaParser({}) as unknown as { functions: Record<string, () => unknown> }).functions;
  const isStub = (f: () => unknown) => /^\(\)=>\{\}$/.test(f.toString().replace(/\s/g, ""));
  const names = new Set<string>();
  for (const [n, f] of Object.entries(builtin)) if (isStub(f)) names.add(n);
  for (const n of Object.keys(fj)) if (/^[A-Z][A-Z0-9]*$/.test(n) && !(n in builtin)) names.add(n);
  const out: Record<string, (...args: Arg[]) => unknown> = {};
  for (const name of names) {
    const fn = lookupFj(name);
    if (!fn) continue;
    out[name] = (...args: Arg[]) => {
      const res = fn(...args.map((a) => clean(a.value)));
      if (res instanceof Error) throw new FormulaParser.FormulaError(res.message || "#VALUE!");
      if (res === undefined) throw new FormulaParser.FormulaError("#N/A");
      return res;
    };
  }
  // XLOOKUP(lookup, lookup_array, return_array, [if_not_found]): exact match, the common case.
  out.XLOOKUP = (lookup: Arg, keys: Arg, values: Arg, notFound?: Arg) => {
    const flat = (a: unknown) => (Array.isArray(a) ? (a as unknown[]).flat(Infinity) : [a]);
    const k = flat(keys.value);
    const vals = flat(values.value);
    const want = typeof lookup.value === "string" ? lookup.value.toLowerCase() : lookup.value;
    const i = k.findIndex((x) => (typeof x === "string" ? x.toLowerCase() : x) === want);
    if (i >= 0) return vals[i] ?? 0;
    if (notFound) return notFound.value;
    throw new FormulaParser.FormulaError("#N/A");
  };
  return out;
})();

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

/** Raw value at Excel (row, col), both 1-based. Row 1 is the header row. Dates become Excel serial numbers. */
function rawAt(sheet: SheetSpec, row: number, col: number): CellValue | undefined {
  if (row === 1) return sheet.columns[col - 1]?.header;
  const cell = sheet.rows[row - 2]?.cells[col - 1];
  const v = cell?.value ?? undefined;
  if (typeof v === "string" && (cell?.format ?? sheet.columns[col - 1]?.format) === "date") {
    const m = ISO_DATE.exec(v);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
  }
  return v;
}

const cache = new WeakMap<WorkbookSpec, Map<string, Computed>>();

/** evaluateWorkbook, memoised per workbook object (the store replaces it on every change). */
export function evaluateCached(wb: WorkbookSpec): Map<string, Computed> {
  let out = cache.get(wb);
  if (!out) {
    out = evaluateWorkbook(wb);
    cache.set(wb, out);
  }
  return out;
}

/**
 * A copy of `wb` with these cells' ai_value set, sharing every untouched sheet, row and cell.
 * ai_value doesn't change what formulas compute, so the copy reuses `wb`'s evaluation.
 */
export function withAiValues(wb: WorkbookSpec, updates: AiValueUpdate[]): WorkbookSpec {
  const sheets = wb.sheets.slice();
  const touched = new Set<string>();
  for (const u of updates) {
    const src = sheets[u.sheet];
    if (!src?.rows[u.row]?.cells[u.col]) continue;
    let sheet = src;
    if (!touched.has(`${u.sheet}`)) {
      sheet = sheets[u.sheet] = { ...src, rows: src.rows.slice() };
      touched.add(`${u.sheet}`);
    }
    if (!touched.has(`${u.sheet}:${u.row}`)) {
      sheet.rows[u.row] = { ...sheet.rows[u.row], cells: sheet.rows[u.row].cells.slice() };
      touched.add(`${u.sheet}:${u.row}`);
    }
    sheet.rows[u.row].cells[u.col] = { ...sheet.rows[u.row].cells[u.col], ai_value: u.value };
  }
  const next = { ...wb, sheets };
  const done = cache.get(wb);
  if (done) cache.set(next, done);
  return next;
}

/** The value a cell shows: its literal, or its formula's result. Grid coordinates (row 0 = header). */
export function valueAt(wb: WorkbookSpec, sheetIndex: number, r: number, c: number): Computed | undefined {
  const sheet = wb.sheets[sheetIndex];
  if (r === 0) return sheet.columns[c]?.header;
  const raw = sheet.rows[r - 1]?.cells[c]?.value;
  if (typeof raw === "string" && raw.startsWith("=")) return evaluateCached(wb).get(`${sheetIndex}:${r + 1}:${c + 1}`);
  return raw ?? undefined;
}

/**
 * Evaluate every formula in the workbook. Returns a lookup keyed by
 * `${sheetIndex}:${row}:${col}` (Excel 1-based coordinates).
 */
export function evaluateWorkbook(wb: WorkbookSpec): Map<string, Computed> {
  const byName = new Map(wb.sheets.map((s, i) => [s.name.toLowerCase(), i]));
  const cache = new Map<string, Computed>();
  const visiting = new Set<string>();

  const sheetIndex = (name: string | undefined, fallback: number) =>
    name === undefined ? fallback : byName.get(name.toLowerCase());

  // =AI(...) bookkeeping: which formula is being computed, and whether it used an AI answer.
  type Frame = { ai: boolean; pending: boolean; failed: boolean };
  const frames: Frame[] = [];
  const stale: AiValueUpdate[] = [];

  const valueOf = (si: number, row: number, col: number): Computed | undefined => {
    const sheet = wb.sheets[si];
    if (!sheet) return { error: "#REF!" };
    const raw = rawAt(sheet, row, col);
    if (!(typeof raw === "string" && raw.startsWith("="))) return raw ?? undefined;

    const key = `${si}:${row}:${col}`;
    if (cache.has(key)) return cache.get(key);
    if (visiting.has(key)) return CYCLE;
    visiting.add(key);
    let result: Computed;
    const frame: Frame = { ai: false, pending: false, failed: false };
    frames.push(frame);
    try {
      // The parser isn't re-entrant: a formula that reads another formula's cell needs its own instance.
      const parser = (parsers[depth] ??= makeParser());
      depth++;
      let out: unknown;
      try {
        out = parser.parse(raw.slice(1), { sheet: sheet.name, row, col });
      } finally {
        depth--;
      }
      const err = errorOf(out);
      result = err ?? (Array.isArray(out) ? (out[0]?.[0] ?? null) : (out as CellValue));
    } catch (e) {
      result = errorOf(e) ?? { error: "#ERROR!" };
    }
    frames.pop();
    if (frame.ai && !frame.pending && !frame.failed && (result === null || typeof result !== "object")) {
      // Keep the cell's ai_value in step with what it shows, so downloads have it.
      const shown = result === null ? "" : typeof result === "boolean" ? (result ? "TRUE" : "FALSE") : String(result);
      const cell = sheet.rows[row - 2]?.cells[col - 1];
      if (cell && (cell.ai_value ?? null) !== shown) stale.push({ sheet: si, row: row - 2, col: col - 1, value: shown });
    }
    visiting.delete(key);
    cache.set(key, result);
    return result;
  };

  // The parser wants plain values; errors propagate as FormulaError-like objects.
  const toParser = (v: Computed | undefined) => {
    if (v && typeof v === "object") return new FormulaParser.FormulaError(v.error);
    return v ?? undefined;
  };

  /** What an argument of AI() shows in the grid: a cell's formatted text, a range's texts joined by ", ". */
  const shownArg = (a: PArg, here: string): string => {
    const text = (v: unknown, sheet: SheetSpec | undefined, r: number, c: number): string => {
      if (v instanceof FormulaParser.FormulaError) throw v;
      if (v === null || v === undefined) return "";
      if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
      if (typeof v === "number" && sheet && r >= 1) return formatValue(v, numberStyleAt(sheet, r, c)).text;
      if (typeof v === "number") return formatValue(v, "text").text;
      return String(v);
    };
    const ref = a.ref;
    const sheet = ref ? wb.sheets[sheetIndex(ref.sheet ?? here, -1) ?? -1] : undefined;
    if (ref?.from && Array.isArray(a.value)) {
      const parts: string[] = [];
      (a.value as unknown[][]).forEach((line, i) =>
        (Array.isArray(line) ? line : [line]).forEach((v, j) => {
          const t = text(v, sheet, ref.from!.row - 1 + i, ref.from!.col - 1 + j);
          if (t !== "") parts.push(t);
        }),
      );
      return parts.join(", ");
    }
    if (Array.isArray(a.value)) return (a.value as unknown[]).flat(Infinity).map((v) => text(v, undefined, 0, 0)).filter(Boolean).join(", ");
    return ref?.row ? text(a.value, sheet, ref.row - 1, (ref.col ?? 1) - 1) : text(a.value, undefined, 0, 0);
  };

  // =AI("instruction", input1, ...): answered from the cache, otherwise "Loading..." while it's fetched in the background.
  const aiFunction = (ctx: { position?: { sheet: string; row: number; col: number } }, ...args: PArg[]) => {
    const pos = ctx.position;
    if (!args.length || args[0].omitted) throw new FormulaParser.FormulaError("#VALUE!");
    const here = pos?.sheet ?? wb.sheets[0]?.name;
    const [prompt, ...inputs] = args.map((a) => shownArg(a, here));
    if (!prompt.trim()) throw new FormulaParser.FormulaError("#VALUE!");
    const frame = frames[frames.length - 1];
    if (frame) frame.ai = true;
    const si = sheetIndex(here, -1) ?? -1;
    const cell = pos ? wb.sheets[si]?.rows[pos.row - 2]?.cells[pos.col - 1] : undefined;
    const res = lookupAi(wb, prompt, inputs, pos && si >= 0 ? { sheet: wb.sheets[si].name, row: pos.row, col: pos.col, aiValue: cell?.ai_value ?? null } : null);
    if (res.status === "loading") {
      if (frame) frame.pending = true;
      return AI_LOADING;
    }
    if (res.status === "error") {
      if (frame) frame.failed = true;
      throw new FormulaParser.FormulaError("#AI!");
    }
    // Plain numbers come back as numbers, so totals over AI cells work.
    return /^-?\d+(\.\d+)?$/.test(res.value.trim()) ? parseFloat(res.value) : res.value;
  };

  const parsers: InstanceType<typeof FormulaParser>[] = [];
  let depth = 0;
  const makeParser = () => {
    const p = new FormulaParser({
    functions: EXTRA_FUNCTIONS,
    functionsNeedContext: { AI: aiFunction },
    onCell: ({ sheet, row, col }: { sheet: string; row: number; col: number }) => {
      const si = sheetIndex(sheet, -1);
      if (si === undefined || si < 0) return new FormulaParser.FormulaError("#REF!");
      return toParser(valueOf(si, row, col));
    },
    onRange: (ref: { sheet: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
      const si = sheetIndex(ref.sheet, -1);
      if (si === undefined || si < 0) return new FormulaParser.FormulaError("#REF!");
      const sheet = wb.sheets[si];
      const lastRow = Math.min(ref.to.row, sheet.rows.length + 1);
      const lastCol = Math.min(ref.to.col, Math.max(sheet.columns.length, 1));
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= lastRow; r++) {
        const line: unknown[] = [];
        for (let c = ref.from.col; c <= lastCol; c++) line.push(toParser(valueOf(si, r, c)));
        out.push(line);
      }
      return out;
    },
  });
    // AI() needs to know which cells its inputs came from, to show them formatted.
    (p as unknown as { funsPreserveRef: string[] }).funsPreserveRef.push("AI");
    return p;
  };

  const results = new Map<string, Computed>();
  wb.sheets.forEach((sheet, si) => {
    sheet.rows.forEach((row, ri) => {
      row.cells.forEach((cell, ci) => {
        if (typeof cell.value === "string" && cell.value.startsWith("=")) {
          results.set(`${si}:${ri + 2}:${ci + 1}`, valueOf(si, ri + 2, ci + 1) ?? null);
        }
      });
    });
  });
  if (stale.length) aiValuesStale(wb, () => withAiValues(wb, stale));
  return results;
}
