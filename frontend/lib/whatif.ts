// What-if analysis: scenarios and goal seek. Pure functions over the workbook model, used by WhatIfDialog.
// Cell addresses are "Sheet!B2" strings (sheet name + A1 reference), the same keys ScenarioSpec uses.

import { columnLetter } from "./format";
import { evaluateWorkbook, type Computed } from "./formula";
import type { CellSpec, ScenarioSpec, WorkbookSpec } from "./types";

export interface CellAddr {
  /** Sheet index. */
  si: number;
  /** Excel row (1 = header row, 2 = sheet.rows[0]). */
  row: number;
  /** 0-based column index. */
  col: number;
}

const colIndex = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** Quote a sheet name for a reference when it needs it ("My Sheet" -> 'My Sheet'). */
export function sheetPrefix(name: string): string {
  return /^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Parse "B2", "Income!B2", "'My Sheet'!$B$2". Sheet defaults to `defaultSheet`. Header-row refs are rejected. */
export function parseCellRef(wb: WorkbookSpec, ref: string, defaultSheet = 0): CellAddr | null {
  const m = /^\s*(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)\s*$/.exec(ref);
  if (!m) return null;
  const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const si = name === undefined ? defaultSheet : wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
  if (si < 0 || !wb.sheets[si]) return null;
  const row = parseInt(m[4], 10);
  if (row < 2) return null;
  return { si, row, col: colIndex(m[3]) };
}

export function formatCellRef(wb: WorkbookSpec, a: CellAddr): string {
  return `${sheetPrefix(wb.sheets[a.si].name)}!${columnLetter(a.col)}${a.row}`;
}

/** Canonical "Sheet!B2" form of a reference, or null if it doesn't point at a data cell. */
export function normalizeRef(wb: WorkbookSpec, ref: string, defaultSheet = 0): string | null {
  const a = parseCellRef(wb, ref, defaultSheet);
  return a ? formatCellRef(wb, a) : null;
}

/** Split "B2, C3; Income!D4" into refs. */
export function splitRefs(text: string): string[] {
  return text
    .split(/[,;\n]+/)
    .map((s) => s.trim())
    .filter(Boolean);
}

export function rawValue(wb: WorkbookSpec, a: CellAddr) {
  return wb.sheets[a.si]?.rows[a.row - 2]?.cells[a.col]?.value ?? null;
}

export const isFormulaValue = (v: unknown) => typeof v === "string" && v.startsWith("=");

/** Set a cell's literal value, creating the row / cell if needed. */
export function setCellValue(wb: WorkbookSpec, a: CellAddr, value: number) {
  const sheet = wb.sheets[a.si];
  while (sheet.columns.length <= a.col) sheet.columns.push({ header: `Column ${sheet.columns.length + 1}`, format: "text" });
  while (sheet.rows.length < a.row - 1) sheet.rows.push({ kind: "data", cells: [] });
  const row = sheet.rows[a.row - 2];
  while (row.cells.length <= a.col) row.cells.push({ value: null } as CellSpec);
  row.cells[a.col].value = value;
}

/** The value a cell shows, given an evaluation map from evaluateWorkbook / evaluateCached. */
export function valueIn(wb: WorkbookSpec, values: Map<string, Computed>, a: CellAddr): Computed | undefined {
  const raw = rawValue(wb, a);
  if (isFormulaValue(raw)) return values.get(`${a.si}:${a.row}:${a.col + 1}`);
  return raw;
}

// ---------------------------------------------------------------- scenarios

/** Write a scenario's input values into the workbook (in place). Returns how many cells were set. */
export function applyScenario(wb: WorkbookSpec, scenario: ScenarioSpec): number {
  let n = 0;
  for (const [ref, v] of Object.entries(scenario.values)) {
    const a = parseCellRef(wb, ref);
    if (!a || isFormulaValue(rawValue(wb, a)) || !Number.isFinite(v)) continue;
    setCellValue(wb, a, v);
    n++;
  }
  return n;
}

/** Result cells' values under each scenario: out[scenario][result]. Each scenario is computed on a copy. */
export function scenarioResults(wb: WorkbookSpec, scenarios: ScenarioSpec[], results: string[]): (Computed | undefined)[][] {
  const addrs = results.map((r) => parseCellRef(wb, r));
  return scenarios.map((s) => {
    const copy = structuredClone(wb);
    applyScenario(copy, s);
    const values = evaluateWorkbook(copy);
    return addrs.map((a) => (a ? valueIn(copy, values, a) : { error: "#REF!" }));
  });
}

// ---------------------------------------------------------------- goal seek

export interface GoalSeekResult {
  found: boolean;
  /** The input value found (or the closest one tried). */
  value: number;
  /** The result cell's value at `value`. */
  result: number | null;
  iterations: number;
  message: string;
}

/** Why goal seek can't run on these cells, or null when it can. */
export function goalSeekProblem(wb: WorkbookSpec, resultRef: string, inputRef: string): string | null {
  const r = parseCellRef(wb, resultRef);
  const i = parseCellRef(wb, inputRef);
  if (!r) return "The result cell isn't a valid cell reference (for example Sheet1!B10).";
  if (!i) return "The changing cell isn't a valid cell reference (for example Sheet1!B2).";
  if (!isFormulaValue(rawValue(wb, r))) return "The result cell must contain a formula.";
  const v = rawValue(wb, i);
  if (isFormulaValue(v)) return "The changing cell must hold a number, not a formula.";
  if (typeof v !== "number") return "The changing cell must hold a number.";
  if (r.si === i.si && r.row === i.row && r.col === i.col) return "The result cell and the changing cell must be different.";
  return null;
}

/**
 * Find the value of `inputRef` that makes `resultRef` equal `target`, like Excel's Goal Seek.
 * Secant steps first; once a sign change is seen, false position with bisection fallback inside
 * the bracket; if the secant stalls without a bracket, search outward for one. Stops at a
 * relative tolerance of `tol` (|f| <= tol * max(1, |target|)) or after `maxIter` iterations.
 */
export function goalSeek(wb: WorkbookSpec, resultRef: string, target: number, inputRef: string, maxIter = 60, tol = 1e-7): GoalSeekResult {
  const problem = goalSeekProblem(wb, resultRef, inputRef);
  const r = parseCellRef(wb, resultRef);
  const inp = parseCellRef(wb, inputRef);
  if (problem || !r || !inp) return { found: false, value: NaN, result: null, iterations: 0, message: problem ?? "Invalid cells." };

  const work = structuredClone(wb);
  const key = `${r.si}:${r.row}:${r.col + 1}`;
  let evals = 0;
  const f = (x: number): number => {
    evals++;
    setCellValue(work, inp, x);
    const v = evaluateWorkbook(work).get(key);
    return typeof v === "number" && Number.isFinite(v) ? v - target : NaN;
  };
  const eps = tol * Math.max(1, Math.abs(target));
  const x0 = rawValue(wb, inp) as number;

  const f0 = f(x0);
  let best = { x: x0, fx: f0 };
  const consider = (x: number, fx: number) => {
    if (Number.isFinite(fx) && (!Number.isFinite(best.fx) || Math.abs(fx) < Math.abs(best.fx))) best = { x, fx };
  };
  const done = (found: boolean, message: string): GoalSeekResult => {
    let { x, fx } = best;
    // Prefer a tidier number when it is just as good: 249.99999998 -> 250.
    if (found)
      for (const p of [4, 6, 8, 10, 12]) {
        const tidy = Number(x.toPrecision(p));
        if (tidy === x) break;
        const ft = f(tidy);
        if (Number.isFinite(ft) && Math.abs(ft) <= eps) {
          x = tidy;
          fx = ft;
          break;
        }
      }
    return { found, value: x, result: Number.isFinite(fx) ? fx + target : null, iterations: evals, message };
  };
  if (Math.abs(best.fx) <= eps) return done(true, "The result already equals the target.");

  // Bracket [lo, hi] with f(lo), f(hi) of opposite sign, once one is known.
  let lo: { x: number; fx: number } | null = null;
  let hi: { x: number; fx: number } | null = null;
  const noteBracket = (a: { x: number; fx: number }, b: { x: number; fx: number }) => {
    if (Number.isFinite(a.fx) && Number.isFinite(b.fx) && Math.sign(a.fx) !== Math.sign(b.fx)) {
      lo = a.x < b.x ? a : b;
      hi = a.x < b.x ? b : a;
    }
  };

  // Phase 1: secant.
  let p0 = best;
  const step0 = x0 === 0 ? 1 : Math.abs(x0) * 0.01;
  let p1 = { x: x0 + step0, fx: f(x0 + step0) };
  consider(p1.x, p1.fx);
  noteBracket(p0, p1);
  while (!lo && evals < maxIter / 2) {
    if (!Number.isFinite(p1.fx) || !Number.isFinite(p0.fx) || p1.fx === p0.fx) break;
    const x2 = p1.x - (p1.fx * (p1.x - p0.x)) / (p1.fx - p0.fx);
    if (!Number.isFinite(x2) || Math.abs(x2) > 1e15) break;
    const p2 = { x: x2, fx: f(x2) };
    consider(p2.x, p2.fx);
    if (Math.abs(p2.fx) <= eps) return done(true, "Found a solution.");
    noteBracket(p1, p2);
    p0 = p1;
    p1 = p2;
  }

  // Phase 2: no sign change yet, search outward from the start for one.
  if (!lo) {
    const start = { x: x0, fx: f0 };
    let d = Math.max(Math.abs(x0), 1) * 0.1;
    let prevUp = start;
    let prevDown = start;
    while (!lo && evals < maxIter && d < 1e15) {
      const up = { x: x0 + d, fx: f(x0 + d) };
      consider(up.x, up.fx);
      noteBracket(prevUp, up);
      if (Number.isFinite(up.fx)) prevUp = up;
      if (lo) break;
      const down = { x: x0 - d, fx: f(x0 - d) };
      consider(down.x, down.fx);
      noteBracket(prevDown, down);
      if (Number.isFinite(down.fx)) prevDown = down;
      d *= 2;
    }
    if (Math.abs(best.fx) <= eps) return done(true, "Found a solution.");
  }

  // Phase 3: false position (Illinois) inside the bracket, bisecting when it stalls.
  if (lo && hi) {
    let a: { x: number; fx: number } = lo;
    let b: { x: number; fx: number } = hi;
    let side = 0;
    let budget = maxIter;
    while (budget-- > 0) {
      let x = b.x - (b.fx * (b.x - a.x)) / (b.fx - a.fx);
      if (!Number.isFinite(x) || x <= Math.min(a.x, b.x) || x >= Math.max(a.x, b.x)) x = (a.x + b.x) / 2;
      const fx = f(x);
      if (!Number.isFinite(fx)) {
        // Undefined here (e.g. #DIV/0!): bisect instead.
        const m = (a.x + b.x) / 2;
        const fm = f(m);
        if (!Number.isFinite(fm)) break;
        consider(m, fm);
        if (Math.sign(fm) === Math.sign(a.fx)) a = { x: m, fx: fm };
        else b = { x: m, fx: fm };
        continue;
      }
      consider(x, fx);
      if (Math.abs(fx) <= eps) return done(true, "Found a solution.");
      if (Math.sign(fx) === Math.sign(b.fx)) {
        b = { x, fx };
        if (side === -1) a = { x: a.x, fx: a.fx / 2 };
        side = -1;
      } else {
        a = { x, fx };
        if (side === 1) b = { x: b.x, fx: b.fx / 2 };
        side = 1;
      }
      if (Math.abs(b.x - a.x) <= 1e-15 * Math.max(1, Math.abs(a.x))) break;
    }
    if (Math.abs(best.fx) <= eps) return done(true, "Found a solution.");
    return done(false, "No solution found: the result jumps past the target without reaching it.");
  }
  return done(false, "No solution found: no value of the changing cell reached the target.");
}
