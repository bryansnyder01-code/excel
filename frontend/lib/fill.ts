// Pure helpers for copy/paste and the fill handle. Kept free of React and the store so
// scripts/check-grid.ts can test them directly.

import { offsetFormula } from "./refs";
import type { CellSpec, CellValue } from "./types";

/** Parse tab-separated text from the clipboard (Excel, Google Sheets, this grid). Handles quoted cells. */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  const s = text.replace(/\r\n?/g, "\n");
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (quoted) {
      if (ch === '"' && s[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else cell += ch;
    } else if (ch === '"' && cell === "") quoted = true;
    else if (ch === "\t") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

export function toTsv(values: string[][]): string {
  const esc = (v: string) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
  return values.map((r) => r.map(esc).join("\t")).join("\n");
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const LISTS = [MONTHS, MONTHS.map((m) => m.slice(0, 3)), DAYS, DAYS.map((d) => d.slice(0, 3)), ["Q1", "Q2", "Q3", "Q4"]];

const ISO = /^(\d{4})-(\d{2})-(\d{2})$/;
const toDays = (iso: string) => {
  const m = ISO.exec(iso)!;
  return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000;
};
const fromDays = (d: number) => new Date(d * 86400000).toISOString().slice(0, 10);

function matchCase(word: string, like: string) {
  if (like === like.toUpperCase()) return word.toUpperCase();
  if (like === like.toLowerCase()) return word.toLowerCase();
  return word;
}

/** Find a known list (months, weekdays, quarters) every source value belongs to, with each value's index. */
function knownList(values: CellValue[]): { list: string[]; idx: number[] } | null {
  for (const list of LISTS) {
    const idx = values.map((v) => (typeof v === "string" ? list.findIndex((x) => x.toLowerCase() === v.toLowerCase()) : -1));
    if (idx.every((i) => i >= 0)) return { list, idx };
  }
  return null;
}

/**
 * Continue a line of source cells (in fill order) by `count` cells, the way Excel's fill handle does:
 * numbers and dates continue their step, "Item 3" counts up, months and weekdays continue,
 * formulas are copied with relative references moved, and everything else repeats.
 * `axis` is the direction of travel; `sign` is +1 for down/right and -1 for up/left.
 */
export function fillSeries(source: CellSpec[], count: number, axis: "row" | "col", sign: 1 | -1): CellSpec[] {
  const k = source.length;
  const values = source.map((c) => c.value ?? null);
  const style = (i: number): CellSpec => {
    const { value: _v, ...rest } = source[i % k];
    return structuredClone(rest);
  };
  const out: CellSpec[] = [];
  const nonEmpty = values.filter((v) => v !== null && v !== "");

  // Formulas anywhere: repeat the pattern, moving references by the distance travelled.
  if (values.some((v) => typeof v === "string" && v.startsWith("="))) {
    for (let p = 0; p < count; p++) {
      const i = p % k;
      const dist = (Math.floor(p / k) + 1) * k * sign;
      const v = values[i];
      const moved = typeof v === "string" && v.startsWith("=") ? offsetFormula(v, axis === "row" ? dist : 0, axis === "col" ? dist : 0) : v;
      out.push({ ...style(i), value: moved });
    }
    return out;
  }

  const step = (nums: number[]) => (nums.length >= 2 ? (nums[nums.length - 1] - nums[0]) / (nums.length - 1) : 0);

  // Numbers: a straight-line series from two or more; a single number repeats, as in Excel.
  if (nonEmpty.length === k && values.every((v) => typeof v === "number")) {
    const nums = values as number[];
    const d = step(nums);
    for (let p = 0; p < count; p++) {
      const v = k === 1 ? nums[0] : nums[k - 1] + d * (p + 1);
      out.push({ ...style(p), value: Math.round(v * 1e10) / 1e10 });
    }
    return out;
  }

  // Dates: one day at a time from a single date, otherwise the same gap.
  if (nonEmpty.length === k && values.every((v) => typeof v === "string" && ISO.test(v))) {
    const days = (values as string[]).map(toDays);
    const d = k === 1 ? 1 : step(days);
    for (let p = 0; p < count; p++) out.push({ ...style(p), value: fromDays(Math.round(days[k - 1] + d * (p + 1))) });
    return out;
  }

  // Months, weekdays, quarters.
  const known = nonEmpty.length === k ? knownList(values) : null;
  if (known) {
    const n = known.list.length;
    const d = k === 1 ? 1 : (((known.idx[k - 1] - known.idx[0]) / (k - 1)) % n + n) % n || 1;
    for (let p = 0; p < count; p++) {
      const idx = (((known.idx[k - 1] + d * (p + 1)) % n) + n) % n;
      out.push({ ...style(p), value: matchCase(known.list[idx], values[0] as string) });
    }
    return out;
  }

  // Text ending in a number ("Item 3", "Week 1"): count on from the last one.
  const tail = /^(.*?)(\d+)$/;
  const parts = values.map((v) => (typeof v === "string" ? tail.exec(v) : null));
  if (nonEmpty.length === k && parts.every((m) => m && m[1] === parts[0]![1])) {
    const nums = parts.map((m) => parseInt(m![2], 10));
    const d = k === 1 ? 1 : step(nums);
    const width = parts[0]![2].length;
    for (let p = 0; p < count; p++) {
      const n = Math.max(0, Math.round(nums[k - 1] + d * (p + 1)));
      out.push({ ...style(p), value: `${parts[0]![1]}${String(n).padStart(parts[0]![2].startsWith("0") ? width : 1, "0")}` });
    }
    return out;
  }

  for (let p = 0; p < count; p++) out.push({ ...style(p), value: values[p % k] });
  return out;
}
