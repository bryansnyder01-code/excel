// Grid styling that mirrors backend/app/services/workbook_builder.py,
// so the preview looks like the downloaded file.

import type { CSSProperties } from "react";
import type { Computed } from "./formula";
import type { CellSpec, CondStyle, ConditionalFormatSpec, RowKind, SheetSpec } from "./types";

// Fills are CSS variables (see globals.css) so dark mode can soften them.
export const THEME = {
  headerFill: "#1F4E78",
  subtotalFill: "var(--sc-subtotal)",
  totalFill: "var(--sc-total)",
  bandFill: "var(--sc-band)",
};

const COND: Record<Exclude<CondStyle, "custom">, { bg: string; fg: string }> = {
  good: { bg: "#C6EFCE", fg: "#006100" },
  bad: { bg: "#FFC7CE", fg: "#9C0006" },
  warning: { bg: "#FFEB9C", fg: "#9C5700" },
};

export interface CondPaint {
  bg?: string;
  fg?: string;
  /** Data bar: share of the cell to fill, 0 to 1, and its colour. */
  bar?: { pct: number; color: string };
}

export function parseRange(range: string) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(range.trim().toUpperCase());
  if (!m) return null;
  const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const c1 = col(m[1]);
  const r1 = +m[2];
  const c2 = m[3] ? col(m[3]) : c1;
  const r2 = m[4] ? +m[4] : r1;
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

function matches(cf: ConditionalFormatSpec, v: Computed | undefined) {
  if (cf.operator === "contains") {
    const needle = (cf.text ?? "").toLowerCase();
    return !!needle && v !== null && v !== undefined && typeof v !== "object" && String(v).toLowerCase().includes(needle);
  }
  if (typeof v !== "number") return false;
  const a = cf.value1 ?? 0;
  switch (cf.operator) {
    case "negative":
      return v < 0;
    case "greater_than":
      return v > a;
    case "less_than":
      return v < a;
    case "equal":
      return v === a;
    case "between":
      return v >= Math.min(a, cf.value2 ?? a) && v <= Math.max(a, cf.value2 ?? a);
    default:
      return false;
  }
}

const hex = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
function mix(a: string, b: string, t: number) {
  const [x, y] = [hex(a), hex(b)];
  return `rgb(${x.map((v, i) => Math.round(v + (y[i] - v) * t)).join(",")})`;
}
const SCALE = ["#F8696B", "#FFEB84", "#63BE7B"];

/**
 * Conditional formatting for one sheet, like Excel's: highlight rules, text contains,
 * 3-colour scales and data bars. `valueAt(row, col)` takes Excel 1-based coordinates.
 */
export function conditionalPainter(sheet: SheetSpec, valueAt: (row: number, col: number) => Computed | undefined) {
  const rules = (sheet.conditional_formats ?? []).flatMap((cf) => {
    const range = parseRange(cf.range);
    if (!range) return [];
    let min = Infinity;
    let max = -Infinity;
    if (cf.operator === "color_scale" || cf.operator === "data_bar") {
      for (let r = range.r1; r <= Math.min(range.r2, sheet.rows.length + 1); r++)
        for (let c = range.c1; c <= range.c2; c++) {
          const v = valueAt(r, c);
          if (typeof v === "number") (min = Math.min(min, v)), (max = Math.max(max, v));
        }
    }
    return [{ cf, range, min, max }];
  });
  return (row: number, col: number, value: Computed | undefined): CondPaint | null => {
    let out: CondPaint | null = null;
    for (const { cf, range, min, max } of rules) {
      if (row < range.r1 || row > range.r2 || col < range.c1 || col > range.c2) continue;
      if (cf.operator === "color_scale") {
        if (typeof value !== "number" || !isFinite(min)) continue;
        let t = max === min ? 0.5 : (value - min) / (max - min);
        if (cf.reverse) t = 1 - t;
        const bg = t < 0.5 ? mix(SCALE[0], SCALE[1], t * 2) : mix(SCALE[1], SCALE[2], (t - 0.5) * 2);
        out = { ...(out ?? {}), bg, fg: "#1f2937" };
      } else if (cf.operator === "data_bar") {
        if (typeof value !== "number" || !isFinite(max)) continue;
        const lo = Math.min(0, min);
        const pct = max === lo ? 1 : (value - lo) / (max - lo);
        out = { ...(out ?? {}), bar: { pct: Math.max(0, Math.min(1, pct)), color: cf.fill_color || "#638EC6" } };
      } else if (!out?.bg && matches(cf, value)) {
        const style = cf.style ?? "good";
        out = { ...(out ?? {}), ...(style === "custom" ? { bg: cf.fill_color ?? undefined, fg: cf.font_color ?? undefined } : COND[style]) };
      }
    }
    return out;
  };
}

export function dataCellStyle(
  cell: CellSpec | undefined,
  kind: RowKind,
  banded: boolean,
  numeric: boolean,
  cond: CondPaint | null,
): CSSProperties {
  const bold = cell?.bold || kind === "total" || kind === "subtotal" || kind === "section";
  const style: CSSProperties = {
    fontWeight: bold ? 600 : 400,
    fontStyle: cell?.italic ? "italic" : undefined,
    textDecoration: cell?.underline ? "underline" : undefined,
    textAlign: cell?.align ?? (numeric ? "right" : "left"),
  };
  if (kind === "total") {
    style.background = THEME.totalFill;
    style.boxShadow = "inset 0 2px 0 var(--sc-total-line)";
  } else if (kind === "subtotal" || kind === "section") style.background = THEME.subtotalFill;
  else if (banded) style.background = THEME.bandFill;

  if (cond?.bg) style.background = cond.bg;
  if (cond?.fg) style.color = cond.fg;
  if (cond?.bar) {
    const base = (style.background as string | undefined) ?? "var(--sc-cell)";
    style.background = `linear-gradient(90deg, ${cond.bar.color}AA ${cond.bar.pct * 100}%, transparent ${cond.bar.pct * 100}%), ${base}`;
  }
  if (cell?.fill_color) {
    style.background = cell.fill_color;
    // Readable on a custom fill in dark mode too.
    if (!cond?.fg) style.color = "#111827";
  }
  if (cell?.font_color) style.color = cell.font_color;
  if (cell?.border === "thick") {
    style.outline = "2px solid var(--sc-strong)";
    style.outlineOffset = "-2px";
  }
  if (cell?.border === "none") style.borderColor = "transparent";
  return style;
}
