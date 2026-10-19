// Row layout for the virtualised grid: which rows are visible (filters hide some), how tall each
// is, and where it sits, so only the rows on screen need to be drawn. Pure, for scripts/check-grid.ts.

import { parseRange } from "./cellStyle";
import type { SheetSpec } from "./types";

export const ROW_H = 28;
const PX_PER_PT = 4 / 3;
const LINE_H = 17;

export interface RowLayout {
  /** Grid rows to draw, in order (row 0, the header, is always first). */
  visible: number[];
  /** Top of each entry of `visible`, in px from the top of row 0. */
  top: number[];
  height: number[];
  total: number;
  hidden: Set<number>;
}

/** Rows a sheet's filters hide (grid coordinates). `textAt` is what a cell displays. */
export function filteredRows(sheet: SheetSpec, textAt: (r: number, c: number) => string): Set<number> {
  const hidden = new Set<number>();
  const filters = (sheet.filters ?? []).filter((f) => f.values);
  if (!filters.length) return hidden;
  sheet.rows.forEach((row, i) => {
    if ((row.kind ?? "data") !== "data") return;
    const r = i + 1;
    if (filters.some((f) => !f.values.includes(textAt(r, f.column)))) hidden.add(r);
  });
  return hidden;
}

export function rowLayout(
  sheet: SheetSpec,
  nRows: number,
  hidden: Set<number>,
  /** Height for a row with wrapped text and no set height, or null. */
  autoHeight: (r: number) => number | null,
): RowLayout {
  const visible: number[] = [];
  const top: number[] = [];
  const height: number[] = [];
  let y = 0;
  for (let r = 0; r < nRows; r++) {
    if (hidden.has(r)) continue;
    const set = r === 0 ? sheet.header_height : sheet.rows[r - 1]?.height;
    const h = set ? Math.max(12, Math.round(set * PX_PER_PT)) : (autoHeight(r) ?? ROW_H);
    visible.push(r);
    top.push(y);
    height.push(h);
    y += h;
  }
  return { visible, top, height, total: y, hidden };
}

/** Lines needed to show `text` wrapped in a column `width` px wide, as a row height. */
export function wrappedHeight(text: string, width: number): number {
  const perLine = Math.max(4, Math.floor((width - 12) / 7));
  const lines = text.split("\n").reduce((n, part) => n + Math.max(1, Math.ceil(part.length / perLine)), 0);
  return Math.max(ROW_H, lines * LINE_H + 10);
}

/** Index into layout.visible of the first row whose bottom is below `y`. */
export function rowAt(layout: RowLayout, y: number): number {
  let lo = 0;
  let hi = layout.visible.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (layout.top[mid] + layout.height[mid] <= y) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

export interface MergeInfo {
  /** Anchor cells "r:c" (grid coordinates) and the area they cover. */
  anchors: Map<string, { r1: number; c1: number; r2: number; c2: number }>;
  /** Cells hidden under a merge, mapped to their anchor "r:c". */
  covered: Map<string, string>;
}

export function mergeInfo(sheet: SheetSpec): MergeInfo {
  const anchors = new Map<string, { r1: number; c1: number; r2: number; c2: number }>();
  const covered = new Map<string, string>();
  for (const m of sheet.merges ?? []) {
    const g = parseRange(m);
    if (!g || (g.r1 === g.r2 && g.c1 === g.c2)) continue;
    const box = { r1: g.r1 - 1, c1: g.c1 - 1, r2: g.r2 - 1, c2: g.c2 - 1 };
    const key = `${box.r1}:${box.c1}`;
    anchors.set(key, box);
    for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) if (r !== box.r1 || c !== box.c1) covered.set(`${r}:${c}`, key);
  }
  return { anchors, covered };
}
