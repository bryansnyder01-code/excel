// Dashboard page helpers: cell references, KPI values, tile layout and a default dashboard.
// Pure functions over the workbook; the components apply changes through useStore.mutate.

import { evaluateCached, type Computed } from "./formula";
import { columnLetter, formatValue, numberStyleAt } from "./format";
import type { DashboardSpec, DashboardTile, WorkbookSpec } from "./types";

export const GRID_COLS = 12;
export const MIN_W = 2;
export const MIN_H = 1;
export const MAX_H = 12;

const SIZES: Record<DashboardTile["kind"], { w: number; h: number }> = {
  kpi: { w: 3, h: 2 },
  chart: { w: 6, h: 5 },
  text: { w: 4, h: 2 },
};

export const defaultSize = (kind: DashboardTile["kind"]) => SIZES[kind];

/** A cell reference resolved against the workbook. `r` / `c` are grid coordinates (row 0 = header). */
export interface CellRef {
  sheet: number;
  r: number;
  c: number;
}

const REF = /^(?:(?:'((?:[^']|'')+)'|([^!'"]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$/;

/**
 * Parse "B4", "Summary!B4" or "'My sheet'!B4". A bare address resolves on `defaultSheet`.
 * Returns null for anything that isn't a single cell on an existing sheet.
 */
export function parseCellRef(wb: WorkbookSpec, ref: string | null | undefined, defaultSheet = 0): CellRef | null {
  if (!ref) return null;
  const m = REF.exec(ref.trim());
  if (!m) return null;
  const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]?.trim();
  const sheet = name === undefined ? defaultSheet : wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
  if (sheet < 0 || !wb.sheets[sheet]) return null;
  const c = [...m[3].toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const row = parseInt(m[4], 10);
  if (row < 1) return null;
  return { sheet, r: row - 1, c };
}

/** "Summary!B4", quoting sheet names that need it. */
export function formatCellRef(sheetName: string, r: number, c: number): string {
  const quoted = /^[A-Za-z_][\w.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
  return `${quoted}!${columnLetter(c)}${r + 1}`;
}

/** The value a cell shows: its literal, or its formula's result. */
export function computedAt(wb: WorkbookSpec, ref: CellRef): Computed | undefined {
  const sheet = wb.sheets[ref.sheet];
  if (!sheet) return undefined;
  if (ref.r === 0) return sheet.columns[ref.c]?.header;
  const raw = sheet.rows[ref.r - 1]?.cells[ref.c]?.value;
  if (typeof raw === "string" && raw.startsWith("=")) return evaluateCached(wb).get(`${ref.sheet}:${ref.r + 1}:${ref.c + 1}`);
  return raw ?? undefined;
}

export interface KpiValue {
  /** The value formatted the way the cell shows it, or a hint when the ref is broken. */
  text: string;
  number: number | null;
  error: boolean;
  /** Change against the comparison cell as a fraction (0.12 = +12%), when both are numbers. */
  change: number | null;
  /** Where the value lives, for jumping to it. */
  at: CellRef | null;
}

export function kpiValue(wb: WorkbookSpec, tile: DashboardTile): KpiValue {
  const at = parseCellRef(wb, tile.ref);
  if (!at) return { text: tile.ref ? "#REF!" : "No cell", number: null, error: true, change: null, at: null };
  const v = computedAt(wb, at);
  const shown = formatValue(v, numberStyleAt(wb.sheets[at.sheet], at.r, at.c));
  const number = typeof v === "number" ? v : null;
  let change: number | null = null;
  const cmpAt = parseCellRef(wb, tile.compare_ref, at.sheet);
  if (number !== null && cmpAt) {
    const base = computedAt(wb, cmpAt);
    if (typeof base === "number" && base !== 0) change = (number - base) / Math.abs(base);
  }
  return { text: shown.text || "–", number, error: shown.error, change, at };
}

/** Every chart in the workbook, for the chart tile picker. */
export function allCharts(wb: WorkbookSpec): { sheet: string; sheetIndex: number; index: number; title: string }[] {
  return wb.sheets.flatMap((s, si) => (s.charts ?? []).map((ch, i) => ({ sheet: s.name, sheetIndex: si, index: i, title: ch.title || `Chart ${i + 1}` })));
}

/** The chart a tile shows, or null when the sheet or chart is gone. */
export function tileChart(wb: WorkbookSpec, tile: DashboardTile) {
  const si = wb.sheets.findIndex((s) => s.name.toLowerCase() === (tile.sheet ?? "").toLowerCase());
  const chart = si >= 0 ? wb.sheets[si].charts?.[tile.chart ?? 0] : undefined;
  return chart ? { sheetIndex: si, chart } : null;
}

// ---------- Layout ----------

const overlaps = (a: DashboardTile, b: DashboardTile) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

/** Keep a tile inside the 12 columns and at a sensible size. */
export function clampTile<T extends DashboardTile>(t: T): T {
  const w = Math.max(MIN_W, Math.min(GRID_COLS, Math.round(t.w)));
  const h = Math.max(MIN_H, Math.min(MAX_H, Math.round(t.h)));
  return { ...t, w, h, x: Math.max(0, Math.min(GRID_COLS - w, Math.round(t.x))), y: Math.max(0, Math.round(t.y)) };
}

/**
 * Resolve overlaps after `pinnedId` moved or grew: the pinned tile stays put and any tile that
 * collides with it (or with a tile pushed before it) moves down just far enough. Gaps are kept.
 */
export function settle(tiles: DashboardTile[], pinnedId?: string): DashboardTile[] {
  const pinned = tiles.find((t) => t.id === pinnedId);
  const rest = tiles.filter((t) => t.id !== pinnedId).map(clampTile).sort((a, b) => a.y - b.y || a.x - b.x);
  const placed: DashboardTile[] = pinned ? [clampTile(pinned)] : [];
  for (const t of rest) {
    let y = t.y;
    for (;;) {
      const probe = { ...t, y };
      const hit = placed.find((p) => overlaps(p, probe));
      if (!hit) break;
      y = hit.y + hit.h;
    }
    placed.push({ ...t, y });
  }
  // Keep the original order so React keys and focus stay stable.
  const byId = new Map(placed.map((t) => [t.id, t]));
  return tiles.map((t) => byId.get(t.id)!);
}

/** Reading order: top to bottom, left to right. Used for tab order and the phone layout. */
export function readingOrder(tiles: DashboardTile[]): DashboardTile[] {
  return [...tiles].sort((a, b) => a.y - b.y || a.x - b.x);
}

/** The first free spot for a new tile of size w×h, scanning rows top to bottom. */
export function freeSpot(tiles: DashboardTile[], w: number, h: number): { x: number; y: number } {
  for (let y = 0; y < 200; y++)
    for (let x = 0; x + w <= GRID_COLS; x++) {
      const probe = { id: "", kind: "text", title: "", x, y, w, h } as DashboardTile;
      if (!tiles.some((t) => overlaps(t, probe))) return { x, y };
    }
  return { x: 0, y: tiles.reduce((m, t) => Math.max(m, t.y + t.h), 0) };
}

export function rowsUsed(tiles: DashboardTile[]): number {
  return tiles.reduce((m, t) => Math.max(m, t.y + t.h), 0);
}

// ---------- Default dashboard ----------

const labelOf = (wb: WorkbookSpec, si: number, rowIndex: number, c: number): string => {
  const sheet = wb.sheets[si];
  const header = sheet.columns[c]?.header?.trim() || columnLetter(c);
  // The row's first text cell is usually its label ("Total", "Net profit").
  const cells = sheet.rows[rowIndex].cells;
  const label = cells.slice(0, c).find((x) => typeof x.value === "string" && x.value && !x.value.startsWith("="))?.value as string | undefined;
  if (!label || /^(grand\s+)?totals?$/i.test(label.trim())) return `Total ${header}`.replace(/^Total total/i, "Total");
  return `${label.trim()} · ${header}`;
};

/**
 * A sensible dashboard from what the workbook already has, no AI involved:
 * one KPI per total row (its last numeric cell, plus the next-to-last numeric cell as the
 * comparison when the row has one), and every chart.
 */
export function suggestDashboard(wb: WorkbookSpec, makeId: () => string): DashboardSpec {
  const computed = evaluateCached(wb);
  const kpis: DashboardTile[] = [];
  wb.sheets.forEach((sheet, si) => {
    sheet.rows.forEach((row, ri) => {
      if (row.kind !== "total") return;
      const numericCols: number[] = [];
      row.cells.forEach((cell, c) => {
        const v = typeof cell.value === "string" && cell.value.startsWith("=") ? computed.get(`${si}:${ri + 2}:${c + 1}`) : cell.value;
        if (typeof v === "number") numericCols.push(c);
      });
      // A column headed "Total" is the row's grand total; otherwise its last number.
      const totalCol = numericCols.find((c) => /^\s*(grand\s+)?total\s*$/i.test(sheet.columns[c]?.header ?? ""));
      const last = totalCol ?? numericCols.at(-1);
      if (last === undefined) return;
      kpis.push({
        id: makeId(),
        kind: "kpi",
        title: labelOf(wb, si, ri, last),
        ref: formatCellRef(sheet.name, ri + 1, last),
        compare_ref: null,
        x: 0,
        y: 0,
        ...SIZES.kpi,
      });
    });
  });
  const tiles: DashboardTile[] = [];
  kpis.slice(0, 8).forEach((t, i) => tiles.push({ ...t, x: (i % 4) * 3, y: Math.floor(i / 4) * SIZES.kpi.h }));
  let y = rowsUsed(tiles);
  allCharts(wb).forEach((ch, i) => {
    tiles.push({
      id: makeId(),
      kind: "chart",
      title: ch.title,
      sheet: ch.sheet,
      chart: ch.index,
      x: (i % 2) * 6,
      y: y + Math.floor(i / 2) * SIZES.chart.h,
      ...SIZES.chart,
    });
  });
  y = rowsUsed(tiles);
  if (!tiles.length)
    tiles.push({
      id: makeId(),
      kind: "text",
      title: "Getting started",
      text: "Mark a row as a Total row or add a chart, then suggest again, or add tiles by hand.",
      x: 0,
      y,
      ...SIZES.text,
    });
  return { title: wb.dashboard?.title || `${wb.title} dashboard`, tiles };
}
