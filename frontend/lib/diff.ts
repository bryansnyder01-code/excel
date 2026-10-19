// What an AI change did to the workbook, cell by cell, for the review banner and grid highlights.

import type { SheetSpec, WorkbookSpec } from "./types";

export interface SheetDiff {
  name: string;
  status: "added" | "changed" | "removed";
  /** Changed cells as "r:c" in grid coordinates (row 0 is the header row). */
  cells: Set<string>;
}

const empty = (v: unknown) => v === null || v === undefined || v === "";

function cellKey(sheet: SheetSpec, r: number, c: number): string {
  if (r === 0) return JSON.stringify(sheet.columns[c]?.header ?? null);
  const cell = sheet.rows[r - 1]?.cells[c];
  if (!cell) return "null";
  const { value, ...style } = cell;
  const clean = Object.fromEntries(Object.entries(style).filter(([, v]) => v !== null && v !== undefined && v !== false));
  return JSON.stringify([empty(value) ? null : value, clean, sheet.rows[r - 1]?.kind ?? "data"]);
}

function diffSheet(before: SheetSpec | undefined, after: SheetSpec): Set<string> {
  const out = new Set<string>();
  const rows = Math.max(after.rows.length, before?.rows.length ?? 0) + 1;
  const cols = Math.max(after.columns.length, before?.columns.length ?? 0, ...after.rows.map((r) => r.cells.length));
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const a = cellKey(after, r, c);
      if (!before) {
        if (a !== "null" && a !== JSON.stringify(null)) out.add(`${r}:${c}`);
      } else if (a !== cellKey(before, r, c)) out.add(`${r}:${c}`);
    }
  return out;
}

export function diffWorkbooks(before: WorkbookSpec, after: WorkbookSpec): SheetDiff[] {
  const out: SheetDiff[] = [];
  const old = new Map(before.sheets.map((s) => [s.name.toLowerCase(), s]));
  for (const sheet of after.sheets) {
    const prev = old.get(sheet.name.toLowerCase());
    const cells = diffSheet(prev, sheet);
    const extras =
      JSON.stringify([sheet.charts ?? [], sheet.validations ?? [], sheet.conditional_formats ?? []]) !==
      JSON.stringify([prev?.charts ?? [], prev?.validations ?? [], prev?.conditional_formats ?? []]);
    if (!prev) out.push({ name: sheet.name, status: "added", cells });
    else if (cells.size || extras) out.push({ name: sheet.name, status: "changed", cells });
    old.delete(sheet.name.toLowerCase());
  }
  for (const s of old.values()) out.push({ name: s.name, status: "removed", cells: new Set() });
  return out;
}
