// Who may change what: viewers change nothing; on a protected sheet, or inside a locked range,
// only the owner may change cells. The server enforces the same rules; this check just stops
// a refused edit before it's made.

import { parseRange } from "./cellStyle";
import type { SheetSpec, WorkbookSpec } from "./types";

export type Role = "owner" | "editor" | "viewer" | null;

/** Is grid cell (r, c) of `sheet` protected from non-owners? */
export function isLocked(sheet: SheetSpec, r: number, c: number): boolean {
  if (sheet.locked) return true;
  return (sheet.locked_ranges ?? []).some((range) => {
    const g = parseRange(range);
    return !!g && r + 1 >= g.r1 && r + 1 <= g.r2 && c + 1 >= g.c1 && c + 1 <= g.c2;
  });
}

const same = (a: unknown, b: unknown) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

/** Why `after` isn't an allowed change of `before` for this role, or null when it's fine. */
export function editRefusal(role: Role, before: WorkbookSpec, after: WorkbookSpec): string | null {
  if (!role || role === "owner") return null;
  if (role === "viewer") return "You can view this workbook but not change it.";
  for (const sheet of before.sheets) {
    if (!sheet.locked && !sheet.locked_ranges?.length) continue;
    const next = after.sheets.find((s) => s.name === sheet.name);
    if (!next) return `"${sheet.name}" is protected, so only the owner can delete it.`;
    if (!same(sheet.locked, next.locked) || !same(sheet.locked_ranges, next.locked_ranges))
      return "Only the owner can change which cells are locked.";
    if (sheet.locked && !same(sheet, next)) return `"${sheet.name}" is locked by its owner.`;
    const rows = Math.max(sheet.rows.length, next.rows.length);
    for (let i = 0; i < rows; i++) {
      const a = sheet.rows[i];
      const b = next.rows[i];
      if (a === b) continue;
      const cols = Math.max(a?.cells.length ?? 0, b?.cells.length ?? 0, sheet.columns.length);
      for (let c = 0; c < cols; c++)
        if (isLocked(sheet, i + 1, c) && !same(a?.cells[c]?.value, b?.cells[c]?.value)) return "That cell is locked by the workbook's owner.";
    }
    for (let c = 0; c < sheet.columns.length; c++)
      if (isLocked(sheet, 0, c) && sheet.columns[c].header !== next.columns[c]?.header) return "That header is locked by the workbook's owner.";
  }
  return null;
}
