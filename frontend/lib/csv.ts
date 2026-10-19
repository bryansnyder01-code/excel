import { evaluateCached } from "./formula";
import type { WorkbookSpec } from "./types";

function escape(v: unknown): string {
  const s = v === null || v === undefined ? "" : String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

/** CSV of one sheet with formulas replaced by their current results. */
export function sheetToCsv(wb: WorkbookSpec, index: number): string {
  const sheet = wb.sheets[index];
  const computed = evaluateCached(wb);
  const lines = [sheet.columns.map((c) => escape(c.header)).join(",")];
  sheet.rows.forEach((row, ri) => {
    const cells = sheet.columns.map((_, ci) => {
      const v = row.cells[ci]?.value;
      if (typeof v === "string" && v.startsWith("=")) {
        const out = computed.get(`${index}:${ri + 2}:${ci + 1}`);
        return escape(out && typeof out === "object" ? out.error : out);
      }
      return escape(v);
    });
    lines.push(cells.join(","));
  });
  return "﻿" + lines.join("\r\n") + "\r\n";
}
