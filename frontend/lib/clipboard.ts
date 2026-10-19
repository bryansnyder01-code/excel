"use client";

// Copy, cut and paste between the grid, Excel and Google Sheets.
// The system clipboard carries tab-separated text (what spreadsheets exchange); a copy made in
// this grid also remembers its cells, so pasting it back keeps formulas, styles and notes.

import { evaluateCached } from "./formula";
import { formatValue, numberStyleAt, parseInput, typedFormat } from "./format";
import { parseTsv, toTsv } from "./fill";
import { offsetFormula } from "./refs";
import { normalize, useStore } from "./store";
import type { CellSpec, Selection } from "./types";

interface Clip {
  tsv: string;
  cells: CellSpec[][];
  sheet: number;
  sel: Selection;
  cut: boolean;
}

let clip: Clip | null = null;

const escHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** The selection as clipboard text (what each cell shows) plus an HTML table for rich paste targets. */
export function copySelection(cut = false): { text: string; html: string } {
  const { workbook, active, selection } = useStore.getState();
  const sheet = workbook.sheets[active];
  const computed = evaluateCached(workbook);
  const s = normalize(selection);
  const lastRow = Math.min(s.r2, sheet.rows.length);
  const lastCol = Math.min(s.c2, Math.max(sheet.columns.length - 1, ...sheet.rows.map((r) => r.cells.length - 1)));
  const shown: string[][] = [];
  const cells: CellSpec[][] = [];
  for (let r = s.r1; r <= Math.max(s.r1, lastRow); r++) {
    const line: string[] = [];
    const raw: CellSpec[] = [];
    for (let c = s.c1; c <= Math.max(s.c1, lastCol); c++) {
      if (r === 0) {
        const h = sheet.columns[c]?.header ?? "";
        line.push(h);
        raw.push({ value: h });
        continue;
      }
      const cell = sheet.rows[r - 1]?.cells[c] ?? { value: null };
      const v = typeof cell.value === "string" && cell.value.startsWith("=") ? computed.get(`${active}:${r + 1}:${c + 1}`) : cell.value;
      line.push(formatValue(v, numberStyleAt(sheet, r, c)).text);
      raw.push(structuredClone(cell));
    }
    shown.push(line);
    cells.push(raw);
  }
  const text = toTsv(shown);
  clip = { tsv: text, cells, sheet: active, sel: { r1: s.r1, c1: s.c1, r2: s.r1 + cells.length - 1, c2: s.c1 + cells[0].length - 1 }, cut };
  const html = `<table>${shown.map((l) => `<tr>${l.map((v) => `<td>${escHtml(v)}</td>`).join("")}</tr>`).join("")}</table>`;
  return { text, html };
}

/** Paste clipboard text at the selection. Returns false when there was nothing to paste. */
export function pasteText(text: string): boolean {
  if (!text) return false;
  const st = useStore.getState();
  const s = normalize(st.selection);
  const internal = clip && clip.tsv === text.replace(/\r\n?/g, "\n").replace(/\n$/, "") ? clip : null;
  let block: CellSpec[][];
  if (internal) block = internal.cells;
  else {
    const rows = parseTsv(text);
    if (!rows.length) return false;
    block = rows.map((line, dr) =>
      line.map((v): CellSpec => {
        if (s.r1 + dr === 0) return { value: v };
        const value = parseInput(v);
        const implied = typeof value === "number" ? typedFormat(v) : null;
        return implied ? { value, format: implied.format, currency: implied.currency ?? null } : { value };
      }),
    );
  }
  const h = block.length;
  const w = Math.max(...block.map((l) => l.length));
  // A single cell, or a block that divides the selection evenly, is repeated across it, like Excel.
  const selH = s.r2 - s.r1 + 1;
  const selW = s.c2 - s.c1 + 1;
  const tile = (h === 1 && w === 1) || (selH % h === 0 && selW % w === 0);
  const outH = tile ? Math.max(h, selH) : h;
  const outW = tile ? Math.max(w, selW) : w;

  st.mutate((wb) => {
    const sheet = wb.sheets[st.active];
    if (internal?.cut) {
      const from = wb.sheets[internal.sheet];
      for (let r = Math.max(1, internal.sel.r1); r <= internal.sel.r2; r++)
        for (let c = internal.sel.c1; c <= internal.sel.c2; c++) {
          const row = from?.rows[r - 1];
          if (row?.cells[c]) row.cells[c] = { value: null };
        }
    }
    for (let dr = 0; dr < outH; dr++)
      for (let dc = 0; dc < outW; dc++) {
        const src = block[dr % h][dc % w] ?? { value: null };
        const r = s.r1 + dr;
        const c = s.c1 + dc;
        if (r === 0) {
          while (sheet.columns.length <= c) sheet.columns.push({ header: `Column ${sheet.columns.length + 1}`, format: "text" });
          sheet.columns[c].header = src.value === null || src.value === undefined ? "" : String(src.value);
          continue;
        }
        const cell: CellSpec = structuredClone(src);
        if (internal && typeof cell.value === "string" && cell.value.startsWith("=") && !internal.cut) {
          cell.value = offsetFormula(cell.value, r - (internal.sel.r1 + (dr % h)), c - (internal.sel.c1 + (dc % w)));
        }
        while (sheet.columns.length <= c) sheet.columns.push({ header: `Column ${sheet.columns.length + 1}`, format: "text" });
        while (sheet.rows.length < r) sheet.rows.push({ kind: "data", cells: [] });
        const row = sheet.rows[r - 1];
        while (row.cells.length <= c) row.cells.push({ value: null });
        row.cells[c] = cell;
      }
  });
  if (internal?.cut) clip = null;
  st.select({ r1: s.r1, c1: s.c1, r2: s.r1 + outH - 1, c2: s.c1 + outW - 1 });
  return true;
}
