// Live data connections: a CSV, Google Sheet or JSON feed pulled into a sheet of the workbook.

import { fetchData, type ConnectionKind, type FetchedTable } from "./aiApi";
import { INVALID_SHEET_CHARS, newId, useStore } from "./store";
import type { CellValue, ColumnSpec, DataConnection, SheetSpec, WorkbookSpec } from "./types";
import { useUi } from "./ui";

export const REFRESH_CHOICES: { minutes: number | null; label: string }[] = [
  { minutes: null, label: "Only when I click Refresh" },
  { minutes: 15, label: "Every 15 minutes" },
  { minutes: 60, label: "Every hour" },
  { minutes: 360, label: "Every 6 hours" },
  { minutes: 1440, label: "Every day" },
];

export const KIND_LABELS: Record<ConnectionKind, string> = { csv: "CSV file", gsheet: "Google Sheet", json: "JSON API" };

/** Best guess at what a URL serves. */
export function detectKind(url: string): ConnectionKind {
  const u = url.trim().toLowerCase();
  if (/docs\.google\.com\/spreadsheets\//.test(u)) return "gsheet";
  if (/\.json(\?|#|$)/.test(u) || /[?&](format|output|alt)=json\b/.test(u) || /\/api\/|^https?:\/\/api\./.test(u)) return "json";
  return "csv";
}

/** A valid sheet name from what was typed. */
export function cleanSheetName(name: string): string {
  return name.replace(new RegExp(INVALID_SHEET_CHARS.source, "g"), " ").replace(/\s+/g, " ").trim().slice(0, 31) || "Data";
}

const NUMBER = /^-?(0|[1-9]\d*)(\.\d+)?$/;

/** Numbers stay numbers, numeric-looking text becomes a number (codes like 00123 stay text). */
export function toCellValue(v: unknown): CellValue {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "boolean") return v;
  const s = typeof v === "string" ? v : JSON.stringify(v);
  const t = s.trim();
  return NUMBER.test(t) ? parseFloat(t) : s;
}

/**
 * Write a fetched table into the sheet named `sheetName` of `wb` (in place; call inside mutate).
 * Creates the sheet if needed, replaces its headers and rows, and keeps each column's format and
 * width when a column of the same header was there before. Returns the sheet's index.
 */
export function writeTable(wb: WorkbookSpec, sheetName: string, table: FetchedTable): { index: number; created: boolean } {
  const name = cleanSheetName(sheetName);
  let index = wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
  const created = index < 0;
  if (created) {
    wb.sheets.push({ name, columns: [], rows: [], freeze_header: true });
    index = wb.sheets.length - 1;
  }
  const sheet: SheetSpec = wb.sheets[index];
  const old = new Map<string, ColumnSpec>();
  for (const col of sheet.columns) if (!old.has(col.header.trim().toLowerCase())) old.set(col.header.trim().toLowerCase(), col);
  const headers = table.columns.length ? table.columns : (table.rows[0] ?? []).map((_, i) => `Column ${i + 1}`);
  sheet.columns = headers.map((h) => {
    const header = h === null || h === undefined ? "" : String(h);
    const prev = old.get(header.trim().toLowerCase());
    return prev ? { ...prev, header } : { header, format: "text" };
  });
  sheet.rows = table.rows.map((row) => ({
    kind: "data",
    cells: headers.map((_, c) => ({ value: toCellValue(row[c]) })),
  }));
  // Filters on columns that no longer exist would hide everything.
  if (sheet.filters?.length) sheet.filters = sheet.filters.filter((f) => f.column < headers.length);
  return { index, created };
}

/** last_refreshed in milliseconds, whether it was saved in seconds (server) or milliseconds (browser). */
export const refreshedAt = (c: DataConnection): number | null =>
  c.last_refreshed ? (c.last_refreshed < 1e12 ? c.last_refreshed * 1000 : c.last_refreshed) : null;

export function isDue(c: DataConnection, at = Date.now()): boolean {
  if (!c.refresh_minutes) return false;
  const last = refreshedAt(c);
  return last === null || at - last >= c.refresh_minutes * 60_000;
}

export function newConnection(fields: Omit<DataConnection, "id">): DataConnection {
  return { ...fields, id: newId(), sheet: cleanSheetName(fields.sheet) };
}

export function saveConnection(conn: DataConnection) {
  useStore.getState().mutate((wb) => {
    const list = wb.connections ?? [];
    const i = list.findIndex((c) => c.id === conn.id);
    wb.connections = i < 0 ? [...list, conn] : list.map((c) => (c.id === conn.id ? conn : c));
  });
}

export function deleteConnection(id: string) {
  useStore.getState().mutate((wb) => {
    wb.connections = (wb.connections ?? []).filter((c) => c.id !== id);
  });
}

const refreshing = new Set<string>();
export const isRefreshing = (id: string) => refreshing.has(id);

/**
 * Fetch a connection's data and write it into its sheet, in one undo step.
 * Returns the table, or throws with the server's message.
 */
export async function refreshConnection(id: string): Promise<FetchedTable | null> {
  const start = useStore.getState();
  const conn = start.workbook.connections?.find((c) => c.id === id);
  if (!conn || refreshing.has(id)) return null;
  refreshing.add(id);
  try {
    const table = await fetchData(conn.kind, conn.url, conn.path);
    const st = useStore.getState();
    // The workbook was switched or the connection removed while fetching.
    if (st.docId !== start.docId || !st.workbook.connections?.some((c) => c.id === id)) return null;
    st.mutate((wb) => {
      const c = wb.connections!.find((x) => x.id === id)!;
      writeTable(wb, c.sheet, table);
      c.last_refreshed = Date.now();
    });
    return table;
  } finally {
    refreshing.delete(id);
  }
}

/**
 * While the app is open, refresh the open workbook's connections that are due, checking every minute.
 * Returns a function that stops the timer.
 */
export function startConnectionTimer(): () => void {
  const failedAt = new Map<string, number>();
  const tick = async () => {
    if (typeof document !== "undefined" && document.hidden) return;
    const st = useStore.getState();
    // Don't fold an AI change that's waiting for review into a refresh.
    if (st.review || st.busy) return;
    for (const c of st.workbook.connections ?? []) {
      if (!isDue(c) || refreshing.has(c.id)) continue;
      const failed = failedAt.get(c.id);
      if (failed && Date.now() - failed < (c.refresh_minutes ?? 15) * 60_000) continue;
      try {
        await refreshConnection(c.id);
        failedAt.delete(c.id);
      } catch (e) {
        failedAt.set(c.id, Date.now());
        useUi.getState().notify(`Couldn't refresh "${c.name}": ${e instanceof Error ? e.message : "the data couldn't be fetched."}`, true);
      }
    }
  };
  const first = setTimeout(tick, 5_000);
  const every = setInterval(tick, 60_000);
  return () => {
    clearTimeout(first);
    clearInterval(every);
  };
}
