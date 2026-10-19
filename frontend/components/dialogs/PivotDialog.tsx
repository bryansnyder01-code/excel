"use client";

import { useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import { Columns3, Hash, Rows3, Sigma, Type, X } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { runPivot } from "@/lib/actions";
import { evaluateCached } from "@/lib/formula";
import { formatValue, type NumberStyle } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { Aggregation, NumberFormat, SheetSpec, WorkbookSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";

const AGGS: { value: Aggregation; label: string }[] = [
  { value: "sum", label: "Sum" },
  { value: "average", label: "Average" },
  { value: "count", label: "Count of rows" },
  { value: "min", label: "Minimum" },
  { value: "max", label: "Maximum" },
];

type Zone = "rows" | "columns" | "values";
const ZONES: { id: Zone; label: string; key: string; hint: string; icon: typeof Rows3 }[] = [
  { id: "rows", label: "Rows", key: "R", hint: "One row per value, e.g. Category", icon: Rows3 },
  { id: "columns", label: "Columns", key: "C", hint: "Optional, e.g. Quarter", icon: Columns3 },
  { id: "values", label: "Values", key: "V", hint: "The numbers to add up", icon: Sigma },
];

// Same limits as the server (backend/app/services/pivot.py).
const MAX_GROUPS = 200;
const MAX_COLUMN_GROUPS = 24;
const PREVIEW_ROWS = 20;
const NUMERIC: NumberFormat[] = ["currency", "decimal", "integer", "percent"];

interface FieldInfo {
  name: string;
  index: number;
  numeric: boolean;
}

/** The sheet's columns, marking those that hold numbers (by format, or by most of their values). */
function fieldsOf(wb: WorkbookSpec, si: number): FieldInfo[] {
  const sheet = wb.sheets[si];
  const computed = evaluateCached(wb);
  const seen = new Set<string>();
  return sheet.columns.flatMap((col, c) => {
    const name = col.header.trim();
    if (!name || seen.has(name.toLowerCase())) return [];
    seen.add(name.toLowerCase());
    let nums = 0;
    let filled = 0;
    sheet.rows.forEach((row, ri) => {
      if ((row.kind ?? "data") !== "data") return;
      const v = valueOf(row.cells[c]?.value, computed, si, ri, c);
      if (v === null || v === undefined || v === "") return;
      filled++;
      if (typeof v === "number") nums++;
    });
    const numeric = NUMERIC.includes(col.format ?? "text") || (filled > 0 && nums / filled >= 0.8);
    return [{ name, index: c, numeric }];
  });
}

function valueOf(raw: unknown, computed: Map<string, unknown>, si: number, ri: number, c: number): unknown {
  return typeof raw === "string" && raw.startsWith("=") ? computed.get(`${si}:${ri + 2}:${c + 1}`) : raw;
}

interface Acc {
  sum: number;
  n: number;
  rows: number;
  min: number;
  max: number;
}
const newAcc = (): Acc => ({ sum: 0, n: 0, rows: 0, min: Infinity, max: -Infinity });
function add(acc: Acc, v: unknown) {
  acc.rows++;
  if (typeof v === "number" && Number.isFinite(v)) {
    acc.sum += v;
    acc.n++;
    acc.min = Math.min(acc.min, v);
    acc.max = Math.max(acc.max, v);
  }
}
function result(acc: Acc | undefined, agg: Aggregation): number | null {
  if (!acc) return agg === "count" || agg === "sum" ? 0 : null;
  switch (agg) {
    case "count":
      return acc.rows;
    case "sum":
      return acc.sum;
    case "average":
      return acc.n ? acc.sum / acc.n : null;
    case "min":
      return acc.n ? acc.min : null;
    case "max":
      return acc.n ? acc.max : null;
  }
}

interface Group {
  key: string;
  label: string;
  sort: [number, number | string];
}

interface Preview {
  rows: Group[];
  cols: Group[];
  cells: Map<string, Acc>;
  rowTotals: Map<string, Acc>;
  colTotals: Map<string, Acc>;
  grand: Acc;
  totalRows: number;
  error: string | null;
}

/**
 * Group-by over the data rows, mirroring the server: groups come from typed-in values (not formulas,
 * not blanks), text compares case-insensitively, numbers sort before text.
 */
function computePreview(wb: WorkbookSpec, si: number, row: FieldInfo, col: FieldInfo | null, val: FieldInfo | null): Preview {
  const sheet = wb.sheets[si];
  const computed = evaluateCached(wb) as Map<string, unknown>;
  const rows = new Map<string, Group>();
  const cols = new Map<string, Group>();
  const cells = new Map<string, Acc>();
  const rowTotals = new Map<string, Acc>();
  const colTotals = new Map<string, Acc>();
  const grand = newAcc();
  const groupOf = (v: unknown, c: number): Group | null => {
    if (v === null || v === undefined || v === "" || (typeof v === "string" && v.startsWith("="))) return null;
    const key = typeof v === "string" ? `s:${v.trim().toLowerCase()}` : `v:${String(v)}`;
    const label = formatValue(v as string | number | boolean, styleOfColumn(sheet, c)).text;
    return { key, label, sort: typeof v === "number" ? [0, v] : [1, String(v).trim().toLowerCase()] };
  };
  sheet.rows.forEach((r, ri) => {
    if ((r.kind ?? "data") !== "data") return;
    const g = groupOf(r.cells[row.index]?.value, row.index);
    if (!g) return;
    let cg: Group | null = null;
    if (col) {
      cg = groupOf(r.cells[col.index]?.value, col.index);
      if (!cg) return;
      if (!cols.has(cg.key)) cols.set(cg.key, cg);
    }
    if (!rows.has(g.key)) rows.set(g.key, g);
    const v = val ? valueOf(r.cells[val.index]?.value, computed, si, ri, val.index) : null;
    const bump = (m: Map<string, Acc>, k: string) => {
      let acc = m.get(k);
      if (!acc) m.set(k, (acc = newAcc()));
      add(acc, v);
    };
    bump(cells, `${g.key}\u0000${cg?.key ?? ""}`);
    bump(rowTotals, g.key);
    if (cg) bump(colTotals, cg.key);
    add(grand, v);
  });
  const order = (a: Group, b: Group) =>
    a.sort[0] - b.sort[0] || (typeof a.sort[1] === "number" && typeof b.sort[1] === "number" ? a.sort[1] - b.sort[1] : String(a.sort[1]).localeCompare(String(b.sort[1]), undefined, { numeric: true }));
  const rowList = [...rows.values()].sort(order);
  const colList = [...cols.values()].sort(order);
  let error: string | null = null;
  if (!rowList.length) error = `The "${row.name}" column has no typed-in values to group by.`;
  else if (rowList.length > MAX_GROUPS) error = `"${row.name}" has ${rowList.length} different values; summaries support up to ${MAX_GROUPS}.`;
  else if (col && colList.length > MAX_COLUMN_GROUPS) error = `"${col.name}" has ${colList.length} different values; columns support up to ${MAX_COLUMN_GROUPS}.`;
  return { rows: rowList, cols: colList, cells, rowTotals, colTotals, grand, totalRows: rowList.length, error };
}

function styleOfColumn(sheet: SheetSpec, c: number): NumberStyle {
  const col = sheet.columns[c];
  return { fmt: col?.format ?? "text", decimals: col?.decimals ?? null, currency: col?.currency ?? null };
}

/** The number format the summary uses, as the server picks it. */
function resultStyle(sheet: SheetSpec, val: FieldInfo | null, agg: Aggregation): NumberStyle {
  if (agg === "count" || !val) return { fmt: "integer" };
  const s = styleOfColumn(sheet, val.index);
  let fmt = s.fmt;
  if (fmt === "text" || fmt === "date") fmt = agg === "average" ? "decimal" : "integer";
  if (agg === "average" && fmt === "integer") fmt = "decimal";
  return { ...s, fmt };
}

/** "Totals by category": drag fields into Rows, Columns and Values, preview live, then build a summary tab. */
export default function PivotDialog() {
  const close = useUi((s) => s.close);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const [sheetName, setSheetName] = useState(workbook.sheets[active].name);
  const si = Math.max(0, workbook.sheets.findIndex((s) => s.name === sheetName));
  const sheet = workbook.sheets[si];
  const fields = useMemo(() => fieldsOf(workbook, si), [workbook, si]);
  const byName = (n: string) => fields.find((f) => f.name === n) ?? null;

  const defaults = (fs: FieldInfo[]) => ({
    rows: fs.find((f) => !f.numeric)?.name ?? fs[0]?.name ?? "",
    columns: "",
    values: fs.find((f) => f.numeric)?.name ?? "",
  });
  const [zones, setZones] = useState<Record<Zone, string>>(() => defaults(fields));
  const [agg, setAgg] = useState<Aggregation>(() => (defaults(fields).values ? "sum" : "count"));
  const [message, setMessage] = useState("");
  const [drag, setDrag] = useState<null | { field: string; x: number; y: number; sx: number; sy: number; active: boolean; over: Zone | null }>(null);
  const zoneRefs = useRef<Record<Zone, HTMLDivElement | null>>({ rows: null, columns: null, values: null });

  const changeSheet = (name: string) => {
    setSheetName(name);
    const i = workbook.sheets.findIndex((s) => s.name === name);
    const fs = fieldsOf(workbook, i);
    const d = defaults(fs);
    setZones(d);
    setAgg(d.values ? "sum" : "count");
    setMessage("");
  };

  /** Put `field` in `zone`. Each zone holds one field; a field can't be both the rows and the columns. */
  const assign = (zone: Zone, field: string) => {
    const f = byName(field);
    if (!f) return;
    const next = { ...zones };
    const notes: string[] = [];
    if (next[zone] && next[zone] !== field) notes.push(`${ZONES.find((z) => z.id === zone)!.label} holds one field, so ${field} replaced ${next[zone]}.`);
    if (zone === "rows" && next.columns === field) next.columns = "";
    if (zone === "columns" && next.rows === field) next.rows = "";
    next[zone] = field;
    setZones(next);
    if (zone === "values" && !f.numeric && agg !== "count" && agg !== "min" && agg !== "max") {
      setAgg("count");
      notes.push(`${field} isn't numbers, so the summary counts rows.`);
    } else if (zone === "values" && f.numeric && agg === "count") {
      setAgg("sum");
    }
    setMessage(notes.join(" ") || `${field} is now in ${ZONES.find((z) => z.id === zone)!.label}.`);
  };
  const clearZone = (zone: Zone) => {
    setZones({ ...zones, [zone]: "" });
    if (zone === "values") setAgg("count");
    setMessage(`${ZONES.find((z) => z.id === zone)!.label} cleared.`);
  };

  // ---- Pointer drag of a chip onto a zone ----
  const zoneAt = (x: number, y: number): Zone | null => {
    for (const z of ZONES) {
      const r = zoneRefs.current[z.id]?.getBoundingClientRect();
      if (r && x >= r.left && x <= r.right && y >= r.top && y <= r.bottom) return z.id;
    }
    return null;
  };
  const onChipDown = (e: ReactPointerEvent<HTMLButtonElement>, field: string) => {
    if (e.button !== 0) return;
    e.currentTarget.setPointerCapture(e.pointerId);
    setDrag({ field, x: e.clientX, y: e.clientY, sx: e.clientX, sy: e.clientY, active: false, over: null });
  };
  const onChipMove = (e: ReactPointerEvent<HTMLButtonElement>) => {
    if (!drag) return;
    const moved = drag.active || Math.hypot(e.clientX - drag.sx, e.clientY - drag.sy) > 4;
    setDrag({ ...drag, x: e.clientX, y: e.clientY, active: moved, over: moved ? zoneAt(e.clientX, e.clientY) : null });
  };
  const onChipUp = (e: ReactPointerEvent<HTMLButtonElement>) => {
    if (!drag) return;
    const zone = drag.active ? zoneAt(e.clientX, e.clientY) : null;
    if (zone) assign(zone, drag.field);
    setDrag(null);
  };
  const onChipKey = (e: ReactKeyboardEvent<HTMLButtonElement>, field: string) => {
    const z = ZONES.find((x) => x.key === e.key.toUpperCase());
    if (z && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      assign(z.id, field);
    }
  };

  const rowF = byName(zones.rows);
  const colF = zones.columns ? byName(zones.columns) : null;
  const valF = zones.values ? byName(zones.values) : null;
  const preview = useMemo(() => (rowF ? computePreview(workbook, si, rowF, colF, agg === "count" ? null : valF) : null), [workbook, si, rowF, colF, valF, agg]);
  const style = resultStyle(sheet, valF, agg);
  const fmt = (acc: Acc | undefined) => {
    const n = result(acc, agg);
    return n === null ? "" : formatValue(n, style).text;
  };
  const shownCols = preview?.cols.slice(0, 8) ?? [];

  const valid = !!rowF && (agg === "count" || !!valF) && zones.rows !== zones.columns && !preview?.error;

  const submit = () => {
    if (!valid) return;
    close();
    runPivot({
      source_sheet: sheet.name,
      row_field: zones.rows,
      column_field: zones.columns || null,
      value_field: agg === "count" ? null : zones.values,
      aggregation: agg,
    });
  };

  const valueTitle = agg === "count" ? "Count" : `${AGGS.find((a) => a.value === agg)!.label} of ${zones.values || "…"}`;

  return (
    <Modal
      title="Summarise by category"
      onClose={close}
      width="max-w-3xl"
      footer={
        <>
          <Button onClick={close}>Cancel</Button>
          <Button primary disabled={!valid} onClick={submit}>
            Create summary tab
          </Button>
        </>
      }
    >
      <p className="mb-4 text-sm text-slate-500">
        Drag fields into Rows, Columns and Values, or focus a field and press <kbd className="rounded border border-slate-300 px-1 font-mono text-[11px]">R</kbd>,{" "}
        <kbd className="rounded border border-slate-300 px-1 font-mono text-[11px]">C</kbd> or <kbd className="rounded border border-slate-300 px-1 font-mono text-[11px]">V</kbd>. The new
        tab uses live SUMIFS-style formulas and a chart, so it updates whenever the data changes.
      </p>
      <div className="grid gap-x-4 sm:grid-cols-2">
        <Field label="Data sheet">
          <select className={inputClass} value={sheet.name} onChange={(e) => changeSheet(e.target.value)}>
            {workbook.sheets.map((s) => (
              <option key={s.name}>{s.name}</option>
            ))}
          </select>
        </Field>
        <Field label="Calculate">
          <select className={inputClass} value={agg} onChange={(e) => setAgg(e.target.value as Aggregation)}>
            {AGGS.map((a) => (
              <option key={a.value} value={a.value}>
                {a.label}
              </option>
            ))}
          </select>
        </Field>
      </div>

      <div className="mb-1 text-sm font-medium text-slate-700" id="pivot-fields-label">
        Fields
      </div>
      <div role="group" aria-labelledby="pivot-fields-label" className="mb-4 flex flex-wrap gap-1.5">
        {fields.length === 0 && <span className="text-sm text-slate-500">This sheet has no column headers.</span>}
        {fields.map((f) => {
          const placed = (Object.keys(zones) as Zone[]).filter((z) => zones[z] === f.name);
          return (
            <button
              key={f.name}
              type="button"
              className={`flex touch-none items-center gap-1 rounded-full border px-2.5 py-1 text-xs select-none focus:outline-none focus-visible:ring-2 focus-visible:ring-brand ${
                placed.length ? "border-brand/50 bg-brand/5 text-brand" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
              } ${drag?.field === f.name && drag.active ? "opacity-40" : ""} cursor-grab`}
              aria-label={`${f.name}${f.numeric ? ", numbers" : ""}${placed.length ? `, in ${placed.join(" and ")}` : ""}. Press R for rows, C for columns, V for values.`}
              title="Drag into Rows, Columns or Values"
              onPointerDown={(e) => onChipDown(e, f.name)}
              onPointerMove={onChipMove}
              onPointerUp={onChipUp}
              onPointerCancel={() => setDrag(null)}
              onKeyDown={(e) => onChipKey(e, f.name)}
            >
              {f.numeric ? <Hash className="h-3 w-3 text-slate-400" aria-hidden="true" /> : <Type className="h-3 w-3 text-slate-400" aria-hidden="true" />}
              {f.name}
            </button>
          );
        })}
      </div>

      <div className="mb-2 grid gap-3 sm:grid-cols-3">
        {ZONES.map((z) => {
          const Icon = z.icon;
          const value = zones[z.id];
          const over = drag?.active && drag.over === z.id;
          return (
            <div
              key={z.id}
              ref={(el) => {
                zoneRefs.current[z.id] = el;
              }}
              data-zone={z.id}
              className={`rounded-lg border-2 border-dashed p-2.5 transition-colors ${over ? "border-brand bg-brand/5" : "border-slate-200 bg-slate-50"}`}
            >
              <div className="mb-1.5 flex items-center gap-1.5 text-xs font-semibold text-slate-600">
                <Icon className="h-3.5 w-3.5" aria-hidden="true" />
                {z.label}
                <kbd className="ml-auto rounded border border-slate-300 bg-white px-1 font-mono text-[10px] font-normal text-slate-500" aria-hidden="true">
                  {z.key}
                </kbd>
              </div>
              {value ? (
                <span className="mb-1.5 flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2 py-1 text-xs text-slate-800">
                  <span className="min-w-0 flex-1 truncate">{z.id === "values" ? valueTitle : value}</span>
                  <button type="button" aria-label={`Remove ${value} from ${z.label}`} className="rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-red-600" onClick={() => clearZone(z.id)}>
                    <X className="h-3 w-3" />
                  </button>
                </span>
              ) : (
                <span className="mb-1.5 block rounded-md px-2 py-1 text-xs text-slate-400">{z.hint}</span>
              )}
              <select
                aria-label={`${z.label} field`}
                className="w-full rounded-md border border-slate-300 bg-white px-2 py-1 text-xs text-slate-700"
                value={value}
                onChange={(e) => (e.target.value ? assign(z.id, e.target.value) : clearZone(z.id))}
              >
                <option value="">{z.id === "rows" ? "Choose a field" : "None"}</option>
                {fields.map((f) => (
                  <option key={f.name} value={f.name}>
                    {f.name}
                  </option>
                ))}
              </select>
            </div>
          );
        })}
      </div>
      <p className="mb-4 min-h-[1.25rem] text-xs text-slate-500" aria-live="polite">
        {message || "Each area holds one field: the summary tab has one row field, an optional column field and one value."}
      </p>

      <section aria-label="Preview">
        <h3 className="mb-1.5 text-sm font-medium text-slate-700">Preview</h3>
        {!rowF ? (
          <p className="rounded-lg border border-slate-200 p-4 text-sm text-slate-500">Put a field in Rows to see the summary.</p>
        ) : agg !== "count" && !valF ? (
          <p className="rounded-lg border border-slate-200 p-4 text-sm text-slate-500">Put a field of numbers in Values, or choose Count of rows.</p>
        ) : preview?.error ? (
          <p className="rounded-lg border border-red-200 bg-red-50 p-4 text-sm text-red-700">{preview.error}</p>
        ) : (
          preview && (
            <>
              <div className="max-h-72 overflow-auto rounded-lg border border-slate-200">
                <table className="w-full text-sm">
                  <caption className="sr-only">
                    {valueTitle} by {zones.rows}
                    {colF ? ` and ${zones.columns}` : ""}
                  </caption>
                  <thead className="sticky top-0 bg-slate-50 text-xs text-slate-500">
                    <tr>
                      <th scope="col" className="px-3 py-2 text-left font-medium">
                        {zones.rows}
                      </th>
                      {colF ? (
                        <>
                          {shownCols.map((c) => (
                            <th key={c.key} scope="col" className="px-3 py-2 text-right font-medium whitespace-nowrap">
                              {c.label}
                            </th>
                          ))}
                          {preview.cols.length > shownCols.length && <th className="px-3 py-2 text-right font-medium">…</th>}
                          <th scope="col" className="px-3 py-2 text-right font-semibold">
                            Total
                          </th>
                        </>
                      ) : (
                        <th scope="col" className="px-3 py-2 text-right font-medium">
                          {valueTitle}
                        </th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.rows.slice(0, PREVIEW_ROWS).map((r) => (
                      <tr key={r.key} className="border-t border-slate-100">
                        <th scope="row" className="px-3 py-1.5 text-left font-normal text-slate-800">
                          {r.label}
                        </th>
                        {colF &&
                          shownCols.map((c) => (
                            <td key={c.key} className="px-3 py-1.5 text-right tabular-nums text-slate-700">
                              {fmt(preview.cells.get(`${r.key}\u0000${c.key}`))}
                            </td>
                          ))}
                        {colF && preview.cols.length > shownCols.length && <td />}
                        <td className={`px-3 py-1.5 text-right tabular-nums ${colF ? "font-medium text-slate-800" : "text-slate-700"}`}>{fmt(preview.rowTotals.get(r.key))}</td>
                      </tr>
                    ))}
                    <tr className="border-t-2 border-slate-300 bg-slate-50 font-semibold">
                      <th scope="row" className="px-3 py-1.5 text-left text-slate-800">
                        Grand total
                      </th>
                      {colF &&
                        shownCols.map((c) => (
                          <td key={c.key} className="px-3 py-1.5 text-right tabular-nums text-slate-800">
                            {fmt(preview.colTotals.get(c.key))}
                          </td>
                        ))}
                      {colF && preview.cols.length > shownCols.length && <td />}
                      <td className="px-3 py-1.5 text-right tabular-nums text-slate-800">{fmt(preview.grand)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="mt-1.5 text-xs text-slate-500">
                {preview.totalRows > PREVIEW_ROWS ? `Showing the first ${PREVIEW_ROWS} of ${preview.totalRows} groups. ` : `${preview.totalRows} group${preview.totalRows === 1 ? "" : "s"}. `}
                {preview.cols.length > shownCols.length ? `${preview.cols.length - shownCols.length} more columns in the summary tab. ` : ""}
                Rows with a blank or formula {zones.rows}
                {colF ? ` or ${zones.columns}` : ""} aren&apos;t grouped.
              </p>
            </>
          )
        )}
      </section>

      {drag?.active && (
        <div
          aria-hidden="true"
          className="pointer-events-none fixed z-[60] rounded-full border border-brand bg-white px-2.5 py-1 text-xs text-brand shadow-lg"
          style={{ left: drag.x + 8, top: drag.y + 8 }}
        >
          {drag.field}
        </div>
      )}
    </Modal>
  );
}
