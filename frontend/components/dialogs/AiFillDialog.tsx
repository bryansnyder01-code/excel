"use client";

import { useMemo, useRef, useState } from "react";
import { Loader2, Sparkles } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { aiFill, AI_FILL_MAX_ROWS, errorText, type AiFillOutput } from "@/lib/aiApi";
import { evaluateCached } from "@/lib/formula";
import { columnLetter, formatValue, numberStyleAt } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { CellSpec, SheetSpec, WorkbookSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";

const NEW = -1;
const PREVIEW_ROWS = 5;

const isEmpty = (cell: CellSpec | undefined) => cell?.value === null || cell?.value === undefined || cell.value === "";

/** What a cell shows, formulas calculated. Grid row r >= 1. */
function shown(wb: WorkbookSpec, si: number, sheet: SheetSpec, r: number, c: number): string {
  const v = sheet.rows[r - 1]?.cells[c]?.value;
  const value = typeof v === "string" && v.startsWith("=") ? evaluateCached(wb).get(`${si}:${r + 1}:${c + 1}`) : v;
  return formatValue(value, numberStyleAt(sheet, r, c)).text;
}

/** Fill a column, row by row, from an instruction and the row's other columns. */
export default function AiFillDialog() {
  const close = useUi((s) => s.close);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const selection = useStore((s) => s.selection);
  const sheet = workbook.sheets[active];
  const sheetName = sheet.name;
  const ncols = sheet.columns.length;

  const initialTarget = Math.min(selection.c1, selection.c2) < ncols ? Math.min(selection.c1, selection.c2) : NEW;
  const [target, setTarget] = useState<number>(initialTarget);
  const [inputs, setInputs] = useState<number[]>(() =>
    Array.from({ length: initialTarget === NEW ? ncols : initialTarget }, (_, i) => i),
  );
  const [newHeader, setNewHeader] = useState("AI result");
  const [instruction, setInstruction] = useState("");
  const [output, setOutput] = useState<AiFillOutput>("text");
  const [onlyEmpty, setOnlyEmpty] = useState(true);
  const [preview, setPreview] = useState<{ rows: number[]; values: (string | number | null)[] } | null>(null);
  const [busy, setBusy] = useState<null | "preview" | "run">(null);
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState<string | null>(null);
  const cancelled = useRef(false);

  const usedInputs = inputs.filter((c) => c !== target && c < ncols).sort((a, b) => a - b);

  // Data rows to fill: grid rows (1-based) with at least one input, and an empty target if asked.
  const rows = useMemo(() => {
    const out: number[] = [];
    sheet.rows.forEach((row, i) => {
      if ((row.kind ?? "data") !== "data") return;
      if (!usedInputs.some((c) => !isEmpty(row.cells[c]))) return;
      if (onlyEmpty && target !== NEW && !isEmpty(row.cells[target])) return;
      out.push(i + 1);
    });
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, usedInputs.join(","), onlyEmpty, target]);

  const colName = (c: number) => `${columnLetter(c)}${sheet.columns[c]?.header ? ` – ${sheet.columns[c].header}` : ""}`;
  const headers = usedInputs.map((c) => sheet.columns[c]?.header || columnLetter(c));
  const rowValues = (wb: WorkbookSpec, rs: number[]) => {
    const si = wb.sheets.findIndex((s) => s.name === sheetName);
    const sh = wb.sheets[si];
    return rs.map((r) => usedInputs.map((c) => shown(wb, si, sh, r, c)));
  };

  const invalid = !instruction.trim() || !usedInputs.length || (target === NEW && !newHeader.trim());

  const runPreview = async () => {
    setError(null);
    setDone(null);
    const sample = rows.slice(0, PREVIEW_ROWS);
    if (!sample.length) return setError("There are no rows to fill with these settings.");
    setBusy("preview");
    try {
      const res = await aiFill(instruction.trim(), headers, rowValues(workbook, sample), output);
      setPreview({ rows: sample, values: res.values });
    } catch (e) {
      setError(errorText(e, "The preview didn't work."));
    } finally {
      setBusy(null);
    }
  };

  const write = (results: Map<number, string | number | null>) => {
    if (!results.size) return 0;
    let n = 0;
    useStore.getState().mutate((wb) => {
      const sh = wb.sheets.find((s) => s.name === sheetName);
      if (!sh) return;
      let col = target;
      if (target === NEW) {
        col = sh.columns.length;
        sh.columns.push({ header: newHeader.trim(), format: "text" });
      }
      for (const [r, value] of results) {
        const row = sh.rows[r - 1];
        if (!row || value === null || value === undefined || value === "") continue;
        while (row.cells.length <= col) row.cells.push({ value: null });
        row.cells[col] = { ...row.cells[col], value: output === "number" && typeof value === "string" && /^-?\d+(\.\d+)?$/.test(value.trim()) ? parseFloat(value) : value };
        n++;
      }
    });
    return n;
  };

  const run = async () => {
    setError(null);
    setDone(null);
    if (!rows.length) return setError("There are no rows to fill with these settings.");
    cancelled.current = false;
    setBusy("run");
    setProgress({ done: 0, total: rows.length });
    // Values from the preview are reused, so those rows aren't asked again.
    const results = new Map<number, string | number | null>();
    if (preview) preview.rows.forEach((r, i) => rows.includes(r) && results.set(r, preview.values[i] ?? null));
    const todo = rows.filter((r) => !results.has(r));
    const snapshot = useStore.getState().workbook;
    let failure: string | null = null;
    setProgress({ done: results.size, total: rows.length });
    for (let i = 0; i < todo.length && !cancelled.current; i += AI_FILL_MAX_ROWS) {
      const batch = todo.slice(i, i + AI_FILL_MAX_ROWS);
      try {
        const res = await aiFill(instruction.trim(), headers, rowValues(snapshot, batch), output);
        batch.forEach((r, j) => results.set(r, res.values[j] ?? null));
        setProgress({ done: results.size, total: rows.length });
      } catch (e) {
        failure = errorText(e, "The AI request failed.");
        break;
      }
    }
    const n = write(results);
    setBusy(null);
    setPreview(null);
    const what = `Filled ${n} ${n === 1 ? "cell" : "cells"}${target === NEW ? ` in the new "${newHeader.trim()}" column` : ` in column ${columnLetter(target)}`}.`;
    if (failure) setError(`${failure} ${n ? what : "Nothing was written."}`);
    else if (cancelled.current) setDone(`Stopped. ${what}`);
    else setDone(what);
    if (n) useUi.getState().notify(`${what} Undo with Ctrl+Z.`);
  };

  const pct = progress.total ? Math.round((progress.done / progress.total) * 100) : 0;

  return (
    <Modal
      title="Fill column with AI"
      onClose={() => {
        cancelled.current = true;
        close();
      }}
      width="max-w-2xl"
      footer={
        busy === "run" ? (
          <Button onClick={() => (cancelled.current = true)}>Cancel</Button>
        ) : (
          <>
            <Button onClick={close}>Close</Button>
            <Button disabled={invalid || !!busy} onClick={runPreview}>
              {busy === "preview" ? "Previewing..." : `Preview ${Math.min(PREVIEW_ROWS, rows.length) || ""} rows`}
            </Button>
            <Button primary disabled={invalid || !!busy || !rows.length} onClick={run}>
              Fill {rows.length} {rows.length === 1 ? "row" : "rows"}
            </Button>
          </>
        )
      }
    >
      <Field label="Instruction" hint='For example "Categorise this expense as Food, Travel, Office or Other" or "Translate to French".'>
        <textarea
          className={`${inputClass} min-h-[4.5rem]`}
          value={instruction}
          autoFocus
          disabled={busy === "run"}
          onChange={(e) => {
            setInstruction(e.target.value);
            setPreview(null);
          }}
        />
      </Field>

      <div className="grid gap-4 sm:grid-cols-2">
        <div className="mb-4">
          <span className="mb-1 block text-sm font-medium text-slate-700">Use these columns</span>
          <div className="max-h-40 overflow-y-auto rounded-md border border-slate-200 p-2" role="group" aria-label="Input columns">
            {sheet.columns.map((_, c) =>
              c === target ? null : (
                <label key={c} className="flex items-center gap-2 py-0.5 text-sm text-slate-700">
                  <input
                    type="checkbox"
                    className="accent-brand"
                    disabled={busy === "run"}
                    checked={inputs.includes(c)}
                    onChange={(e) => {
                      setInputs(e.target.checked ? [...inputs, c] : inputs.filter((x) => x !== c));
                      setPreview(null);
                    }}
                  />
                  <span className="truncate">{colName(c)}</span>
                </label>
              ),
            )}
          </div>
        </div>
        <div>
          <Field label="Write results to">
            <select
              className={inputClass}
              value={target}
              disabled={busy === "run"}
              onChange={(e) => {
                setTarget(Number(e.target.value));
                setPreview(null);
              }}
            >
              {sheet.columns.map((_, c) => (
                <option key={c} value={c}>
                  {colName(c)}
                </option>
              ))}
              <option value={NEW}>A new column at the end</option>
            </select>
          </Field>
          {target === NEW && (
            <Field label="New column header">
              <input className={inputClass} value={newHeader} onChange={(e) => setNewHeader(e.target.value)} disabled={busy === "run"} />
            </Field>
          )}
          <div className="mb-3 flex gap-4 text-sm text-slate-700" role="radiogroup" aria-label="Result type">
            {(["text", "number"] as AiFillOutput[]).map((o) => (
              <label key={o} className="flex items-center gap-1.5">
                <input type="radio" className="accent-brand" checked={output === o} onChange={() => (setOutput(o), setPreview(null))} disabled={busy === "run"} />
                {o === "text" ? "Text" : "Numbers"}
              </label>
            ))}
          </div>
          {target !== NEW && (
            <label className="flex items-center gap-2 text-sm text-slate-700">
              <input type="checkbox" className="accent-brand" checked={onlyEmpty} onChange={(e) => setOnlyEmpty(e.target.checked)} disabled={busy === "run"} />
              Only fill empty cells
            </label>
          )}
        </div>
      </div>

      <p className="mb-3 text-xs text-slate-500">
        {rows.length} data {rows.length === 1 ? "row" : "rows"} on {sheetName} will be filled. Total, subtotal and section rows are skipped. Everything
        is written in one step you can undo.
      </p>

      {busy === "run" && (
        <div className="mb-3">
          <div className="mb-1 flex items-center gap-2 text-sm text-slate-600">
            <Loader2 className="h-4 w-4 animate-spin" /> Filled {progress.done} of {progress.total} rows...
          </div>
          <div className="h-2 overflow-hidden rounded-full bg-slate-100" role="progressbar" aria-valuenow={pct} aria-valuemin={0} aria-valuemax={100} aria-label="Rows filled">
            <div className="h-full bg-brand transition-all" style={{ width: `${pct}%` }} />
          </div>
        </div>
      )}

      {error && <p className="mb-3 text-sm text-red-600">{error}</p>}
      {done && <p className="mb-3 text-sm text-emerald-700">{done}</p>}

      {preview && (
        <div className="overflow-x-auto rounded-md border border-slate-200">
          <table className="w-full text-left text-xs">
            <caption className="px-3 py-2 text-left text-xs font-medium text-slate-600">
              <Sparkles className="mr-1 inline h-3.5 w-3.5 text-violet-600" />
              Preview of the first {preview.rows.length} rows (nothing is written yet)
            </caption>
            <thead className="bg-slate-50 text-slate-600">
              <tr>
                {headers.map((h, i) => (
                  <th key={i} className="px-3 py-1.5 font-medium">
                    {h}
                  </th>
                ))}
                <th className="px-3 py-1.5 font-medium text-violet-700">{target === NEW ? newHeader : sheet.columns[target]?.header || columnLetter(target)}</th>
              </tr>
            </thead>
            <tbody>
              {rowValues(workbook, preview.rows).map((vals, i) => (
                <tr key={i} className="border-t border-slate-100">
                  {vals.map((v, j) => (
                    <td key={j} className="max-w-[12rem] truncate px-3 py-1.5 text-slate-700">
                      {v}
                    </td>
                  ))}
                  <td className="px-3 py-1.5 font-medium text-slate-900">{preview.values[i] ?? <span className="text-slate-400">(blank)</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Modal>
  );
}
