"use client";

import { useMemo, useState } from "react";
import { Eye, Plus, Target, Trash2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { normalize, useStore } from "@/lib/store";
import { evaluateCached, type Computed } from "@/lib/formula";
import { columnLetter, formatValue, numberStyleAt } from "@/lib/format";
import type { ScenarioSpec, Selection, WorkbookSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";
import {
  applyScenario,
  formatCellRef,
  goalSeek,
  goalSeekProblem,
  isFormulaValue,
  normalizeRef,
  parseCellRef,
  rawValue,
  scenarioResults,
  setCellValue,
  splitRefs,
  valueIn,
  type CellAddr,
  type GoalSeekResult,
} from "@/lib/whatif";

interface Draft {
  name: string;
  /** Typed values keyed by canonical ref; empty means "keep the current value". */
  values: Record<string, string>;
}

const MAX_DEFAULT_CELLS = 20;

/** Cells of the selection (data rows only), split into constants and formulas. */
function selectionCells(wb: WorkbookSpec, si: number, sel: Selection) {
  const s = normalize(sel);
  const inputs: string[] = [];
  const results: string[] = [];
  for (let r = Math.max(1, s.r1); r <= s.r2; r++)
    for (let c = s.c1; c <= s.c2; c++) {
      const a: CellAddr = { si, row: r + 1, col: c };
      const v = rawValue(wb, a);
      if (isFormulaValue(v)) results.length < MAX_DEFAULT_CELLS && results.push(formatCellRef(wb, a));
      else if (typeof v === "number" || v === null) inputs.length < MAX_DEFAULT_CELLS && inputs.push(formatCellRef(wb, a));
    }
  return { inputs, results };
}

/** A short label for a cell: the nearest text to its left in the same row, else its column header. */
function labelFor(wb: WorkbookSpec, a: CellAddr): string {
  const sheet = wb.sheets[a.si];
  const row = sheet.rows[a.row - 2];
  for (let c = a.col - 1; c >= 0; c--) {
    const v = row?.cells[c]?.value;
    if (typeof v === "string" && v.trim() && !v.startsWith("=")) return v.trim();
  }
  return sheet.columns[a.col]?.header ?? "";
}

function show(wb: WorkbookSpec, a: CellAddr | null, v: Computed | undefined): string {
  if (!a) return "";
  return formatValue(v, numberStyleAt(wb.sheets[a.si], a.row - 1, a.col)).text;
}

const parseNum = (s: string): number | null => {
  const t = s.trim().replace(/,/g, "");
  if (!t) return null;
  const pct = t.endsWith("%");
  const n = Number(pct ? t.slice(0, -1) : t);
  return Number.isFinite(n) ? (pct ? n / 100 : n) : null;
};

/** Parse a comma-separated list of refs into canonical refs and the ones that didn't parse. */
function readRefs(wb: WorkbookSpec, text: string, si: number) {
  const ok: string[] = [];
  const bad: string[] = [];
  for (const r of splitRefs(text)) {
    const n = normalizeRef(wb, r, si);
    if (!n) bad.push(r);
    else if (!ok.includes(n)) ok.push(n);
  }
  return { ok, bad };
}

/** Scenarios (compare named sets of inputs) and Goal Seek (solve for an input). */
export default function WhatIfDialog() {
  const close = useUi((s) => s.close);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const [tab, setTab] = useState<"scenarios" | "goal">("scenarios");

  return (
    <Modal title="What-if analysis" onClose={close} width="max-w-3xl">
      <div role="tablist" aria-label="What-if tools" className="mb-4 flex gap-1 border-b border-slate-200">
        {(
          [
            ["scenarios", "Scenarios"],
            ["goal", "Goal seek"],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            role="tab"
            type="button"
            aria-selected={tab === id}
            onClick={() => setTab(id)}
            className={`-mb-px border-b-2 px-3 py-2 text-sm font-medium ${
              tab === id ? "border-b-brand text-brand" : "border-transparent text-slate-500 hover:text-slate-700"
            }`}
          >
            {label}
          </button>
        ))}
      </div>
      {tab === "scenarios" ? <Scenarios workbook={workbook} active={active} onClose={close} /> : <GoalSeek workbook={workbook} active={active} onClose={close} />}
    </Modal>
  );
}

// ---------------------------------------------------------------- scenarios

function Scenarios({ workbook, active, onClose }: { workbook: WorkbookSpec; active: number; onClose: () => void }) {
  const [init] = useState(() => {
    const st = useStore.getState();
    const fromSel = selectionCells(st.workbook, st.active, st.selection);
    const saved = st.workbook.scenarios ?? [];
    const savedKeys = [...new Set(saved.flatMap((s) => Object.keys(s.values)))];
    return {
      inputs: (savedKeys.length ? savedKeys : fromSel.inputs).join(", "),
      results: fromSel.results.join(", "),
      drafts: saved.map((s) => ({
        name: s.name,
        values: Object.fromEntries(Object.entries(s.values).map(([k, v]) => [normalizeRef(st.workbook, k) ?? k, String(v)])),
      })),
    };
  });
  const [inputsText, setInputsText] = useState(init.inputs);
  const [resultsText, setResultsText] = useState(init.results);
  const [drafts, setDrafts] = useState<Draft[]>(init.drafts);

  const inputs = useMemo(() => readRefs(workbook, inputsText, active), [workbook, inputsText, active]);
  const results = useMemo(() => readRefs(workbook, resultsText, active), [workbook, resultsText, active]);
  const values = evaluateCached(workbook);
  const addr = (ref: string) => parseCellRef(workbook, ref);
  const formulaInputs = inputs.ok.filter((r) => {
    const a = addr(r);
    return a && isFormulaValue(rawValue(workbook, a));
  });
  const nonFormulaResults = results.ok.filter((r) => {
    const a = addr(r);
    return a && !isFormulaValue(rawValue(workbook, a));
  });
  const usableInputs = inputs.ok.filter((r) => !formulaInputs.includes(r));

  const currentNumber = (ref: string): number | null => {
    const a = addr(ref);
    const v = a ? rawValue(workbook, a) : null;
    return typeof v === "number" ? v : v === null ? 0 : null;
  };

  /** A draft as a full ScenarioSpec: blanks take the cell's current value. */
  const toSpec = (d: Draft): ScenarioSpec => {
    const out: Record<string, number> = {};
    for (const ref of usableInputs) {
      const typed = parseNum(d.values[ref] ?? "");
      const v = typed ?? currentNumber(ref);
      if (v !== null) out[ref] = v;
    }
    return { name: d.name.trim() || "Scenario", values: out };
  };
  const specs = drafts.map(toSpec);
  const badValues = drafts.some((d) => usableInputs.some((ref) => (d.values[ref] ?? "").trim() !== "" && parseNum(d.values[ref]) === null));

  const specsKey = JSON.stringify(specs);
  const table = useMemo(() => scenarioResults(workbook, JSON.parse(specsKey) as ScenarioSpec[], results.ok), [workbook, specsKey, results.ok]);

  const addDraft = (name: string) =>
    setDrafts((ds) => [...ds, { name, values: Object.fromEntries(usableInputs.map((r) => [r, String(currentNumber(r) ?? "")])) }]);
  const setValue = (i: number, ref: string, v: string) =>
    setDrafts((ds) => ds.map((d, j) => (j === i ? { ...d, values: { ...d.values, [ref]: v } } : d)));

  const showScenario = (i: number) => {
    const spec = specs[i];
    let n = 0;
    useStore.getState().mutate((wb) => {
      n = applyScenario(wb, spec);
    });
    useUi.getState().notify(`Showing "${spec.name}" (${n} ${n === 1 ? "cell" : "cells"} changed). Undo with Ctrl+Z.`);
  };
  const save = () => {
    useStore.getState().mutate((wb) => {
      wb.scenarios = specs.length ? specs : undefined;
    });
    useUi.getState().notify(specs.length ? `Saved ${specs.length} ${specs.length === 1 ? "scenario" : "scenarios"} with the workbook.` : "Removed saved scenarios.");
  };

  const cellHead = (ref: string) => {
    const a = addr(ref);
    const label = a ? labelFor(workbook, a) : "";
    return (
      <th scope="row" className="py-1.5 pr-3 text-left align-middle font-normal">
        <span className="block font-mono text-xs text-slate-500">{ref}</span>
        {label && <span className="block max-w-40 truncate text-slate-700">{label}</span>}
      </th>
    );
  };

  return (
    <div>
      <div className="grid gap-x-4 sm:grid-cols-2">
        <Field label="Changing cells (inputs)" hint="Comma-separated, e.g. B2, B3 or Inputs!B2. Defaults to your selection.">
          <input aria-label="Changing cells" className={inputClass} value={inputsText} onChange={(e) => setInputsText(e.target.value)} placeholder="B2, B3" />
        </Field>
        <Field label="Result cells to watch" hint="Cells with formulas, e.g. Income!B10.">
          <input aria-label="Result cells" className={inputClass} value={resultsText} onChange={(e) => setResultsText(e.target.value)} placeholder="Income!B10" />
        </Field>
      </div>
      {(inputs.bad.length > 0 || results.bad.length > 0 || formulaInputs.length > 0 || nonFormulaResults.length > 0) && (
        <ul className="mb-3 space-y-0.5 rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-800">
          {[...inputs.bad, ...results.bad].map((r) => (
            <li key={`bad-${r}`}>&ldquo;{r}&rdquo; isn&rsquo;t a cell reference (header cells can&rsquo;t be used).</li>
          ))}
          {formulaInputs.map((r) => (
            <li key={`f-${r}`}>{r} has a formula, so it can&rsquo;t be an input and is ignored.</li>
          ))}
          {nonFormulaResults.map((r) => (
            <li key={`n-${r}`}>{r} has no formula, so it won&rsquo;t change between scenarios.</li>
          ))}
        </ul>
      )}

      <div className="mb-3 flex flex-wrap items-center gap-2">
        <Button onClick={() => addDraft(`Scenario ${drafts.length + 1}`)} disabled={!usableInputs.length}>
          <span className="inline-flex items-center gap-1.5">
            <Plus className="h-4 w-4" /> Add scenario
          </span>
        </Button>
        {drafts.length === 0 && (
          <Button onClick={() => ["Best case", "Base case", "Worst case"].forEach(addDraft)} disabled={!usableInputs.length}>
            Add Best / Base / Worst
          </Button>
        )}
        <span className="text-xs text-slate-500">New scenarios start from the current values; edit them in the table.</span>
      </div>

      {usableInputs.length > 0 && (
        <div className="overflow-x-auto rounded-md border border-slate-200">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-xs text-slate-500">
              <tr>
                <th className="px-3 py-2 text-left font-medium">Cell</th>
                <th className="px-2 py-2 text-right font-medium">Current</th>
                {drafts.map((d, i) => (
                  <th key={i} className="min-w-32 px-2 py-2 text-right font-medium">
                    <input
                      aria-label={`Scenario ${i + 1} name`}
                      className="w-full rounded border border-transparent bg-transparent px-1 py-0.5 text-right text-sm font-medium text-slate-800 hover:border-slate-300 focus:border-brand focus:outline-none"
                      value={d.name}
                      onChange={(e) => setDrafts((ds) => ds.map((x, j) => (j === i ? { ...x, name: e.target.value } : x)))}
                    />
                    <div className="mt-1 flex justify-end gap-1">
                      <button
                        type="button"
                        title="Show this scenario on the sheet"
                        aria-label={`Show ${d.name}`}
                        disabled={badValues}
                        onClick={() => showScenario(i)}
                        className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-xs text-slate-600 hover:bg-slate-100 disabled:opacity-50"
                      >
                        <Eye className="h-3.5 w-3.5" /> Show
                      </button>
                      <button
                        type="button"
                        aria-label={`Delete ${d.name}`}
                        onClick={() => setDrafts((ds) => ds.filter((_, j) => j !== i))}
                        className="rounded p-0.5 text-slate-400 hover:bg-red-50 hover:text-red-600"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              <tr>
                <td colSpan={2 + drafts.length} className="bg-slate-50/60 px-3 pt-2 pb-1 text-xs font-semibold tracking-wide text-slate-500 uppercase">
                  Inputs
                </td>
              </tr>
              {usableInputs.map((ref) => {
                const a = addr(ref);
                return (
                  <tr key={ref}>
                    <td className="px-3">{cellHead(ref)}</td>
                    <td className="px-2 text-right text-slate-600 tabular-nums">{a ? show(workbook, a, rawValue(workbook, a)) : ""}</td>
                    {drafts.map((d, i) => {
                      const text = d.values[ref] ?? "";
                      const bad = text.trim() !== "" && parseNum(text) === null;
                      return (
                        <td key={i} className="px-2 py-1">
                          <input
                            aria-label={`${d.name} value for ${ref}`}
                            inputMode="decimal"
                            className={`w-full rounded border px-2 py-1 text-right text-sm tabular-nums focus:border-brand focus:outline-none ${bad ? "border-red-600" : "border-slate-300"}`}
                            value={text}
                            placeholder={String(currentNumber(ref) ?? "")}
                            onChange={(e) => setValue(i, ref, e.target.value)}
                          />
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
              <tr>
                <td colSpan={2 + drafts.length} className="bg-slate-50/60 px-3 pt-2 pb-1 text-xs font-semibold tracking-wide text-slate-500 uppercase">
                  Results
                </td>
              </tr>
              {results.ok.length === 0 && (
                <tr>
                  <td colSpan={2 + drafts.length} className="px-3 py-2 text-sm text-slate-500">
                    Add result cells above to compare them.
                  </td>
                </tr>
              )}
              {results.ok.map((ref, k) => {
                const a = addr(ref);
                const now = a ? valueIn(workbook, values, a) : undefined;
                return (
                  <tr key={ref}>
                    <td className="px-3">{cellHead(ref)}</td>
                    <td className="px-2 text-right text-slate-600 tabular-nums">{show(workbook, a, now)}</td>
                    {drafts.map((_, i) => {
                      const v = table[i]?.[k];
                      const changed = typeof v === "number" && typeof now === "number" && Math.abs(v - now) > 1e-9 * Math.max(1, Math.abs(now));
                      return (
                        <td key={i} className={`px-3 py-1.5 text-right tabular-nums ${changed ? "font-semibold text-slate-900" : "text-slate-700"}`}>
                          {show(workbook, a, v)}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {badValues && <p className="mt-2 text-xs text-red-600">Some scenario values aren&rsquo;t numbers.</p>}

      <div className="mt-4 flex justify-end gap-2 border-t border-slate-200 pt-3">
        <Button onClick={onClose}>Close</Button>
        <Button primary disabled={badValues || (!drafts.length && !(workbook.scenarios ?? []).length)} onClick={save}>
          Save scenarios
        </Button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------- goal seek

function GoalSeek({ workbook, active, onClose }: { workbook: WorkbookSpec; active: number; onClose: () => void }) {
  const [init] = useState(() => {
    const st = useStore.getState();
    const s = normalize(st.selection);
    const here = formatCellRef(st.workbook, { si: st.active, row: Math.max(1, s.r1) + 1, col: s.c1 });
    const isF = isFormulaValue(rawValue(st.workbook, { si: st.active, row: Math.max(1, s.r1) + 1, col: s.c1 }));
    return { result: isF ? here : "", input: isF ? "" : here };
  });
  const [resultRef, setResultRef] = useState(init.result);
  const [targetText, setTargetText] = useState("");
  const [inputRef, setInputRef] = useState(init.input);
  const [busy, setBusy] = useState(false);
  const [outcome, setOutcome] = useState<(GoalSeekResult & { result_ref: string; input_ref: string; target: number }) | null>(null);

  const r = normalizeRef(workbook, resultRef, active);
  const i = normalizeRef(workbook, inputRef, active);
  const target = parseNum(targetText);
  const problem = !resultRef.trim() || !inputRef.trim() ? null : goalSeekProblem(workbook, r ?? resultRef, i ?? inputRef);
  const ready = !!r && !!i && target !== null && !problem;

  const solve = () => {
    if (!ready || !r || !i || target === null) return;
    setBusy(true);
    setOutcome(null);
    // Let the "Solving…" state paint before the (synchronous) search runs.
    setTimeout(() => {
      const res = goalSeek(workbook, r, target, i);
      setOutcome({ ...res, result_ref: r, input_ref: i, target });
      setBusy(false);
    }, 20);
  };

  const apply = () => {
    if (!outcome) return;
    const a = parseCellRef(useStore.getState().workbook, outcome.input_ref);
    if (!a) return;
    useStore.getState().mutate((wb) => setCellValue(wb, a, outcome.value));
    useUi.getState().notify(`Set ${outcome.input_ref} to ${outcome.value}. Undo with Ctrl+Z.`);
    setOutcome(null);
  };

  const ra = r ? parseCellRef(workbook, r) : null;
  const ia = i ? parseCellRef(workbook, i) : null;
  const fmtNum = (a: CellAddr | null, v: number | null) => (a && v !== null ? show(workbook, a, v) || String(v) : "–");

  return (
    <div>
      <p className="mb-4 text-sm text-slate-600">
        Find the input value that makes a formula give the answer you want. Works on a copy until you press Apply.
      </p>
      <div className="grid gap-x-4 sm:grid-cols-3">
        <Field label="Set cell" hint="Must contain a formula">
          <input aria-label="Set cell" className={inputClass} value={resultRef} onChange={(e) => setResultRef(e.target.value)} placeholder="Income!B10" />
        </Field>
        <Field label="To value">
          <input aria-label="Target value" inputMode="decimal" className={inputClass} value={targetText} onChange={(e) => setTargetText(e.target.value)} placeholder="10000" />
        </Field>
        <Field label="By changing cell" hint="Must hold a number">
          <input aria-label="By changing cell" className={inputClass} value={inputRef} onChange={(e) => setInputRef(e.target.value)} placeholder="Inputs!B2" />
        </Field>
      </div>
      {ra && (
        <p className="-mt-2 mb-3 text-xs text-slate-500">
          {r} ({labelFor(workbook, ra) || columnLetter(ra.col)}) is now <span className="tabular-nums">{show(workbook, ra, valueIn(workbook, evaluateCached(workbook), ra)) || "empty"}</span>
          {ia && (
            <>
              {" "}
              with {i} = <span className="tabular-nums">{show(workbook, ia, rawValue(workbook, ia)) || "empty"}</span>
            </>
          )}
          .
        </p>
      )}
      {problem && <p className="mb-3 rounded-md bg-amber-50 px-3 py-2 text-sm text-amber-800">{problem}</p>}
      {targetText.trim() !== "" && target === null && <p className="mb-3 text-sm text-red-600">The target must be a number.</p>}

      {outcome && (
        <div aria-live="polite" className={`mb-3 rounded-md border px-3 py-2 text-sm ${outcome.found ? "border-green-100 bg-green-50" : "border-amber-200 bg-amber-50"}`}>
          {outcome.found ? (
            <p className="text-slate-800">
              <span className="font-medium text-green-700">Solution found.</span> Setting {outcome.input_ref} to{" "}
              <strong className="tabular-nums">{Number(outcome.value.toPrecision(12))}</strong> makes {outcome.result_ref} ={" "}
              <span className="tabular-nums">{fmtNum(ra, outcome.result)}</span>.
            </p>
          ) : (
            <p className="text-amber-800">
              {outcome.message}
              {Number.isFinite(outcome.value) && outcome.result !== null && (
                <>
                  {" "}
                  Closest: {outcome.input_ref} = {Number(outcome.value.toPrecision(10))} gives {fmtNum(ra, outcome.result)}.
                </>
              )}
            </p>
          )}
          <p className="mt-0.5 text-xs text-slate-500">{outcome.iterations} evaluations.</p>
        </div>
      )}

      <div className="mt-4 flex justify-end gap-2 border-t border-slate-200 pt-3">
        <Button onClick={onClose}>Close</Button>
        {outcome?.found && (
          <Button onClick={apply}>
            Apply to {outcome.input_ref}
          </Button>
        )}
        <Button primary disabled={!ready || busy} onClick={solve}>
          <span className="inline-flex items-center gap-1.5">
            <Target className="h-4 w-4" /> {busy ? "Solving…" : "Solve"}
          </span>
        </Button>
      </div>
    </div>
  );
}
