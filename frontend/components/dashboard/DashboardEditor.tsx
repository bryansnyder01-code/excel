"use client";

import { useState } from "react";
import { Crosshair } from "lucide-react";
import { inputClass } from "@/components/ui/Modal";
import { allCharts, formatCellRef, parseCellRef } from "@/lib/dashboard";
import { useStore } from "@/lib/store";
import type { DashboardTile } from "@/lib/types";

export type TileDraft = Pick<DashboardTile, "kind" | "title" | "ref" | "compare_ref" | "sheet" | "chart" | "text">;

const KINDS: { value: DashboardTile["kind"]; label: string }[] = [
  { value: "kpi", label: "Number (KPI)" },
  { value: "chart", label: "Chart" },
  { value: "text", label: "Text" },
];

/** The selected grid cell as a reference like "Summary!B4". */
function selectedRef(): string {
  const { workbook, active, selection } = useStore.getState();
  const r = Math.min(selection.r1, selection.r2);
  const c = Math.min(selection.c1, selection.c2);
  return formatCellRef(workbook.sheets[active].name, r, c);
}

/**
 * Add a tile, or change an existing one's settings. Inline panel above the dashboard grid.
 */
export default function TileForm({
  initial,
  onSave,
  onCancel,
}: {
  initial?: DashboardTile;
  onSave: (draft: TileDraft) => void;
  onCancel: () => void;
}) {
  const workbook = useStore((s) => s.workbook);
  const charts = allCharts(workbook);
  const [kind, setKind] = useState<DashboardTile["kind"]>(initial?.kind ?? "kpi");
  const [title, setTitle] = useState(initial?.title ?? "");
  const [ref, setRef] = useState(initial?.ref ?? "");
  const [compare, setCompare] = useState(initial?.compare_ref ?? "");
  const firstChart = charts[0] ? `${charts[0].sheet}\u0000${charts[0].index}` : "";
  const [chartKey, setChartKey] = useState(initial?.kind === "chart" ? `${initial.sheet}\u0000${initial.chart ?? 0}` : firstChart);
  const [text, setText] = useState(initial?.text ?? "");

  const refOk = !!parseCellRef(workbook, ref, useStore.getState().active);
  const compareOk = !compare.trim() || !!parseCellRef(workbook, compare, parseCellRef(workbook, ref)?.sheet ?? 0);
  const chosen = charts.find((c) => `${c.sheet}\u0000${c.index}` === chartKey);
  const valid = kind === "kpi" ? refOk && compareOk : kind === "chart" ? !!chosen : !!text.trim() || !!title.trim();

  // A bare "B4" is stored with the active sheet's name so it keeps pointing at the same place.
  const qualify = (raw: string) => {
    const at = parseCellRef(workbook, raw, useStore.getState().active);
    return at ? formatCellRef(workbook.sheets[at.sheet].name, at.r, at.c) : raw.trim();
  };

  const save = () => {
    if (!valid) return;
    if (kind === "kpi") {
      const at = parseCellRef(workbook, ref, useStore.getState().active)!;
      const fallback = workbook.sheets[at.sheet].columns[at.c]?.header || qualify(ref);
      onSave({ kind, title: title.trim() || fallback, ref: qualify(ref), compare_ref: compare.trim() ? qualifyCompare(compare, at.sheet) : null });
    } else if (kind === "chart") {
      onSave({ kind, title: title.trim() || chosen!.title, sheet: chosen!.sheet, chart: chosen!.index });
    } else onSave({ kind, title: title.trim(), text: text.trim() });
  };

  const qualifyCompare = (raw: string, sheet: number) => {
    const at = parseCellRef(workbook, raw, sheet);
    return at ? formatCellRef(workbook.sheets[at.sheet].name, at.r, at.c) : raw.trim();
  };

  return (
    <form
      className="mb-4 rounded-xl border border-slate-200 bg-slate-50 p-4"
      aria-label={initial ? "Edit tile" : "Add a tile"}
      onSubmit={(e) => {
        e.preventDefault();
        save();
      }}
      onKeyDown={(e) => e.key === "Escape" && (e.stopPropagation(), onCancel())}
    >
      <h3 className="mb-3 text-sm font-semibold text-slate-800">{initial ? "Edit tile" : "Add a tile"}</h3>
      {!initial && (
        <div role="radiogroup" aria-label="Tile type" className="mb-3 flex flex-wrap gap-2">
          {KINDS.map((k) => (
            <label
              key={k.value}
              className={`cursor-pointer rounded-md border px-3 py-1.5 text-sm has-[:focus-visible]:ring-2 has-[:focus-visible]:ring-brand/40 ${
                kind === k.value ? "border-brand bg-white font-medium text-brand" : "border-slate-300 bg-white text-slate-600 hover:bg-slate-100"
              }`}
            >
              <input type="radio" name="tile-kind" className="sr-only" checked={kind === k.value} onChange={() => setKind(k.value)} />
              {k.label}
            </label>
          ))}
        </div>
      )}
      <div className="grid gap-3 sm:grid-cols-2">
        <label className="block">
          <span className="mb-1 block text-xs font-medium text-slate-600">Title</span>
          <input className={inputClass} value={title} onChange={(e) => setTitle(e.target.value)} placeholder={kind === "kpi" ? "e.g. Total revenue" : "Optional"} />
        </label>
        {kind === "kpi" && (
          <>
            <div>
              <label htmlFor="tile-ref" className="mb-1 block text-xs font-medium text-slate-600">
                Cell
              </label>
              <div className="flex gap-2">
                <input
                  id="tile-ref"
                  className={`${inputClass} font-mono`}
                  value={ref}
                  onChange={(e) => setRef(e.target.value)}
                  placeholder="Summary!B4"
                  aria-invalid={!!ref && !refOk}
                  aria-describedby="tile-ref-hint"
                />
                <button
                  type="button"
                  className="flex shrink-0 items-center gap-1 rounded-md border border-slate-300 bg-white px-2.5 text-xs text-slate-700 hover:bg-slate-100"
                  onClick={() => setRef(selectedRef())}
                  title="Use the cell selected in the sheet"
                >
                  <Crosshair className="h-3.5 w-3.5" /> Use selected cell
                </button>
              </div>
              <span id="tile-ref-hint" className={`mt-1 block text-xs ${ref && !refOk ? "text-red-600" : "text-slate-500"}`}>
                {ref && !refOk ? "That isn't a cell in this workbook. Try Sheet!B4." : `Selected now: ${selectedRef()}`}
              </span>
            </div>
            <label className="block sm:col-span-2">
              <span className="mb-1 block text-xs font-medium text-slate-600">Compare with (optional)</span>
              <input
                className={`${inputClass} font-mono`}
                value={compare}
                onChange={(e) => setCompare(e.target.value)}
                placeholder="e.g. last year's figure, Summary!C4"
                aria-invalid={!compareOk}
              />
              <span className={`mt-1 block text-xs ${compareOk ? "text-slate-500" : "text-red-600"}`}>
                {compareOk ? "Shows the change as ▲ / ▼ percent." : "That isn't a cell in this workbook."}
              </span>
            </label>
          </>
        )}
        {kind === "chart" && (
          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-600">Chart</span>
            {charts.length ? (
              <select className={inputClass} value={chartKey} onChange={(e) => setChartKey(e.target.value)}>
                {charts.map((c) => (
                  <option key={`${c.sheet}-${c.index}`} value={`${c.sheet}\u0000${c.index}`}>
                    {c.title} ({c.sheet})
                  </option>
                ))}
              </select>
            ) : (
              <span className="block text-sm text-slate-500">This workbook has no charts yet. Add one from the toolbar first.</span>
            )}
          </label>
        )}
        {kind === "text" && (
          <label className="block sm:col-span-2">
            <span className="mb-1 block text-xs font-medium text-slate-600">Text</span>
            <textarea className={inputClass} rows={3} value={text} onChange={(e) => setText(e.target.value)} placeholder="A note for readers of this dashboard" />
          </label>
        )}
      </div>
      <div className="mt-3 flex justify-end gap-2">
        <button type="button" className="rounded-md border border-slate-300 bg-white px-3 py-1.5 text-sm text-slate-700 hover:bg-slate-100" onClick={onCancel}>
          Cancel
        </button>
        <button type="submit" disabled={!valid} className="rounded-md bg-brand px-3 py-1.5 text-sm font-medium text-white hover:bg-brand-dark disabled:opacity-50">
          {initial ? "Save tile" : "Add tile"}
        </button>
      </div>
    </form>
  );
}
