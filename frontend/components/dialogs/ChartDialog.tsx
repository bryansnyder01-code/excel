"use client";

import { useEffect, useState } from "react";
import { Plus, RotateCcw, Trash2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { clearChartEdit, useChartEdit } from "@/components/charts/chartEdit";
import { SERIES_COLORS } from "@/components/charts/MiniChart";
import { columnLetter } from "@/lib/format";
import { normalize, useStore } from "@/lib/store";
import type { ChartSeriesSpec, ChartSpec, ChartType } from "@/lib/types";
import { useUi } from "@/lib/ui";

const TYPES: { value: ChartType; label: string; hint: string }[] = [
  { value: "column", label: "Column", hint: "Compare categories" },
  { value: "bar", label: "Bar", hint: "Long category names" },
  { value: "line", label: "Line", hint: "Trends over time" },
  { value: "area", label: "Area", hint: "Volume over time" },
  { value: "pie", label: "Pie", hint: "Shares of a whole" },
  { value: "doughnut", label: "Doughnut", hint: "Shares, with a hole" },
  { value: "scatter", label: "Scatter", hint: "Numbers vs numbers" },
];

const STACKABLE: ChartType[] = ["column", "bar", "area"];
const SINGLE_SERIES: ChartType[] = ["pie", "doughnut"];
const RANGE_RE = /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/;

interface Draft {
  type: ChartType;
  title: string;
  categories: string;
  series: ChartSeriesSpec[];
  stacked: boolean;
  xTitle: string;
  yTitle: string;
}

function fromChart(c: ChartSpec): Draft {
  return {
    type: c.type,
    title: c.title,
    categories: c.categories,
    series: c.series.map((s) => ({ name: s.name, range: s.range, color: s.color ?? null, kind: s.kind ?? null })),
    stacked: !!c.stacked,
    xTitle: c.x_title ?? "",
    yTitle: c.y_title ?? "",
  };
}

/** Build the saved spec, dropping options that do not apply to the chosen type. */
function toChart(d: Draft, anchor: string | null | undefined): ChartSpec {
  const single = SINGLE_SERIES.includes(d.type);
  const series = (single ? d.series.slice(0, 1) : d.series).map((s) => ({
    name: s.name.trim() || s.range,
    range: s.range.trim().toUpperCase(),
    color: s.color || null,
    kind: d.type === "column" && s.kind === "line" ? ("line" as const) : null,
  }));
  const chart: ChartSpec = {
    type: d.type,
    title: d.title.trim(),
    categories: d.categories.trim().toUpperCase(),
    series,
    stacked: STACKABLE.includes(d.type) && d.stacked,
    x_title: single ? null : d.xTitle.trim() || null,
    y_title: single ? null : d.yTitle.trim() || null,
  };
  if (anchor !== undefined) chart.anchor = anchor;
  return chart;
}

/**
 * Add or edit a chart. Adding charts the selected cells: the first selected column holds the
 * labels, every other column becomes a series named by its header. Editing (opened from the
 * charts panel) starts from the saved chart and writes it back in place.
 */
export default function ChartDialog() {
  const close = useUi((s) => s.close);
  const sheet = useStore((s) => s.workbook.sheets[s.active]);
  // Read once: the dialog keeps editing the chart it was opened for.
  const [target] = useState(() => useChartEdit.getState().target);
  const existing = target ? useStore.getState().workbook.sheets[target.sheet]?.charts?.[target.index] : undefined;
  const editing = !!(target && existing);
  useEffect(() => clearChartEdit, []);

  const [draft, setDraft] = useState<Draft>(() => {
    if (existing) return fromChart(existing);
    const sel = normalize(useStore.getState().selection);
    const firstRow = Math.max(1, sel.r1);
    const dataRows = sheet.rows.filter((r) => r.kind !== "total").length;
    // A selection that is a single cell or starts at the header means "all the data rows".
    const lastRow = sel.r2 > firstRow ? sel.r2 : dataRows;
    const labelCol = sel.c1;
    const seriesCols = sel.c2 > sel.c1 && lastRow >= firstRow ? Array.from({ length: sel.c2 - sel.c1 }, (_, i) => sel.c1 + 1 + i) : [];
    const range = (c: number) => `${columnLetter(c)}${firstRow + 1}:${columnLetter(c)}${lastRow + 1}`;
    return {
      type: "column",
      title: seriesCols.length ? `${seriesCols.map((c) => sheet.columns[c]?.header).join(" and ")} by ${sheet.columns[labelCol]?.header}` : "",
      categories: seriesCols.length ? range(labelCol) : "",
      series: seriesCols.map((c) => ({ name: sheet.columns[c]?.header ?? columnLetter(c), range: range(c), color: null, kind: null })),
      stacked: false,
      xTitle: "",
      yTitle: "",
    };
  });

  const set = (patch: Partial<Draft>) => setDraft((d) => ({ ...d, ...patch }));
  const setSeries = (i: number, patch: Partial<ChartSeriesSpec>) =>
    setDraft((d) => ({ ...d, series: d.series.map((s, j) => (j === i ? { ...s, ...patch } : s)) }));

  const single = SINGLE_SERIES.includes(draft.type);
  const shownSeries = single ? draft.series.slice(0, 1) : draft.series;
  const rangesOk = RANGE_RE.test(draft.categories.trim()) && shownSeries.every((s) => RANGE_RE.test(s.range.trim()));
  const hasData = draft.series.length > 0;
  const ok = hasData && rangesOk && !!draft.title.trim();

  const save = () => {
    if (!ok) return;
    if (editing && target) {
      const spec = toChart(draft, existing?.anchor);
      useStore.getState().mutate((wb) => {
        const charts = wb.sheets[target.sheet]?.charts;
        if (charts && charts[target.index]) charts[target.index] = spec;
      });
    } else {
      useStore.getState().addChart(toChart(draft, undefined));
    }
    close();
  };

  const xLabel = draft.type === "bar" ? "Category axis title (left)" : draft.type === "scatter" ? "X axis title" : "Category axis title";
  const yLabel = draft.type === "bar" ? "Value axis title (bottom)" : draft.type === "scatter" ? "Y axis title" : "Value axis title";

  return (
    <Modal
      title={editing ? "Edit chart" : "Add a chart"}
      onClose={close}
      width="max-w-2xl"
      footer={
        <>
          <Button onClick={close}>Cancel</Button>
          <Button primary disabled={!ok} onClick={save}>
            {editing ? "Save chart" : "Add chart"}
          </Button>
        </>
      }
    >
      {!hasData && !editing ? (
        <p className="text-sm text-slate-600">
          Select the labels column and at least one column of numbers next to it (for example, drag across A to C), then try again.
          You can also ask the AI, e.g. &ldquo;add a line chart of monthly profit&rdquo;.
        </p>
      ) : (
        <>
          <div className="mb-4 grid grid-cols-2 gap-2 sm:grid-cols-4">
            {TYPES.map((t) => (
              <button
                key={t.value}
                type="button"
                aria-pressed={draft.type === t.value}
                onClick={() => set({ type: t.value })}
                className={`rounded-lg border p-2 text-left text-sm ${draft.type === t.value ? "border-brand bg-green-50" : "border-slate-200 hover:bg-slate-50"}`}
              >
                <span className="block font-medium text-slate-800">{t.label}</span>
                <span className="text-xs text-slate-500">{t.hint}</span>
              </button>
            ))}
          </div>

          <Field label="Title">
            <input className={inputClass} value={draft.title} onChange={(e) => set({ title: e.target.value })} />
          </Field>

          {STACKABLE.includes(draft.type) && (
            <label className="mb-4 flex items-center gap-2 text-sm text-slate-700">
              <input type="checkbox" className="accent-brand" checked={draft.stacked} onChange={(e) => set({ stacked: e.target.checked })} />
              Stacked
              <span className="text-xs text-slate-500">(series sit on top of each other)</span>
            </label>
          )}

          <Field
            label={draft.type === "scatter" ? "X values (or labels)" : "Labels"}
            hint={draft.type === "scatter" ? "Numbers here become the X position of each point." : undefined}
          >
            <input
              className={`${inputClass} font-mono`}
              value={draft.categories}
              placeholder="A2:A13"
              onChange={(e) => set({ categories: e.target.value })}
            />
          </Field>

          <div className="mb-4">
            <div className="mb-1 flex items-center justify-between">
              <span className="text-sm font-medium text-slate-700">{single ? "Values" : "Series"}</span>
              {single && draft.series.length > 1 && <span className="text-xs text-slate-500">Pie and doughnut charts use the first series only.</span>}
            </div>
            <div className="space-y-2">
              {shownSeries.map((s, i) => {
                const color = s.color || SERIES_COLORS[i % SERIES_COLORS.length];
                return (
                  <div key={i} className="flex flex-wrap items-center gap-2 rounded-md border border-slate-200 p-2">
                    {!single && (
                      <span className="flex items-center gap-1">
                        <input
                          type="color"
                          aria-label={`Colour of ${s.name || `series ${i + 1}`}`}
                          title="Series colour"
                          value={color}
                          onChange={(e) => setSeries(i, { color: e.target.value })}
                          className="h-7 w-8 cursor-pointer rounded border border-slate-300 bg-transparent p-0.5"
                        />
                        {s.color && (
                          <button
                            type="button"
                            title="Use the default colour"
                            aria-label="Use the default colour"
                            className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                            onClick={() => setSeries(i, { color: null })}
                          >
                            <RotateCcw className="h-3.5 w-3.5" />
                          </button>
                        )}
                      </span>
                    )}
                    <input
                      aria-label="Series name"
                      className={`${inputClass} min-w-0 flex-1 py-1.5`}
                      value={s.name}
                      onChange={(e) => setSeries(i, { name: e.target.value })}
                    />
                    <input
                      aria-label="Series range"
                      className={`${inputClass} w-32 py-1.5 font-mono ${RANGE_RE.test(s.range.trim()) ? "" : "bg-red-50"}`}
                      value={s.range}
                      placeholder="B2:B13"
                      onChange={(e) => setSeries(i, { range: e.target.value })}
                    />
                    {draft.type === "column" && (
                      <label className="flex items-center gap-1 text-xs whitespace-nowrap text-slate-600" title="Draw this series as a line over the columns (a combo chart)">
                        <input
                          type="checkbox"
                          className="accent-brand"
                          checked={s.kind === "line"}
                          onChange={(e) => setSeries(i, { kind: e.target.checked ? "line" : null })}
                        />
                        Show as line
                      </label>
                    )}
                    {!single && draft.series.length > 1 && (
                      <button
                        type="button"
                        aria-label={`Remove series ${s.name}`}
                        title="Remove series"
                        className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-red-600"
                        onClick={() => setDraft((d) => ({ ...d, series: d.series.filter((_, j) => j !== i) }))}
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
            {!single && (
              <button
                type="button"
                className="mt-2 flex items-center gap-1 rounded px-2 py-1 text-xs font-medium text-brand hover:bg-slate-100"
                onClick={() => setDraft((d) => ({ ...d, series: [...d.series, { name: `Series ${d.series.length + 1}`, range: "", color: null, kind: null }] }))}
              >
                <Plus className="h-3.5 w-3.5" /> Add series
              </button>
            )}
            {!rangesOk && <p className="mt-2 text-xs text-red-600">Use cell ranges like B2:B13 for the labels and every series.</p>}
          </div>

          {!single && (
            <div className="grid gap-x-3 sm:grid-cols-2">
              <Field label={xLabel}>
                <input className={inputClass} value={draft.xTitle} placeholder="Optional" onChange={(e) => set({ xTitle: e.target.value })} />
              </Field>
              <Field label={yLabel}>
                <input className={inputClass} value={draft.yTitle} placeholder="Optional" onChange={(e) => set({ yTitle: e.target.value })} />
              </Field>
            </div>
          )}
        </>
      )}
    </Modal>
  );
}
