"use client";

import { useState } from "react";
import { BarChart3, ChevronDown, ChevronUp, Pencil, Trash2 } from "lucide-react";
import { valueAt } from "@/lib/formula";
import { CURRENCIES, formatValue, numberStyleAt } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { ChartSpec, NumberFormat, SheetSpec, WorkbookSpec } from "@/lib/types";
import { editChart } from "./chartEdit";
import MiniChart, { type ChartData } from "./MiniChart";

function cells(range: string): { r: number; c: number }[] {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(range.trim().toUpperCase());
  if (!m) return [];
  const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const c1 = col(m[1]);
  const c2 = m[3] ? col(m[3]) : c1;
  const r1 = +m[2] - 1;
  const r2 = m[4] ? +m[4] - 1 : r1;
  const out = [];
  for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push({ r, c });
  return out;
}

function chartData(wb: WorkbookSpec, si: number, sheet: SheetSpec, chart: ChartSpec): ChartData {
  const catValues = cells(chart.categories).map(({ r, c }) => ({ v: valueAt(wb, si, r, c), r, c }));
  const labels = catValues.map(({ v, r, c }) => formatValue(v, numberStyleAt(sheet, r, c)).text);
  const numeric = catValues.length > 0 && catValues.every(({ v }) => typeof v === "number");
  const series = chart.series.map((s) => ({
    name: s.name,
    color: s.color ?? null,
    kind: s.kind ?? null,
    values: cells(s.range).map(({ r, c }) => {
      const v = valueAt(wb, si, r, c);
      return typeof v === "number" ? v : 0;
    }),
  }));
  const firstCol = cells(chart.series[0]?.range ?? "")[0]?.c ?? 0;
  const fmt: NumberFormat = sheet.columns[firstCol]?.format ?? "decimal";
  const symbol = (CURRENCIES.find((c) => c.code === sheet.columns[firstCol]?.currency) ?? CURRENCIES[0]).symbol;
  const compact = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 });
  return {
    type: chart.type,
    labels,
    xValues: numeric ? catValues.map(({ v }) => v as number) : null,
    series,
    stacked: !!chart.stacked,
    xTitle: chart.x_title ?? null,
    yTitle: chart.y_title ?? null,
    format: (n) => (fmt === "percent" ? `${(n * 100).toFixed(0)}%` : fmt === "currency" ? `${symbol}${compact.format(n)}` : compact.format(n)),
  };
}

/** Previews of the active sheet's charts. The download contains real, editable Excel charts. */
export default function ChartsPanel() {
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const sheet = workbook.sheets[active];
  const charts = sheet.charts ?? [];
  const [open, setOpen] = useState(true);
  if (!charts.length) return null;

  return (
    <div className="shrink-0 border-t border-slate-200 bg-slate-50">
      <button
        className="flex w-full items-center gap-2 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
        onClick={() => setOpen(!open)}
      >
        <BarChart3 className="h-3.5 w-3.5 text-brand" />
        Charts on this sheet ({charts.length})
        {open ? <ChevronDown className="ml-auto h-3.5 w-3.5" /> : <ChevronUp className="ml-auto h-3.5 w-3.5" />}
      </button>
      {open && (
        <div className="flex gap-3 overflow-x-auto px-3 pb-3">
          {charts.map((chart, i) => (
            <figure key={i} className="group relative w-[380px] shrink-0 rounded-lg border border-slate-200 bg-white p-3">
              <figcaption className="mb-1 pr-12 text-sm font-medium text-slate-700">{chart.title}</figcaption>
              <MiniChart data={chartData(workbook, active, sheet, chart)} />
              <div className="absolute top-2 right-2 flex gap-0.5 opacity-0 group-focus-within:opacity-100 group-hover:opacity-100">
                <button
                  aria-label={`Edit chart ${chart.title}`}
                  title="Edit chart"
                  className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                  onClick={() => editChart(i)}
                >
                  <Pencil className="h-3.5 w-3.5" />
                </button>
                <button
                  aria-label={`Remove chart ${chart.title}`}
                  title="Remove chart"
                  className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-red-600"
                  onClick={() => useStore.getState().removeChart(i)}
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              </div>
            </figure>
          ))}
        </div>
      )}
    </div>
  );
}
