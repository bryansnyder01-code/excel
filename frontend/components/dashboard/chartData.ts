// Chart preview data for dashboard chart tiles. Same mapping as ChartsPanel's (which keeps its own
// copy private); if ChartsPanel starts exporting chartData, this file can re-export it instead.

import { valueAt } from "@/lib/formula";
import { CURRENCIES, formatValue, numberStyleAt } from "@/lib/format";
import type { ChartSpec, NumberFormat, SheetSpec, WorkbookSpec } from "@/lib/types";
import type { ChartData } from "@/components/charts/MiniChart";

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

export function chartData(wb: WorkbookSpec, si: number, sheet: SheetSpec, chart: ChartSpec): ChartData {
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
