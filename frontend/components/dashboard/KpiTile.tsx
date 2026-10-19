"use client";

import type { KpiValue } from "@/lib/dashboard";

/** The body of a KPI tile: the big number and its change against the comparison cell. */
export default function KpiTile({ value, compact }: { value: KpiValue; compact?: boolean }) {
  const { change } = value;
  const up = change !== null && change > 0;
  const down = change !== null && change < 0;
  const pct = change === null ? "" : `${Math.abs(change * 100).toFixed(Math.abs(change) < 0.1 ? 1 : 0)}%`;
  return (
    <div className="flex min-h-0 flex-1 flex-col justify-center">
      <div
        className={`truncate font-semibold tabular-nums ${compact ? "text-2xl" : "text-3xl"} ${
          value.error ? "text-red-600" : value.number !== null && value.number < 0 ? "text-red-700" : "text-slate-900"
        }`}
        title={value.text}
      >
        {value.text}
      </div>
      {change !== null && (
        <div
          className={`mt-1 text-sm font-medium ${up ? "text-green-700" : down ? "text-red-600" : "text-slate-500"}`}
          aria-label={up ? `Up ${pct} on the comparison` : down ? `Down ${pct} on the comparison` : "No change on the comparison"}
        >
          <span aria-hidden="true">{up ? "▲" : down ? "▼" : "•"}</span> {pct}
          <span className="ml-1 font-normal text-slate-500">vs comparison</span>
        </div>
      )}
    </div>
  );
}
