"use client";

import type { ChartType } from "@/lib/types";

export const SERIES_COLORS = ["#1F4E78", "#3E8E5A", "#C55A11", "#7B5EA7", "#B08A1E", "#2E86C1"];

export interface ChartSeriesData {
  name: string;
  values: number[];
  /** Hex colour; the default palette is used when missing. */
  color?: string | null;
  /** In a column chart, "line" draws this series as a line (a combo chart). */
  kind?: "column" | "line" | null;
}

export interface ChartData {
  type: ChartType;
  labels: string[];
  /** Scatter: numeric X for each category, when the category cells are all numbers. */
  xValues?: number[] | null;
  series: ChartSeriesData[];
  stacked?: boolean;
  xTitle?: string | null;
  yTitle?: string | null;
  format: (n: number) => string;
}

const W = 420;
const H = 230;
const compact = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 });

// Text and lines use slate utilities so the dark theme's palette remap applies to them.
const TICK = "fill-slate-500";
const LABEL = "fill-slate-700";
const GRID = "stroke-slate-200";
const AXIS = "stroke-slate-400";
const TITLE = "fill-slate-600";

export function seriesColor(s: ChartSeriesData, i: number) {
  return s.color || SERIES_COLORS[i % SERIES_COLORS.length];
}

function niceMax(v: number) {
  if (v <= 0) return 1;
  const mag = 10 ** Math.floor(Math.log10(v));
  const n = v / mag;
  return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * mag;
}

/** Rounded bounds around [lo, hi] split into four steps. */
function niceRange(lo: number, hi: number): [number, number] {
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) return [0, 1];
  if (lo === hi) {
    const pad = Math.abs(lo) || 1;
    lo -= pad;
    hi += pad;
  }
  const step = niceMax((hi - lo) / 4);
  return [Math.floor(lo / step) * step, Math.ceil(hi / step) * step];
}

function clip(s: string, n: number) {
  return s.length > n ? `${s.slice(0, n - 1)}…` : s;
}

/** A small, dependency-free SVG preview of an Excel chart. */
export default function MiniChart({ data }: { data: ChartData }) {
  if (data.type === "pie" || data.type === "doughnut") return <Pie data={data} hole={data.type === "doughnut"} />;
  if (data.type === "scatter") return <Scatter data={data} />;
  return <Cartesian data={data} />;
}

function Legend({ data, x, y, width }: { data: ChartData; x: number; y: number; width: number }) {
  const step = Math.min(110, width / Math.max(1, data.series.length));
  const asLine = (s: ChartSeriesData) => data.type === "line" || (data.type === "column" && s.kind === "line");
  return (
    <g transform={`translate(${x}, ${y})`}>
      {data.series.map((s, i) => (
        <g key={i} transform={`translate(${i * step}, 0)`}>
          {asLine(s) ? (
            <line x1="0" x2="10" y1="-4" y2="-4" stroke={seriesColor(s, i)} strokeWidth="2.5" />
          ) : (
            <rect width="9" height="9" y="-8" fill={seriesColor(s, i)} rx="1" />
          )}
          <text x="13" fontSize="10" className={LABEL}>
            {clip(s.name, Math.max(4, Math.floor((step - 16) / 5.5)))}
          </text>
        </g>
      ))}
    </g>
  );
}

function AxisTitles({ left, bottom, plot }: { left?: string | null; bottom?: string | null; plot: { l: number; t: number; w: number; h: number; bottomY: number } }) {
  return (
    <>
      {left && (
        <text
          transform={`translate(11, ${plot.t + plot.h / 2}) rotate(-90)`}
          fontSize="10"
          fontWeight="500"
          textAnchor="middle"
          className={TITLE}
        >
          {clip(left, 34)}
        </text>
      )}
      {bottom && (
        <text x={plot.l + plot.w / 2} y={plot.bottomY} fontSize="10" fontWeight="500" textAnchor="middle" className={TITLE}>
          {clip(bottom, 60)}
        </text>
      )}
    </>
  );
}

/** Column, bar, line and area charts, including stacked and column + line combos. */
function Cartesian({ data }: { data: ChartData }) {
  const { type, series } = data;
  const horizontal = type === "bar";
  const stacked = !!data.stacked && (type === "column" || type === "bar" || type === "area");
  const idx = series.map((s, i) => ({ s, i }));
  const bars = type === "bar" ? idx : type === "column" ? idx.filter(({ s }) => s.kind !== "line") : [];
  const lines = type === "line" ? idx : type === "column" ? idx.filter(({ s }) => s.kind === "line") : [];
  const areas = type === "area" ? idx : [];
  const n = Math.max(1, data.labels.length, ...series.map((s) => s.values.length));
  const val = (s: ChartSeriesData, k: number) => s.values[k] ?? 0;

  // Stack offsets: bars stack positives and negatives apart, areas stack cumulatively.
  const barBase: { pos: number; neg: number }[][] = [];
  const areaBase: number[][] = [];
  const extents: number[] = [0];
  if (stacked && bars.length) {
    for (let k = 0; k < n; k++) {
      let pos = 0;
      let neg = 0;
      bars.forEach(({ s }, j) => {
        (barBase[j] ??= [])[k] = { pos, neg };
        const v = val(s, k);
        if (v >= 0) pos += v;
        else neg += v;
      });
      extents.push(pos, neg);
    }
  } else bars.forEach(({ s }) => extents.push(...s.values));
  if (stacked && areas.length) {
    for (let k = 0; k < n; k++) {
      let acc = 0;
      areas.forEach(({ s }, j) => {
        (areaBase[j] ??= [])[k] = acc;
        acc += val(s, k);
        extents.push(acc);
      });
    }
  } else areas.forEach(({ s }) => extents.push(...s.values));
  lines.forEach(({ s }) => extents.push(...s.values));

  const hi = Math.max(...extents);
  const lo = Math.min(...extents);
  const maxV = niceMax(Math.max(0, hi));
  const minV = lo < 0 ? -niceMax(-lo) : 0;

  const legend = series.length > 1;
  const leftTitle = horizontal ? data.xTitle : data.yTitle;
  const bottomTitle = horizontal ? data.yTitle : data.xTitle;
  const PAD = {
    l: (horizontal ? 64 : 52) + (leftTitle ? 14 : 0),
    r: 12,
    t: 12,
    b: 20 + (bottomTitle ? 14 : 0) + (legend ? 16 : 0),
  };
  const plotW = W - PAD.l - PAD.r;
  const plotH = H - PAD.t - PAD.b;
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((f) => minV + (maxV - minV) * f);
  const scale = (v: number) => ((v - minV) / (maxV - minV || 1)) * (horizontal ? plotW : plotH);
  const band = (horizontal ? plotH : plotW) / n;
  const labelEvery = Math.ceil(n / (horizontal ? 12 : 8));
  const vy = (v: number) => PAD.t + plotH - scale(v);
  const cx = (k: number) => PAD.l + band * (k + 0.5);

  /** A bar for category k spanning values a..b, `offset`/`thick` measured along the category axis. */
  const bar = (key: string, k: number, offset: number, thick: number, a: number, b: number, fill: string, tip: string) => {
    const p0 = scale(Math.min(a, b));
    const p1 = scale(Math.max(a, b));
    return horizontal ? (
      <rect key={key} x={PAD.l + p0} y={PAD.t + band * k + offset} width={Math.max(0, p1 - p0)} height={thick} fill={fill} rx="1">
        <title>{tip}</title>
      </rect>
    ) : (
      <rect key={key} x={PAD.l + band * k + offset} y={PAD.t + plotH - p1} width={thick} height={Math.max(0, p1 - p0)} fill={fill} rx="1">
        <title>{tip}</title>
      </rect>
    );
  };
  const tip = (k: number, s: ChartSeriesData, v: number) => `${data.labels[k] ?? k + 1} · ${s.name}: ${data.format(v)}`;

  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="h-auto w-full" role="img">
      {ticks.map((t, i) => {
        const p = scale(t);
        return horizontal ? (
          <g key={i}>
            <line x1={PAD.l + p} x2={PAD.l + p} y1={PAD.t} y2={PAD.t + plotH} className={GRID} />
            <text x={PAD.l + p} y={PAD.t + plotH + 14} fontSize="10" className={TICK} textAnchor="middle">
              {data.format(t)}
            </text>
          </g>
        ) : (
          <g key={i}>
            <line x1={PAD.l} x2={W - PAD.r} y1={PAD.t + plotH - p} y2={PAD.t + plotH - p} className={GRID} />
            <text x={PAD.l - 6} y={PAD.t + plotH - p + 3} fontSize="10" className={TICK} textAnchor="end">
              {data.format(t)}
            </text>
          </g>
        );
      })}

      {areas.map(({ s, i }, j) => {
        const base = (k: number) => (stacked ? (areaBase[j]?.[k] ?? 0) : 0);
        const top = Array.from({ length: n }, (_, k) => `${cx(k)},${vy(base(k) + val(s, k))}`);
        const bottom = Array.from({ length: n }, (_, k) => `${cx(n - 1 - k)},${vy(base(n - 1 - k))}`);
        const color = seriesColor(s, i);
        return (
          <g key={`a${i}`}>
            <polygon points={[...top, ...bottom].join(" ")} fill={color} fillOpacity={stacked ? 0.75 : 0.3} />
            <polyline points={top.join(" ")} fill="none" stroke={color} strokeWidth="1.5" />
            {n <= 24 &&
              s.values.map((v, k) => (
                <circle key={k} cx={cx(k)} cy={vy(base(k) + v)} r="6" fill="transparent">
                  <title>{tip(k, s, v)}</title>
                </circle>
              ))}
          </g>
        );
      })}

      {stacked
        ? bars.map(({ s, i }, j) =>
            Array.from({ length: n }, (_, k) => {
              const v = val(s, k);
              const b = barBase[j]?.[k] ?? { pos: 0, neg: 0 };
              const start = v >= 0 ? b.pos : b.neg;
              return bar(`${i}-${k}`, k, band * 0.2, band * 0.6, start, start + v, seriesColor(s, i), tip(k, s, v));
            }),
          )
        : bars.map(({ s, i }, j) => {
            const inner = (band * 0.75) / bars.length;
            return s.values.map((v, k) => bar(`${i}-${k}`, k, band * 0.125 + inner * j, inner, 0, v, seriesColor(s, i), tip(k, s, v)));
          })}

      {lines.map(({ s, i }) => {
        const color = seriesColor(s, i);
        const pts = s.values.map((v, k) => `${cx(k)},${vy(v)}`).join(" ");
        return (
          <g key={`l${i}`}>
            <polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" />
            {n <= 24 &&
              s.values.map((v, k) => (
                <circle key={k} cx={cx(k)} cy={vy(v)} r="2.5" fill={color}>
                  <title>{tip(k, s, v)}</title>
                </circle>
              ))}
          </g>
        );
      })}

      {data.labels.map((l, i) =>
        i % labelEvery !== 0 ? null : horizontal ? (
          <text key={i} x={PAD.l - 6} y={PAD.t + band * (i + 0.5) + 3} fontSize="10" className={LABEL} textAnchor="end">
            {clip(l, 10)}
          </text>
        ) : (
          <text key={i} x={cx(i)} y={PAD.t + plotH + 14} fontSize="10" className={LABEL} textAnchor="middle">
            {clip(l, Math.max(4, Math.floor((band * labelEvery) / 6)))}
          </text>
        ),
      )}
      {horizontal ? (
        <line x1={PAD.l + scale(0)} x2={PAD.l + scale(0)} y1={PAD.t} y2={PAD.t + plotH} className={AXIS} />
      ) : (
        <line x1={PAD.l} x2={W - PAD.r} y1={vy(0)} y2={vy(0)} className={AXIS} />
      )}
      <AxisTitles left={leftTitle} bottom={bottomTitle} plot={{ l: PAD.l, t: PAD.t, w: plotW, h: plotH, bottomY: PAD.t + plotH + 29 }} />
      {legend && <Legend data={data} x={PAD.l} y={H - 6} width={plotW} />}
    </svg>
  );
}

/** X from numeric category cells (or 1, 2, 3, ... when they are labels), Y from each series. */
function Scatter({ data }: { data: ChartData }) {
  const n = Math.max(data.labels.length, ...data.series.map((s) => s.values.length));
  const xs = data.xValues && data.xValues.length ? data.xValues : Array.from({ length: n }, (_, k) => k + 1);
  const ys = data.series.flatMap((s) => s.values);
  const [x0, x1] = niceRange(Math.min(...xs), Math.max(...xs));
  const [y0, y1] = niceRange(Math.min(0, ...ys), Math.max(...ys, 0));
  const legend = data.series.length > 1;
  const PAD = { l: 52 + (data.yTitle ? 14 : 0), r: 14, t: 12, b: 20 + (data.xTitle ? 14 : 0) + (legend ? 16 : 0) };
  const plotW = W - PAD.l - PAD.r;
  const plotH = H - PAD.t - PAD.b;
  const sx = (v: number) => PAD.l + ((v - x0) / (x1 - x0 || 1)) * plotW;
  const sy = (v: number) => PAD.t + plotH - ((v - y0) / (y1 - y0 || 1)) * plotH;
  const fx = data.xValues && data.xValues.length ? (v: number) => compact.format(v) : (v: number) => (Number.isInteger(v) ? String(v) : "");
  const ticks = [0, 0.25, 0.5, 0.75, 1];

  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="h-auto w-full" role="img">
      {ticks.map((f, i) => {
        const yv = y0 + (y1 - y0) * f;
        const xv = x0 + (x1 - x0) * f;
        return (
          <g key={i}>
            <line x1={PAD.l} x2={W - PAD.r} y1={sy(yv)} y2={sy(yv)} className={GRID} />
            <text x={PAD.l - 6} y={sy(yv) + 3} fontSize="10" className={TICK} textAnchor="end">
              {data.format(yv)}
            </text>
            <line x1={sx(xv)} x2={sx(xv)} y1={PAD.t} y2={PAD.t + plotH} className={GRID} />
            <text x={sx(xv)} y={PAD.t + plotH + 14} fontSize="10" className={TICK} textAnchor="middle">
              {fx(xv)}
            </text>
          </g>
        );
      })}
      <line x1={PAD.l} x2={PAD.l} y1={PAD.t} y2={PAD.t + plotH} className={AXIS} />
      <line x1={PAD.l} x2={W - PAD.r} y1={PAD.t + plotH} y2={PAD.t + plotH} className={AXIS} />
      {data.series.map((s, i) =>
        s.values.map((v, k) =>
          xs[k] === undefined ? null : (
            <circle key={`${i}-${k}`} cx={sx(xs[k])} cy={sy(v)} r="3.5" fill={seriesColor(s, i)} fillOpacity="0.85">
              <title>{`${data.labels[k] ?? xs[k]} · ${s.name}: ${data.format(v)}`}</title>
            </circle>
          ),
        ),
      )}
      <AxisTitles left={data.yTitle} bottom={data.xTitle} plot={{ l: PAD.l, t: PAD.t, w: plotW, h: plotH, bottomY: PAD.t + plotH + 29 }} />
      {legend && <Legend data={data} x={PAD.l} y={H - 6} width={plotW} />}
    </svg>
  );
}

function Pie({ data, hole }: { data: ChartData; hole: boolean }) {
  const values = (data.series[0]?.values ?? []).map((v) => Math.max(0, v));
  const sum = values.reduce((a, b) => a + b, 0);
  const total = sum || 1;
  const cx = 110;
  const cy = H / 2;
  const r = 90;
  const ri = hole ? r * 0.56 : 0;
  const pt = (rad: number, a: number) => `${cx + rad * Math.cos(a)} ${cy + rad * Math.sin(a)}`;
  let angle = -Math.PI / 2;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="h-auto w-full" role="img">
      {values.map((v, i) => {
        if (v <= 0) return null;
        const a0 = angle;
        const a1 = (angle += (v / total) * Math.PI * 2);
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const full = v >= total;
        const circle = (rad: number, sweep: 0 | 1) =>
          `M ${cx - rad} ${cy} a ${rad} ${rad} 0 1 ${sweep} ${2 * rad} 0 a ${rad} ${rad} 0 1 ${sweep} ${-2 * rad} 0 Z`;
        const d = full
          ? circle(r, 0) + (hole ? ` ${circle(ri, 1)}` : "")
          : hole
            ? `M ${pt(r, a0)} A ${r} ${r} 0 ${large} 1 ${pt(r, a1)} L ${pt(ri, a1)} A ${ri} ${ri} 0 ${large} 0 ${pt(ri, a0)} Z`
            : `M ${cx} ${cy} L ${pt(r, a0)} A ${r} ${r} 0 ${large} 1 ${pt(r, a1)} Z`;
        return (
          <path key={i} d={d} fill={SERIES_COLORS[i % SERIES_COLORS.length]} fillRule="evenodd" style={{ stroke: "var(--sc-cell)" }} strokeWidth="1.5">
            <title>{`${data.labels[i]}: ${data.format(v)} (${((v / total) * 100).toFixed(1)}%)`}</title>
          </path>
        );
      })}
      {hole && (
        <text x={cx} y={cy + 4} fontSize="13" fontWeight="600" textAnchor="middle" className={LABEL}>
          {data.format(sum)}
        </text>
      )}
      {data.labels.slice(0, 10).map((l, i) => (
        <g key={i} transform={`translate(230, ${24 + i * 19})`}>
          <rect width="10" height="10" y="-9" fill={SERIES_COLORS[i % SERIES_COLORS.length]} rx="2" />
          <text x="16" fontSize="11" className={LABEL}>
            {clip(l, 16)} · {(((values[i] ?? 0) / total) * 100).toFixed(0)}%
          </text>
        </g>
      ))}
    </svg>
  );
}
