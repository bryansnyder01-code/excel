"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { ArrowDownAZ, ArrowUpAZ } from "lucide-react";
import { evaluateCached } from "@/lib/formula";
import { formatValue, numberStyleAt } from "@/lib/format";
import { useStore } from "@/lib/store";

const BLANK = "";

/** The dropdown on a header cell: sort the column, or show only some of its values. */
export default function FilterMenu({ column, onClose }: { column: number; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const sheet = workbook.sheets[active];
  const current = (sheet.filters ?? []).find((f) => f.column === column);

  // Every value the column shows in its data rows, numbers in order, then text, blanks last.
  const values = useMemo(() => {
    const computed = evaluateCached(workbook);
    const seen = new Map<string, number | string>();
    sheet.rows.forEach((row, i) => {
      if ((row.kind ?? "data") !== "data") return;
      const raw = row.cells[column]?.value;
      const v = typeof raw === "string" && raw.startsWith("=") ? computed.get(`${active}:${i + 2}:${column + 1}`) : raw;
      const text = formatValue(v, numberStyleAt(sheet, i + 1, column)).text;
      if (!seen.has(text)) seen.set(text, typeof v === "number" ? v : text);
    });
    return [...seen.entries()]
      .sort(([ta, a], [tb, b]) => {
        if (ta === BLANK || tb === BLANK) return ta === BLANK ? 1 : -1;
        if (typeof a === "number" && typeof b === "number") return a - b;
        if (typeof a === "number" || typeof b === "number") return typeof a === "number" ? -1 : 1;
        return ta.localeCompare(tb, undefined, { numeric: true, sensitivity: "base" });
      })
      .map(([t]) => t);
  }, [workbook, active, sheet, column]);

  const [checked, setChecked] = useState<Set<string>>(() => new Set(current ? current.values : values));
  const [search, setSearch] = useState("");
  const shown = values.filter((v) => !search || v.toLowerCase().includes(search.toLowerCase()));

  useEffect(() => {
    const close = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e.type === "mousedown" && ref.current?.contains(e.target as Node)) return;
      onClose();
    };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", close);
    return () => {
      window.removeEventListener("mousedown", close);
      window.removeEventListener("keydown", close);
    };
  }, [onClose]);

  const st = useStore.getState();
  const apply = () => {
    st.setFilter(column, checked.size === values.length ? null : [...checked]);
    onClose();
  };

  return (
    <div
      ref={ref}
      role="dialog"
      aria-label="Sort and filter"
      className="absolute top-full left-0 z-50 mt-0.5 w-60 rounded-lg border border-slate-200 bg-white p-2 text-left text-[13px] font-normal text-slate-800 shadow-xl"
      onMouseDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => e.stopPropagation()}
    >
      <button
        className="flex w-full items-center gap-2 rounded px-2 py-1.5 hover:bg-slate-100"
        onClick={() => {
          st.sortBy(column, false);
          onClose();
        }}
      >
        <ArrowDownAZ className="h-4 w-4 text-slate-500" /> Sort A to Z (smallest first)
      </button>
      <button
        className="flex w-full items-center gap-2 rounded px-2 py-1.5 hover:bg-slate-100"
        onClick={() => {
          st.sortBy(column, true);
          onClose();
        }}
      >
        <ArrowUpAZ className="h-4 w-4 text-slate-500" /> Sort Z to A (largest first)
      </button>
      <div className="my-1.5 h-px bg-slate-200" />
      <input
        aria-label="Search values"
        placeholder="Search"
        className="mb-1.5 h-7 w-full rounded-md border border-slate-300 bg-white px-2 text-xs focus:border-brand focus:outline-none"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      <div className="max-h-48 overflow-y-auto rounded border border-slate-100 py-1">
        <label className="flex items-center gap-2 px-2 py-0.5 text-xs font-medium">
          <input
            type="checkbox"
            checked={checked.size === values.length}
            onChange={(e) => setChecked(new Set(e.target.checked ? values : []))}
          />
          Select all
        </label>
        {shown.map((v) => (
          <label key={v} className="flex items-center gap-2 px-2 py-0.5 text-xs">
            <input
              type="checkbox"
              checked={checked.has(v)}
              onChange={(e) => {
                const next = new Set(checked);
                if (e.target.checked) next.add(v);
                else next.delete(v);
                setChecked(next);
              }}
            />
            <span className="truncate">{v === BLANK ? <em className="text-slate-400">(Blanks)</em> : v}</span>
          </label>
        ))}
      </div>
      <div className="mt-2 flex items-center justify-between gap-2">
        <button
          className="text-xs text-slate-600 hover:underline disabled:opacity-40"
          disabled={!current}
          onClick={() => {
            st.setFilter(column, null);
            onClose();
          }}
        >
          Clear filter
        </button>
        <button className="rounded-md bg-brand px-3 py-1 text-xs font-semibold text-white hover:bg-brand-dark disabled:opacity-50" disabled={!checked.size} onClick={apply}>
          Apply
        </button>
      </div>
    </div>
  );
}
