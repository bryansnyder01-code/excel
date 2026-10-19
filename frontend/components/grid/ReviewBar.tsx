"use client";

import { useMemo, useState } from "react";
import { Check, Eye, Undo2 } from "lucide-react";
import { getAutoApply, setAutoApply } from "@/lib/actions";
import { diffWorkbooks } from "@/lib/diff";
import { useStore } from "@/lib/store";

/** Above the grid while an AI change waits: what changed, and Keep or Discard. */
export default function ReviewBar() {
  const review = useStore((s) => s.review);
  const workbook = useStore((s) => s.workbook);
  const [auto, setAuto] = useState(getAutoApply);
  const diff = useMemo(() => (review ? diffWorkbooks(review.before, workbook) : []), [review, workbook]);
  if (!review) return null;

  const cells = diff.reduce((n, d) => n + d.cells.size, 0);
  const st = useStore.getState();
  const go = (name: string) => {
    const i = workbook.sheets.findIndex((s) => s.name === name);
    if (i >= 0) st.setActive(i);
  };

  return (
    <div role="region" aria-label="Review AI changes" className="flex shrink-0 flex-wrap items-center gap-2 border-b border-amber-300 bg-amber-50 px-3 py-2 text-sm">
      <Eye className="h-4 w-4 shrink-0 text-amber-700" />
      <span className="font-medium text-slate-800">
        Review the AI&apos;s change: {cells} {cells === 1 ? "cell" : "cells"}
      </span>
      <span className="flex flex-wrap gap-1">
        {diff.map((d) => (
          <button
            key={d.name}
            disabled={d.status === "removed"}
            onClick={() => go(d.name)}
            className="rounded-full border border-amber-300 bg-white px-2 py-0.5 text-xs text-slate-700 hover:bg-amber-100 disabled:opacity-60"
          >
            {d.name} <span className="text-slate-500">{d.status === "added" ? "new" : d.status === "removed" ? "removed" : d.cells.size}</span>
          </button>
        ))}
      </span>
      <label className="ml-auto flex items-center gap-1.5 text-xs text-slate-600">
        <input
          type="checkbox"
          checked={auto}
          onChange={(e) => {
            setAuto(e.target.checked);
            setAutoApply(e.target.checked);
          }}
        />
        Apply future changes without asking
      </label>
      <button className="flex items-center gap-1 rounded-md border border-slate-300 bg-white px-3 py-1 text-xs text-slate-700 hover:bg-slate-50" onClick={st.discardReview}>
        <Undo2 className="h-3.5 w-3.5" /> Discard
      </button>
      <button className="flex items-center gap-1 rounded-md bg-brand px-3 py-1 text-xs font-semibold text-white hover:bg-brand-dark" onClick={st.acceptReview}>
        <Check className="h-3.5 w-3.5" /> Keep
      </button>
    </div>
  );
}
