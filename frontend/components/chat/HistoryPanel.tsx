"use client";

import { useState } from "react";
import { Bookmark, GitCompare, RotateCcw } from "lucide-react";
import { compareWithCurrent } from "@/components/dialogs/CompareDialog";
import { CHECKPOINT_LIMIT, useStore } from "@/lib/store";

/** Saved versions of this workbook: every AI change, import, template and summary, plus manual saves. */
export default function HistoryPanel() {
  const checkpoints = useStore((s) => s.checkpoints);
  const workbook = useStore((s) => s.workbook);
  const [label, setLabel] = useState("");
  const current = checkpoints.at(-1)?.workbook === workbook ? checkpoints.at(-1)?.id : undefined;

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <form
        className="flex gap-2 border-b border-slate-200 p-3"
        onSubmit={(e) => {
          e.preventDefault();
          useStore.getState().saveCheckpoint(label.trim() || "Saved version");
          setLabel("");
        }}
      >
        <input
          aria-label="Version name"
          className="flex-1 rounded-md border border-slate-300 px-2.5 py-1.5 text-sm focus:border-brand focus:outline-none"
          placeholder="Name this version (optional)"
          value={label}
          onChange={(e) => setLabel(e.target.value)}
        />
        <button type="submit" className="flex items-center gap-1 rounded-md bg-brand px-3 text-sm font-medium text-white hover:bg-brand-dark">
          <Bookmark className="h-3.5 w-3.5" /> Save
        </button>
      </form>
      <div className="flex-1 overflow-y-auto p-3">
        {checkpoints.length === 0 && (
          <p className="text-sm text-slate-500">
            Versions appear here each time the AI changes the workbook, or when you save one. You can jump back to any of them.
          </p>
        )}
        <ol className="space-y-2">
          {[...checkpoints].reverse().map((cp) => (
            <li key={cp.id} className="rounded-lg border border-slate-200 p-2.5">
              <div className="flex items-start gap-2">
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium text-slate-800" title={cp.label}>
                    {cp.label}
                  </div>
                  <div className="text-xs text-slate-500">
                    {new Date(cp.at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })} ·{" "}
                    {cp.workbook.sheets.map((s) => s.name).join(", ")}
                  </div>
                </div>
                {cp.id === current ? (
                  <span className="rounded bg-green-50 px-2 py-0.5 text-xs text-brand">Current</span>
                ) : (
                  <div className="flex shrink-0 items-center">
                    <button
                      className="flex items-center gap-1 rounded px-2 py-1 text-xs text-slate-600 hover:bg-slate-100 hover:text-brand"
                      aria-label={`Compare "${cp.label}" with the current workbook`}
                      title="Compare with the current workbook"
                      onClick={() => compareWithCurrent(cp.id)}
                    >
                      <GitCompare className="h-3 w-3" aria-hidden="true" /> Compare
                    </button>
                    <button
                      className="flex items-center gap-1 rounded px-2 py-1 text-xs text-slate-600 hover:bg-slate-100 hover:text-brand"
                      aria-label={`Restore "${cp.label}"`}
                      onClick={() => useStore.getState().restoreCheckpoint(cp.id)}
                    >
                      <RotateCcw className="h-3 w-3" aria-hidden="true" /> Restore
                    </button>
                  </div>
                )}
              </div>
            </li>
          ))}
        </ol>
        {checkpoints.length >= CHECKPOINT_LIMIT && (
          <p className="mt-3 text-xs text-slate-400">Only the latest {CHECKPOINT_LIMIT} versions are kept.</p>
        )}
      </div>
    </div>
  );
}
