"use client";

import { useMemo, useState } from "react";
import { diffWorkbooks } from "@/lib/diff";
import { LayoutDashboard, Plus, X } from "lucide-react";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

export default function SheetTabs() {
  const sheets = useStore((s) => s.workbook.sheets);
  const active = useStore((s) => s.active);
  const view = useUi((s) => s.view);
  const hasDashboard = useStore((s) => (s.workbook.dashboard?.tiles.length ?? 0) > 0);
  const review = useStore((s) => s.review);
  const changed = useMemo(() => (review ? new Set(diffWorkbooks(review.before, useStore.getState().workbook).map((d) => d.name)) : null), [review, sheets]);
  const { setActive, addSheet, renameSheet, deleteSheet } = useStore.getState();
  const [renaming, setRenaming] = useState<number | null>(null);
  const [draft, setDraft] = useState("");
  const [error, setError] = useState<string | null>(null);

  const showSheet = (i: number) => {
    if (i !== active) setActive(i);
    if (view !== "sheet") useUi.setState({ view: "sheet" });
  };

  const finishRename = () => {
    if (renaming === null) return;
    const err = renameSheet(renaming, draft);
    setError(err);
    if (!err) setRenaming(null);
  };

  return (
    <div className="flex h-10 shrink-0 items-stretch border-t border-slate-200 bg-slate-50 text-sm">
      <button
        aria-label="Add sheet"
        title="Add sheet"
        className="flex w-10 items-center justify-center text-slate-500 hover:bg-slate-200"
        onClick={addSheet}
      >
        <Plus className="h-4 w-4" />
      </button>
      <div role="tablist" aria-label="Sheets" className="flex min-w-0 items-stretch overflow-x-auto">
        <button
          type="button"
          role="tab"
          aria-selected={view === "dashboard"}
          title={hasDashboard ? "Dashboard" : "Dashboard: put key numbers and charts on one page"}
          className={`flex items-center gap-1.5 border-r border-slate-200 px-3 whitespace-nowrap focus:outline-none focus-visible:ring-2 focus-visible:ring-brand focus-visible:ring-inset ${
            view === "dashboard" ? "border-b-2 border-b-brand bg-white font-semibold text-brand" : "text-slate-600 hover:bg-slate-100"
          }`}
          onClick={() => useUi.setState({ view: "dashboard" })}
        >
          <LayoutDashboard className="h-4 w-4" aria-hidden="true" />
          Dashboard
        </button>
        {sheets.map((s, i) => (
          <div
            key={i}
            role="tab"
            tabIndex={0}
            aria-selected={view === "sheet" && i === active}
            className={`group flex items-center gap-1 border-r border-slate-200 px-3 focus:outline-none focus-visible:ring-2 focus-visible:ring-brand focus-visible:ring-inset ${
              view === "sheet" && i === active
                ? "border-b-2 border-b-brand bg-white font-semibold text-brand"
                : "cursor-pointer text-slate-600 hover:bg-slate-100"
            }`}
            onClick={() => showSheet(i)}
            onKeyDown={(e) => {
              if (e.target !== e.currentTarget) return;
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                showSheet(i);
              } else if (e.key === "F2") {
                e.preventDefault();
                setRenaming(i);
                setDraft(s.name);
                setError(null);
              }
            }}
            onDoubleClick={() => {
              setRenaming(i);
              setDraft(s.name);
              setError(null);
            }}
            title="Double-click (or F2) to rename"
          >
            {renaming === i ? (
              <input
                autoFocus
                aria-label="Sheet name"
                className="w-32 rounded border border-brand px-1 py-0.5 font-normal text-slate-800 focus:outline-none"
                value={draft}
                maxLength={31}
                onChange={(e) => setDraft(e.target.value)}
                onBlur={finishRename}
                onKeyDown={(e) => {
                  if (e.key === "Enter") finishRename();
                  if (e.key === "Escape") {
                    setRenaming(null);
                    setError(null);
                  }
                }}
              />
            ) : (
              <span className="whitespace-nowrap">
                {s.name}
                {changed?.has(s.name) && <span title="Changed by the AI" className="ml-1.5 inline-block h-1.5 w-1.5 rounded-full bg-amber-500 align-middle" />}
              </span>
            )}
            {sheets.length > 1 && renaming !== i && (
              <button
                aria-label={`Delete ${s.name}`}
                title="Delete sheet"
                className="rounded p-0.5 text-slate-400 opacity-0 hover:bg-slate-200 hover:text-red-600 group-hover:opacity-100"
                onClick={(e) => {
                  e.stopPropagation();
                  if (confirm(`Delete the sheet "${s.name}"? Formulas that use it will show #REF!.`)) deleteSheet(i);
                }}
              >
                <X className="h-3.5 w-3.5" />
              </button>
            )}
          </div>
        ))}
      </div>
      {error && <span className="ml-3 self-center text-xs text-red-600">{error}</span>}
    </div>
  );
}
