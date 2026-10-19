"use client";

import { useState } from "react";
import { useStore } from "@/lib/store";

/** A small comment box anchored to a cell. Saved into the .xlsx as an Excel comment. */
export default function NoteEditor({ r, c, initial, onDone }: { r: number; c: number; initial: string; onDone: () => void }) {
  const [text, setText] = useState(initial);
  const save = () => {
    if (text.trim() !== initial.trim()) useStore.getState().setNote(r, c, text);
    onDone();
  };
  return (
    <div
      className="absolute top-0 left-full z-50 ml-1 w-64 rounded-md border border-amber-300 bg-amber-50 p-2 text-left font-normal whitespace-normal text-slate-800 shadow-lg"
      onMouseDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
    >
      <textarea
        autoFocus
        aria-label="Comment"
        rows={4}
        className="w-full resize-none rounded border border-amber-200 bg-white p-1.5 text-xs focus:border-brand focus:outline-none"
        placeholder="Add a comment for this cell"
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === "Escape") onDone();
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) save();
        }}
      />
      <div className="mt-1.5 flex items-center justify-end gap-1.5 text-xs">
        {initial && (
          <button
            className="mr-auto rounded px-2 py-1 text-red-600 hover:bg-red-50"
            onClick={() => {
              useStore.getState().setNote(r, c, null);
              onDone();
            }}
          >
            Delete
          </button>
        )}
        <button className="rounded px-2 py-1 text-slate-600 hover:bg-amber-100" onClick={onDone}>
          Cancel
        </button>
        <button className="rounded bg-brand px-2.5 py-1 font-medium text-white hover:bg-brand-dark" onClick={save}>
          Save
        </button>
      </div>
    </div>
  );
}
