"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { ChevronDown, ChevronUp, X } from "lucide-react";
import { evaluateCached } from "@/lib/formula";
import { formatValue, numberStyleAt } from "@/lib/format";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

interface Hit {
  sheet: number;
  r: number;
  c: number;
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Find (Ctrl+F) and replace (Ctrl+H) across the sheet or the whole workbook. */
export default function FindBar() {
  const mode = useUi((s) => s.find);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [allSheets, setAllSheets] = useState(false);
  const [index, setIndex] = useState(-1);
  const findRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (mode) findRef.current?.select();
  }, [mode]);

  const pattern = useMemo(() => {
    if (!query) return null;
    const body = escapeRe(query);
    return new RegExp(wholeCell ? `^${body}$` : body, matchCase ? "g" : "gi");
  }, [query, matchCase, wholeCell]);

  // Matches what each cell shows, and its formula too, so both "1,200" and "SUM" are found.
  const hits = useMemo<Hit[]>(() => {
    if (!pattern) return [];
    const computed = evaluateCached(workbook);
    const out: Hit[] = [];
    const test = (s: string) => {
      pattern.lastIndex = 0;
      return pattern.test(s);
    };
    workbook.sheets.forEach((sheet, si) => {
      if (!allSheets && si !== active) return;
      sheet.columns.forEach((col, c) => test(col.header) && out.push({ sheet: si, r: 0, c }));
      sheet.rows.forEach((row, ri) =>
        row.cells.forEach((cell, c) => {
          const raw = cell.value;
          if (raw === null || raw === undefined || raw === "") return;
          const isFormula = typeof raw === "string" && raw.startsWith("=");
          const shown = formatValue(isFormula ? computed.get(`${si}:${ri + 2}:${c + 1}`) : raw, numberStyleAt(sheet, ri + 1, c)).text;
          if (test(shown) || test(String(raw))) out.push({ sheet: si, r: ri + 1, c });
        }),
      );
    });
    return out;
  }, [pattern, workbook, active, allSheets]);

  useEffect(() => setIndex(-1), [query, matchCase, wholeCell, allSheets]);

  if (!mode) return null;

  const go = (step: 1 | -1) => {
    if (!hits.length) return;
    const next = index < 0 ? (step === 1 ? 0 : hits.length - 1) : (index + step + hits.length) % hits.length;
    setIndex(next);
    const h = hits[next];
    const st = useStore.getState();
    if (h.sheet !== st.active) st.setActive(h.sheet);
    st.select({ r1: h.r, c1: h.c, r2: h.r, c2: h.c });
  };

  const replace = (all: boolean) => {
    if (!pattern || !hits.length) return;
    const targets = all ? hits : [hits[Math.max(0, index)]];
    const n = useStore.getState().replaceIn(targets, new RegExp(pattern.source, pattern.flags), replacement.replace(/\$/g, "$$$$"));
    useUi.getState().notify(n ? `Replaced ${n} ${n === 1 ? "cell" : "cells"}.` : "Nothing to replace. Matches inside calculated results can't be replaced.");
    if (!all) go(1);
  };

  const close = () => {
    useUi.setState({ find: null });
    (document.querySelector(".sc-grid") as HTMLElement | null)?.focus();
  };

  const input = "h-7 rounded-md border border-slate-300 bg-white px-2 text-xs focus:border-brand focus:outline-none";
  const toggle = (on: boolean) => `h-7 rounded-md border px-2 text-xs ${on ? "border-brand bg-green-50 text-brand" : "border-slate-300 text-slate-600 hover:bg-slate-100"}`;

  return (
    <div
      role="search"
      className="flex shrink-0 flex-wrap items-center gap-1.5 border-b border-slate-200 bg-white px-2 py-1.5"
      onKeyDown={(e) => {
        if (e.key === "Escape") close();
      }}
    >
      <input
        ref={findRef}
        autoFocus
        aria-label="Find"
        placeholder="Find"
        className={`${input} w-44`}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            go(e.shiftKey ? -1 : 1);
          }
        }}
      />
      <span className="min-w-14 text-xs text-slate-500" aria-live="polite">
        {query ? (hits.length ? `${index < 0 ? "–" : index + 1} of ${hits.length}` : "No matches") : ""}
      </span>
      <button aria-label="Previous match" className="rounded p-1 text-slate-500 hover:bg-slate-100" onClick={() => go(-1)}>
        <ChevronUp className="h-4 w-4" />
      </button>
      <button aria-label="Next match" className="rounded p-1 text-slate-500 hover:bg-slate-100" onClick={() => go(1)}>
        <ChevronDown className="h-4 w-4" />
      </button>
      {mode === "replace" && (
        <>
          <input
            aria-label="Replace with"
            placeholder="Replace with"
            className={`${input} ml-1 w-40`}
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && replace(false)}
          />
          <button className={toggle(false)} disabled={!hits.length} onClick={() => replace(false)}>
            Replace
          </button>
          <button className={toggle(false)} disabled={!hits.length} onClick={() => replace(true)}>
            Replace all
          </button>
        </>
      )}
      <span className="mx-1 h-5 w-px bg-slate-200" />
      <button className={toggle(matchCase)} aria-pressed={matchCase} onClick={() => setMatchCase(!matchCase)} title="Match case">
        Aa
      </button>
      <button className={toggle(wholeCell)} aria-pressed={wholeCell} onClick={() => setWholeCell(!wholeCell)}>
        Whole cell
      </button>
      <button className={toggle(allSheets)} aria-pressed={allSheets} onClick={() => setAllSheets(!allSheets)}>
        All sheets
      </button>
      {mode === "find" && (
        <button className="text-xs text-brand hover:underline" onClick={() => useUi.setState({ find: "replace" })}>
          Replace...
        </button>
      )}
      <button aria-label="Close find" className="ml-auto rounded p-1 text-slate-400 hover:bg-slate-100" onClick={close}>
        <X className="h-4 w-4" />
      </button>
    </div>
  );
}
