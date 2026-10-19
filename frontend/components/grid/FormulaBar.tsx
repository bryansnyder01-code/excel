"use client";

import { useEffect, useState } from "react";
import { Loader2, Sparkles, Wand2, X } from "lucide-react";
import { explainCell } from "@/lib/api";
import { valueAt } from "@/lib/formula";
import { columnLetter, formatValue, numberStyleAt, parseInput } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { ExplainResult } from "@/lib/types";
import FormulaInput from "./FormulaInput";

/** Shows and edits the raw value or formula of the active cell, with an AI "explain / fix" helper. */
export default function FormulaBar() {
  const { r2: r, c2: c } = useStore((s) => s.selection);
  const raw = useStore((s) => {
    const sheet = s.workbook.sheets[s.active];
    if (r === 0) return sheet.columns[c]?.header ?? "";
    const v = sheet.rows[r - 1]?.cells[c]?.value;
    return v === null || v === undefined ? "" : String(v);
  });
  const [value, setValue] = useState(raw);
  useEffect(() => setValue(raw), [raw, r, c]);

  const [explain, setExplain] = useState<{ loading: boolean; result?: ExplainResult; error?: string } | null>(null);
  const [question, setQuestion] = useState("");
  useEffect(() => {
    setExplain(null);
    setQuestion("");
  }, [r, c]);

  const address = `${columnLetter(c)}${r + 1}`;
  const isError = useStore((s) => {
    const v = valueAt(s.workbook, s.active, r, c);
    return !!v && typeof v === "object";
  });

  const commit = () => {
    if (value === raw) return;
    useStore.getState().setCellValue(r, c, r === 0 ? value : parseInput(value), value);
  };

  const ask = async (q?: string) => {
    const { workbook, active } = useStore.getState();
    const sheet = workbook.sheets[active];
    const shown = formatValue(valueAt(workbook, active, r, c), numberStyleAt(sheet, r, c)).text;
    setExplain({ loading: true });
    try {
      const result = await explainCell(workbook, sheet.name, address, shown, q);
      setExplain({ loading: false, result });
    } catch (e) {
      setExplain({ loading: false, error: e instanceof Error ? e.message : "Something went wrong." });
    }
  };

  const applyFix = (fix: string) => {
    const st = useStore.getState();
    st.setCellValue(r, c, r === 0 ? fix : parseInput(fix));
    st.saveCheckpoint(`Fixed ${address}`);
    setExplain(null);
  };

  return (
    <div className="relative shrink-0">
      <div className="flex h-9 items-center border-b border-slate-200 text-sm">
        <div className="w-16 border-r border-slate-200 px-2 text-center font-mono text-xs text-slate-600">{address}</div>
        <span className="px-2 font-serif text-sm italic text-slate-400">fx</span>
        <div className="relative h-full min-w-0 flex-1">
          <FormulaInput
            ariaLabel="Cell contents"
            popup="below-wide"
            className="h-full w-full bg-transparent px-1 font-mono text-[13px] focus:outline-none"
            value={value}
            onChange={setValue}
            onBlur={commit}
            onKeyDown={(e) => {
              if (e.key === "Enter") (e.target as HTMLInputElement).blur();
              if (e.key === "Escape") {
                setValue(raw);
                (e.target as HTMLInputElement).blur();
              }
            }}
          />
        </div>
        {raw !== "" && (
          <button
            onClick={() => ask()}
            className={`mr-2 flex items-center gap-1 rounded-md px-2 py-1 text-xs font-medium ${
              isError ? "bg-red-50 text-red-700 hover:bg-red-100" : "text-brand hover:bg-green-50"
            }`}
          >
            <Sparkles className="h-3.5 w-3.5" />
            {isError ? "Fix this error" : "Explain"}
          </button>
        )}
      </div>

      {explain && (
        <div className="absolute top-full right-2 z-40 mt-1 w-[440px] rounded-lg border border-slate-200 bg-white p-4 text-sm shadow-lg">
          <button aria-label="Close" className="absolute top-2 right-2 rounded p-1 text-slate-400 hover:bg-slate-100" onClick={() => setExplain(null)}>
            <X className="h-3.5 w-3.5" />
          </button>
          <div className="mb-2 flex items-center gap-1.5 text-xs font-semibold tracking-wide text-brand uppercase">
            <Sparkles className="h-3.5 w-3.5" /> {address}
          </div>
          {explain.loading && (
            <p className="flex items-center gap-2 text-slate-500">
              <Loader2 className="h-4 w-4 animate-spin" /> Looking at this cell...
            </p>
          )}
          {explain.error && <p className="text-red-600">{explain.error}</p>}
          {explain.result && (
            <>
              <p className="leading-relaxed whitespace-pre-line text-slate-700">{explain.result.explanation}</p>
              {explain.result.fixed_value && (
                <div className="mt-3 rounded-md bg-green-50 p-2.5">
                  <div className="mb-1 text-xs text-slate-500">Suggested fix</div>
                  <code className="block font-mono text-[13px] break-all text-slate-800">{explain.result.fixed_value}</code>
                  <button
                    onClick={() => applyFix(explain.result!.fixed_value!)}
                    className="mt-2 flex items-center gap-1.5 rounded-md bg-brand px-3 py-1.5 text-xs font-semibold text-white hover:bg-brand-dark"
                  >
                    <Wand2 className="h-3.5 w-3.5" /> Apply fix
                  </button>
                </div>
              )}
            </>
          )}
          {!explain.loading && (
            <form
              className="mt-3 flex gap-2"
              onSubmit={(e) => {
                e.preventDefault();
                if (question.trim()) ask(question.trim());
              }}
            >
              <input
                aria-label="Ask about this cell"
                className="flex-1 rounded-md border border-slate-300 px-2 py-1.5 text-xs focus:border-brand focus:outline-none"
                placeholder="Ask something else about this cell"
                value={question}
                onChange={(e) => setQuestion(e.target.value)}
              />
              <button type="submit" disabled={!question.trim()} className="rounded-md border border-slate-300 px-2.5 text-xs disabled:opacity-40">
                Ask
              </button>
            </form>
          )}
        </div>
      )}
    </div>
  );
}
