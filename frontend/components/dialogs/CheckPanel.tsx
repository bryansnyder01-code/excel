"use client";

import { useEffect, useState } from "react";
import { AlertTriangle, CheckCircle2, Info, Loader2, ShieldCheck, Sparkles, Wand2, X, XCircle } from "lucide-react";
import { displayedValues } from "@/lib/actions";
import { checkWorkbook } from "@/lib/api";
import { parseRange } from "@/lib/cellStyle";
import { useStore } from "@/lib/store";
import type { Finding } from "@/lib/types";
import { useUi } from "@/lib/ui";

const ICON = {
  error: <XCircle className="h-4 w-4 shrink-0 text-red-600" />,
  warning: <AlertTriangle className="h-4 w-4 shrink-0 text-amber-600" />,
  info: <Info className="h-4 w-4 shrink-0 text-sky-600" />,
};

/** "Check my workbook": a side panel listing likely mistakes, each with a jump to the cell and, where possible, a fix. */
export default function CheckPanel() {
  const close = useUi((s) => s.close);
  const [useAi, setUseAi] = useState(true);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ summary: string; findings: Finding[] } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState<Set<string>>(new Set());

  const run = async (ai = useAi) => {
    setBusy(true);
    setError(null);
    try {
      setResult(await checkWorkbook(useStore.getState().workbook, displayedValues(), ai));
      setDone(new Set());
    } catch (e) {
      setError(e instanceof Error ? e.message : "The check didn't run.");
    } finally {
      setBusy(false);
    }
  };

  useEffect(() => {
    // The free automatic checks run as soon as the panel opens; the AI review is one click away.
    run(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const go = (f: Finding) => {
    const st = useStore.getState();
    const i = st.workbook.sheets.findIndex((s) => s.name === f.sheet);
    const g = parseRange(f.cell);
    if (i < 0 || !g) return false;
    if (i !== st.active) st.setActive(i);
    st.select({ r1: g.r1 - 1, c1: g.c1 - 1, r2: g.r2 - 1, c2: g.c2 - 1 });
    useUi.setState({ highlight: [{ sheet: f.sheet, range: f.cell }] });
    return true;
  };

  const fix = (f: Finding) => {
    if (!f.fix || !go(f)) return;
    const g = parseRange(f.cell)!;
    const value = f.fix;
    useStore.getState().mutate((wb) => {
      const sheet = wb.sheets.find((s) => s.name === f.sheet);
      const row = sheet?.rows[g.r1 - 2];
      if (!row) return;
      while (row.cells.length < g.c1) row.cells.push({ value: null });
      row.cells[g.c1 - 1] = { ...row.cells[g.c1 - 1], value };
    });
    setDone(new Set(done).add(f.id));
  };

  const open = result?.findings.filter((f) => !done.has(f.id)) ?? [];

  return (
    <aside
      role="complementary"
      aria-label="Check my workbook"
      className="fixed top-16 right-3 bottom-14 z-40 flex w-[min(380px,calc(100vw-24px))] flex-col rounded-xl border border-slate-200 bg-white shadow-2xl"
    >
      <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3">
        <ShieldCheck className="h-5 w-5 text-brand" />
        <h2 className="flex-1 text-sm font-semibold text-slate-800">Check my workbook</h2>
        <button aria-label="Close" className="rounded p-1 text-slate-500 hover:bg-slate-100" onClick={close}>
          <X className="h-4 w-4" />
        </button>
      </div>
      <div className="flex items-center gap-2 border-b border-slate-100 px-4 py-2.5">
        <label className="flex flex-1 items-center gap-1.5 text-xs text-slate-600">
          <input type="checkbox" checked={useAi} onChange={(e) => setUseAi(e.target.checked)} />
          <Sparkles className="h-3.5 w-3.5 text-violet-600" /> Include an AI review of the logic
        </label>
        <button
          className="flex items-center gap-1 rounded-md bg-brand px-3 py-1 text-xs font-semibold text-white hover:bg-brand-dark disabled:opacity-50"
          disabled={busy}
          onClick={() => run()}
        >
          {busy && <Loader2 className="h-3.5 w-3.5 animate-spin" />} {result ? "Check again" : "Check"}
        </button>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-4 py-3">
        {busy && !result && (
          <p className="flex items-center gap-2 text-sm text-slate-500">
            <Loader2 className="h-4 w-4 animate-spin" /> Checking formulas, totals and references...
          </p>
        )}
        {error && <p className="text-sm text-red-600" role="alert">{error}</p>}
        {result && (
          <>
            <p className="mb-3 text-sm font-medium text-slate-700">{result.summary}</p>
            {open.length === 0 && (
              <p className="flex items-center gap-2 rounded-lg bg-green-50 px-3 py-2 text-sm text-green-800">
                <CheckCircle2 className="h-4 w-4" /> {result.findings.length ? "Everything found has been fixed." : "Nothing looks wrong."}
              </p>
            )}
            <ul className="space-y-2">
              {open.map((f) => (
                <li key={f.id} className="rounded-lg border border-slate-200 p-2.5">
                  <div className="flex items-start gap-2">
                    {ICON[f.severity] ?? ICON.info}
                    <div className="min-w-0 flex-1">
                      <button className="text-xs font-semibold text-brand hover:underline" onClick={() => go(f)}>
                        {f.sheet}!{f.cell}
                      </button>
                      <p className="mt-0.5 text-sm text-slate-700">{f.message}</p>
                      {f.fix && (
                        <div className="mt-1.5 flex items-center gap-2">
                          <code className="min-w-0 flex-1 truncate rounded bg-slate-100 px-1.5 py-0.5 text-xs text-slate-700">{f.fix}</code>
                          <button
                            className="flex shrink-0 items-center gap-1 rounded-md border border-slate-300 px-2 py-0.5 text-xs text-slate-700 hover:bg-slate-50"
                            onClick={() => fix(f)}
                          >
                            <Wand2 className="h-3 w-3" /> Fix
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </aside>
  );
}
