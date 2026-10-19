"use client";

import { useEffect, useState } from "react";
import { History, Loader2, X } from "lucide-react";
import { useAccount } from "@/lib/account";
import { send } from "@/lib/api";
import { columnLetter } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { CellSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";

interface Change {
  user: { id: string; email: string };
  at: number;
  value: CellSpec | null;
}

interface Activity {
  user: { id: string; email: string };
  at: number;
  sheet: string;
  cell: string | null;
  kind: "edit" | "restructure" | "replace";
  summary: string;
}

const when = (ms: number) => new Date(ms).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
const show = (v: CellSpec | null) => (v?.value === null || v?.value === undefined || v.value === "" ? "(empty)" : String(v.value));

/** Who changed the selected cell, and when; plus recent activity across the workbook. */
export default function CellHistoryPanel() {
  const close = useUi((s) => s.close);
  const user = useAccount((s) => s.user);
  const docId = useStore((s) => s.docId);
  const sheet = useStore((s) => s.workbook.sheets[s.active].name);
  const { r2, c2 } = useStore((s) => s.selection);
  const cell = `${columnLetter(c2)}${r2 + 1}`;
  const [tab, setTab] = useState<"cell" | "all">("cell");
  const [changes, setChanges] = useState<Change[] | null>(null);
  const [activity, setActivity] = useState<Activity[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!user) return;
    let gone = false;
    setError(null);
    const id = encodeURIComponent(docId);
    const req =
      tab === "cell"
        ? send<Change[]>("GET", `/api/docs/${id}/history?sheet=${encodeURIComponent(sheet)}&cell=${cell}`).then((x) => !gone && setChanges(x))
        : send<Activity[]>("GET", `/api/docs/${id}/activity?limit=50`).then((x) => !gone && setActivity(x));
    req.catch((e) => !gone && setError(/not found/i.test(e.message) ? "History starts once this workbook is in your account and edited live." : e.message));
    return () => {
      gone = true;
    };
  }, [user, docId, sheet, cell, tab]);

  return (
    <aside
      role="complementary"
      aria-label="Edit history"
      className="fixed top-16 right-3 bottom-14 z-40 flex w-[min(360px,calc(100vw-24px))] flex-col rounded-xl border border-slate-200 bg-white shadow-2xl"
    >
      <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3">
        <History className="h-5 w-5 text-brand" />
        <h2 className="flex-1 text-sm font-semibold text-slate-800">Who changed what</h2>
        <button aria-label="Close" className="rounded p-1 text-slate-500 hover:bg-slate-100" onClick={close}>
          <X className="h-4 w-4" />
        </button>
      </div>
      <div role="tablist" aria-label="History scope" className="flex gap-1 border-b border-slate-100 px-3 py-2 text-xs">
        {(
          [
            ["cell", `${sheet}!${cell}`],
            ["all", "Whole workbook"],
          ] as const
        ).map(([k, label]) => (
          <button
            key={k}
            role="tab"
            aria-selected={tab === k}
            className={`truncate rounded-full px-2.5 py-1 ${tab === k ? "bg-brand text-white" : "text-slate-600 hover:bg-slate-100"}`}
            onClick={() => setTab(k)}
          >
            {label}
          </button>
        ))}
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-4 py-3 text-sm">
        {!user && <p className="text-slate-600">Sign in and share a workbook to see who changed each cell.</p>}
        {error && <p className="text-slate-500">{error}</p>}
        {user && !error && (tab === "cell" ? changes : activity) === null && (
          <p className="flex items-center gap-2 text-slate-500">
            <Loader2 className="h-4 w-4 animate-spin" /> Loading...
          </p>
        )}
        {tab === "cell" && changes && !error && (
          <ol className="space-y-2.5">
            {!changes.length && <li className="text-slate-400">No recorded changes to this cell yet.</li>}
            {changes.map((ch, i) => (
              <li key={i} className="border-l-2 border-slate-200 pl-3">
                <div className="text-xs text-slate-500">
                  <span className="font-medium text-slate-700">{ch.user.email}</span> · {when(ch.at)}
                </div>
                <div className="truncate font-mono text-xs text-slate-800">{show(ch.value)}</div>
              </li>
            ))}
          </ol>
        )}
        {tab === "all" && activity && !error && (
          <ol className="space-y-2.5">
            {!activity.length && <li className="text-slate-400">No recorded changes yet.</li>}
            {activity.map((a, i) => (
              <li key={i} className="border-l-2 border-slate-200 pl-3">
                <div className="text-xs text-slate-500">
                  <span className="font-medium text-slate-700">{a.user.email}</span> · {when(a.at)}
                </div>
                <div className="text-xs text-slate-800">
                  {a.cell ? `${a.sheet}!${a.cell}: ` : ""}
                  {a.summary}
                </div>
              </li>
            ))}
          </ol>
        )}
      </div>
    </aside>
  );
}
