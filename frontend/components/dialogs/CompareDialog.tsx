"use client";

import { useEffect, useMemo, useState } from "react";
import { create } from "zustand";
import { ArrowRight, RotateCcw } from "lucide-react";
import Modal, { Button, inputClass } from "@/components/ui/Modal";
import { diffWorkbooks } from "@/lib/diff";
import { evaluateCached } from "@/lib/formula";
import { columnLetter, formatValue, numberStyleAt } from "@/lib/format";
import { useStore } from "@/lib/store";
import type { Checkpoint, WorkbookSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";

/** "current" or a checkpoint id. */
export type VersionId = string;
export const CURRENT = "current";

/**
 * Which versions the compare dialog opens with. Set it, then open("compare"):
 * `useComparePick.getState().pick(checkpointId)` compares that version with the current one.
 */
export const useComparePick = create<{ from: VersionId | null; to: VersionId | null; pick: (from: VersionId, to?: VersionId) => void; clear: () => void }>(
  (set) => ({
    from: null,
    to: null,
    pick: (from, to = CURRENT) => set({ from, to }),
    clear: () => set({ from: null, to: null }),
  }),
);

/** Open the compare dialog with `id` (a checkpoint) against the current workbook. */
export function compareWithCurrent(id: VersionId) {
  useComparePick.getState().pick(id, CURRENT);
  useUi.getState().open("compare");
}

interface Version {
  id: VersionId;
  label: string;
  at: number;
  workbook: WorkbookSpec;
}

interface Side {
  raw: string;
  shown: string;
  formula: boolean;
}

interface Change {
  sheet: string;
  address: string;
  kind: "added" | "removed" | "changed" | "format";
  before: Side;
  after: Side;
}

interface SheetSummary {
  name: string;
  status: "added" | "removed" | "changed";
  added: number;
  removed: number;
  changed: number;
  formatting: number;
  other: boolean;
}

const MAX_ROWS = 500;
const time = (at: number) => new Date(at).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });

function sideAt(wb: WorkbookSpec, si: number, r: number, c: number): Side {
  const sheet = wb.sheets[si];
  if (si < 0 || !sheet) return { raw: "", shown: "", formula: false };
  if (r === 0) {
    const h = sheet.columns[c]?.header ?? "";
    return { raw: h, shown: h, formula: false };
  }
  const v = sheet.rows[r - 1]?.cells[c]?.value;
  const formula = typeof v === "string" && v.startsWith("=");
  const value = formula ? evaluateCached(wb).get(`${si}:${r + 1}:${c + 1}`) : v;
  return { raw: v === null || v === undefined ? "" : String(v), shown: formatValue(value ?? null, numberStyleAt(sheet, r, c)).text, formula };
}

function compare(before: WorkbookSpec, after: WorkbookSpec): { sheets: SheetSummary[]; changes: Change[] } {
  const idx = (wb: WorkbookSpec, name: string) => wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
  const sheets: SheetSummary[] = [];
  const changes: Change[] = [];
  for (const d of diffWorkbooks(before, after)) {
    const bi = idx(before, d.name);
    const ai = idx(after, d.name);
    const summary: SheetSummary = { name: d.name, status: d.status, added: 0, removed: 0, changed: 0, formatting: 0, other: false };
    let keys = [...d.cells];
    if (d.status === "removed") {
      // A removed sheet: every cell that had something in it is gone.
      const s = before.sheets[bi];
      keys = [];
      s.columns.forEach((_, c) => keys.push(`0:${c}`));
      s.rows.forEach((row, ri) => row.cells.forEach((cell, c) => cell.value !== null && cell.value !== undefined && cell.value !== "" && keys.push(`${ri + 1}:${c}`)));
    }
    keys
      .map((k) => k.split(":").map(Number) as [number, number])
      .sort((a, b) => a[0] - b[0] || a[1] - b[1])
      .forEach(([r, c]) => {
        const b = sideAt(before, bi, r, c);
        const a = sideAt(after, ai, r, c);
        const kind: Change["kind"] = !b.raw && a.raw ? "added" : b.raw && !a.raw ? "removed" : b.raw !== a.raw ? "changed" : "format";
        summary[kind === "format" ? "formatting" : kind]++;
        changes.push({ sheet: d.name, address: `${columnLetter(c)}${r + 1}`, kind, before: b, after: a });
      });
    if (d.status === "changed" && !d.cells.size) summary.other = true; // charts, dropdowns or rules only
    sheets.push(summary);
  }
  return { sheets, changes };
}

const KIND_STYLE: Record<Change["kind"], string> = {
  added: "bg-green-50 text-green-700",
  removed: "bg-red-50 text-red-700",
  changed: "bg-amber-50 text-amber-800",
  format: "bg-slate-100 text-slate-600",
};
const KIND_LABEL: Record<Change["kind"], string> = { added: "Added", removed: "Removed", changed: "Changed", format: "Formatting" };

function Cell({ side, strike }: { side: Side; strike?: boolean }) {
  if (!side.raw) return <span className="text-slate-400">(empty)</span>;
  return (
    <div className={strike ? "text-slate-500" : "text-slate-800"}>
      <div className="break-words">{side.shown || <span className="text-slate-400">(blank)</span>}</div>
      {side.formula && <code className="mt-0.5 block font-mono text-[11px] break-all text-slate-500">{side.raw}</code>}
    </div>
  );
}

/** Compare any two saved versions (or one with the current workbook), cell by cell. */
export default function CompareDialog() {
  const close = useUi((s) => s.close);
  const checkpoints = useStore((s) => s.checkpoints);
  const workbook = useStore((s) => s.workbook);

  const versions: Version[] = useMemo(
    () => [
      { id: CURRENT, label: "Current workbook", at: Number.POSITIVE_INFINITY, workbook },
      ...[...checkpoints].reverse().map((cp: Checkpoint) => ({ id: cp.id, label: cp.label, at: cp.at, workbook: cp.workbook })),
    ],
    [checkpoints, workbook],
  );

  const [initial] = useState(() => {
    const pre = useComparePick.getState();
    if (pre.from && versions.some((v) => v.id === pre.from)) return { from: pre.from, to: pre.to && versions.some((v) => v.id === pre.to) ? pre.to : CURRENT };
    // Default: the latest checkpoint vs now; if that checkpoint *is* now, the one before it.
    const latest = checkpoints.at(-1);
    const from = latest && latest.workbook === workbook && checkpoints.length > 1 ? checkpoints.at(-2)! : latest;
    return { from: from?.id ?? CURRENT, to: CURRENT };
  });
  // The preselection is used once; the next open starts from the default again.
  useEffect(() => useComparePick.getState().clear(), []);
  const [fromId, setFromId] = useState<VersionId>(initial.from);
  const [toId, setToId] = useState<VersionId>(initial.to);
  const [sheetFilter, setSheetFilter] = useState("");
  const [showFormatting, setShowFormatting] = useState(false);

  const from = versions.find((v) => v.id === fromId) ?? versions[0];
  const to = versions.find((v) => v.id === toId) ?? versions[0];
  // Always show the older one as "before".
  const [older, newer] = from.at <= to.at ? [from, to] : [to, from];
  const result = useMemo(() => compare(older.workbook, newer.workbook), [older, newer]);
  const visible = result.changes.filter((c) => (!sheetFilter || c.sheet === sheetFilter) && (showFormatting || c.kind !== "format"));
  const totals = result.sheets.reduce((t, s) => ({ added: t.added + s.added, removed: t.removed + s.removed, changed: t.changed + s.changed }), { added: 0, removed: 0, changed: 0 });

  const picker = (label: string, value: VersionId, onChange: (v: VersionId) => void) => (
    <label className="block min-w-0 flex-1">
      <span className="mb-1 block text-xs font-medium text-slate-600">{label}</span>
      <select className={inputClass} value={value} onChange={(e) => onChange(e.target.value)}>
        {versions.map((v) => (
          <option key={v.id} value={v.id}>
            {v.id === CURRENT ? v.label : `${v.label} · ${time(v.at)}`}
          </option>
        ))}
      </select>
    </label>
  );

  const restore = () => {
    if (older.id === CURRENT) return;
    if (!confirm(`Restore "${older.label}"? The current workbook stays in the version history.`)) return;
    useStore.getState().restoreCheckpoint(older.id);
    close();
    useUi.getState().notify(`Restored "${older.label}".`);
  };

  return (
    <Modal
      title="Compare versions"
      onClose={close}
      width="max-w-4xl"
      footer={
        <>
          {older.id !== CURRENT && (
            <span className="mr-auto">
              <Button onClick={restore}>
                <span className="flex items-center gap-1.5">
                  <RotateCcw className="h-3.5 w-3.5" aria-hidden="true" /> Restore this version
                </span>
              </Button>
            </span>
          )}
          <Button primary onClick={close}>
            Done
          </Button>
        </>
      }
    >
      {checkpoints.length === 0 ? (
        <p className="text-sm text-slate-500">
          There are no saved versions yet. Versions are saved each time the AI changes the workbook, or when you save one in the History panel.
        </p>
      ) : (
        <>
          <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-end">
            {picker("Before", fromId, setFromId)}
            <ArrowRight className="mx-auto hidden h-4 w-4 shrink-0 text-slate-400 sm:mb-3 sm:block" aria-hidden="true" />
            {picker("After", toId, setToId)}
          </div>
          {from.id === to.id ? (
            <p className="text-sm text-slate-500">Pick two different versions to compare.</p>
          ) : (
            <>
              <p className="mb-3 text-sm text-slate-600" aria-live="polite">
                From <strong className="font-medium">{older.label}</strong> to <strong className="font-medium">{newer.label}</strong>:{" "}
                {result.sheets.length === 0
                  ? "no differences."
                  : `${totals.changed} cell${totals.changed === 1 ? "" : "s"} changed, ${totals.added} added, ${totals.removed} removed.`}
              </p>

              {result.sheets.length > 0 && (
                <table className="mb-4 w-full text-left text-sm">
                  <caption className="sr-only">Changes per sheet</caption>
                  <thead>
                    <tr className="border-b border-slate-200 text-xs text-slate-500">
                      <th scope="col" className="py-1.5 pr-2 font-medium">Sheet</th>
                      <th scope="col" className="px-2 py-1.5 text-right font-medium">Changed</th>
                      <th scope="col" className="px-2 py-1.5 text-right font-medium">Added</th>
                      <th scope="col" className="px-2 py-1.5 text-right font-medium">Removed</th>
                      <th scope="col" className="py-1.5 pl-2 text-right font-medium">Formatting</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.sheets.map((s) => (
                      <tr key={s.name} className="border-b border-slate-100">
                        <th scope="row" className="py-1.5 pr-2 font-medium text-slate-800">
                          <button
                            type="button"
                            className="text-left hover:text-brand hover:underline"
                            aria-pressed={sheetFilter === s.name}
                            onClick={() => setSheetFilter(sheetFilter === s.name ? "" : s.name)}
                            title="Show only this sheet's changes"
                          >
                            {s.name}
                          </button>
                          {s.status !== "changed" && (
                            <span className={`ml-2 rounded px-1.5 py-0.5 text-[11px] ${s.status === "added" ? KIND_STYLE.added : KIND_STYLE.removed}`}>
                              Sheet {s.status}
                            </span>
                          )}
                          {s.other && <span className="ml-2 text-xs font-normal text-slate-500">charts, dropdowns or rules changed</span>}
                        </th>
                        <td className="px-2 py-1.5 text-right tabular-nums text-slate-700">{s.changed}</td>
                        <td className="px-2 py-1.5 text-right tabular-nums text-green-700">{s.added}</td>
                        <td className="px-2 py-1.5 text-right tabular-nums text-red-600">{s.removed}</td>
                        <td className="py-1.5 pl-2 text-right tabular-nums text-slate-500">{s.formatting}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}

              {result.changes.length > 0 && (
                <>
                  <div className="mb-2 flex flex-wrap items-center gap-3">
                    <label className="flex items-center gap-2 text-sm text-slate-600">
                      Sheet
                      <select className="rounded-md border border-slate-300 bg-white px-2 py-1 text-sm" value={sheetFilter} onChange={(e) => setSheetFilter(e.target.value)}>
                        <option value="">All sheets</option>
                        {result.sheets.map((s) => (
                          <option key={s.name} value={s.name}>
                            {s.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="flex items-center gap-2 text-sm text-slate-600">
                      <input type="checkbox" checked={showFormatting} onChange={(e) => setShowFormatting(e.target.checked)} />
                      Include formatting-only changes
                    </label>
                    <span className="ml-auto text-xs text-slate-500">
                      {visible.length} change{visible.length === 1 ? "" : "s"}
                      {visible.length > MAX_ROWS ? `, first ${MAX_ROWS} shown` : ""}
                    </span>
                  </div>
                  <div className="overflow-x-auto rounded-lg border border-slate-200">
                    <table className="w-full min-w-[560px] text-left text-sm">
                      <caption className="sr-only">Changed cells</caption>
                      <thead className="bg-slate-50 text-xs text-slate-500">
                        <tr>
                          <th scope="col" className="px-3 py-2 font-medium">Cell</th>
                          <th scope="col" className="px-3 py-2 font-medium">Change</th>
                          <th scope="col" className="w-[38%] px-3 py-2 font-medium">Before</th>
                          <th scope="col" className="w-[38%] px-3 py-2 font-medium">After</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visible.slice(0, MAX_ROWS).map((c, i) => (
                          <tr key={`${c.sheet}!${c.address}-${i}`} className="border-t border-slate-100 align-top">
                            <th scope="row" className="px-3 py-2 font-mono text-xs font-medium whitespace-nowrap text-slate-700">
                              {!sheetFilter && <span className="font-sans font-normal text-slate-500">{c.sheet}!</span>}
                              {c.address}
                            </th>
                            <td className="px-3 py-2">
                              <span className={`rounded px-1.5 py-0.5 text-[11px] font-medium ${KIND_STYLE[c.kind]}`}>{KIND_LABEL[c.kind]}</span>
                            </td>
                            <td className="px-3 py-2">
                              <Cell side={c.before} strike />
                            </td>
                            <td className="px-3 py-2">
                              <Cell side={c.after} />
                            </td>
                          </tr>
                        ))}
                        {visible.length === 0 && (
                          <tr>
                            <td colSpan={4} className="px-3 py-4 text-center text-sm text-slate-500">
                              Only formatting changed here. Tick “Include formatting-only changes” to see those cells.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </>
          )}
        </>
      )}
    </Modal>
  );
}
