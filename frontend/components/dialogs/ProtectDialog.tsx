"use client";

import { Lock, LockOpen, X } from "lucide-react";
import Modal, { Button } from "@/components/ui/Modal";
import { columnLetter } from "@/lib/format";
import { useLive } from "@/lib/live";
import { normalize, useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

/** Owner-only protection: lock a whole sheet, or just some cells, against editors' changes. */
export default function ProtectDialog() {
  const close = useUi((s) => s.close);
  const sheet = useStore((s) => s.workbook.sheets[s.active]);
  const selection = useStore((s) => s.selection);
  const role = useLive((s) => (s.status === "live" ? s.role : null));
  const owner = !role || role === "owner";
  const sel = normalize(selection);
  const first = `${columnLetter(sel.c1)}${sel.r1 + 1}`;
  const selRange = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? first : `${first}:${columnLetter(sel.c2)}${sel.r2 + 1}`;
  const ranges = sheet.locked_ranges ?? [];

  const update = (fn: (s: typeof sheet) => void) =>
    useStore.getState().mutate((wb) => {
      const s = wb.sheets.find((x) => x.name === sheet.name);
      if (s) fn(s);
    });

  return (
    <Modal title="Protect cells" onClose={close} width="max-w-md">
      <p className="mb-4 text-sm text-slate-600">
        Locked cells can only be changed by you, the owner. People you invite can still see them, and edit everything else. Downloads keep the protection (no password).
      </p>
      {!owner && <p className="mb-3 rounded-md bg-amber-50 px-3 py-2 text-sm text-amber-800">Only the workbook&apos;s owner can change protection.</p>}
      <label className="mb-4 flex items-center gap-2 text-sm text-slate-700">
        <input type="checkbox" disabled={!owner} checked={!!sheet.locked} onChange={(e) => update((s) => void (s.locked = e.target.checked))} />
        Lock the whole &ldquo;{sheet.name}&rdquo; sheet
      </label>
      <div className="mb-2 flex items-center justify-between">
        <h3 className="text-sm font-semibold text-slate-700">Locked ranges on this sheet</h3>
        <Button
          disabled={!owner || ranges.includes(selRange)}
          onClick={() => update((s) => void (s.locked_ranges = [...(s.locked_ranges ?? []), selRange]))}
        >
          <span className="flex items-center gap-1.5">
            <Lock className="h-3.5 w-3.5" /> Lock {selRange}
          </span>
        </Button>
      </div>
      {ranges.length ? (
        <ul className="divide-y divide-slate-100 rounded-md border border-slate-200 text-sm" aria-label="Locked ranges">
          {ranges.map((r) => (
            <li key={r} className="flex items-center gap-2 px-3 py-1.5">
              <Lock className="h-3.5 w-3.5 text-slate-400" />
              <span className="flex-1 font-mono text-xs text-slate-700">{r}</span>
              {owner && (
                <button
                  aria-label={`Unlock ${r}`}
                  className="flex items-center gap-1 rounded px-1.5 py-0.5 text-xs text-slate-600 hover:bg-slate-100"
                  onClick={() => update((s) => void (s.locked_ranges = (s.locked_ranges ?? []).filter((x) => x !== r)))}
                >
                  <LockOpen className="h-3.5 w-3.5" /> Unlock
                </button>
              )}
            </li>
          ))}
        </ul>
      ) : (
        <p className="text-sm text-slate-400">None yet. Select cells, then lock them here.</p>
      )}
      <div className="mt-5 flex justify-end">
        <Button onClick={close}>
          <span className="flex items-center gap-1.5">
            <X className="h-3.5 w-3.5" /> Done
          </span>
        </Button>
      </div>
    </Modal>
  );
}
