"use client";

import Modal from "@/components/ui/Modal";
import { useUi } from "@/lib/ui";

const GROUPS: { title: string; keys: [string, string][] }[] = [
  {
    title: "Find anything",
    keys: [["Ctrl + K", "Command palette: search every action, sheet and cell"]],
  },
  {
    title: "Moving and selecting",
    keys: [
      ["Arrow keys", "Move one cell"],
      ["Shift + arrows", "Extend the selection"],
      ["Tab / Shift + Tab", "Move right / left"],
      ["Ctrl + A", "Select all data"],
      ["Click a column letter or row number", "Select the whole column or row"],
    ],
  },
  {
    title: "Editing",
    keys: [
      ["Enter or F2", "Edit the cell"],
      ["Type", "Replace what's in the cell"],
      ["Esc", "Cancel the edit"],
      ["Delete", "Clear the selected cells"],
      ["Tab (while typing a formula)", "Insert the suggested function"],
      ["Alt + Down", "Open a cell's dropdown list"],
      ["Shift + F2", "Add or edit a comment"],
    ],
  },
  {
    title: "Copy and fill",
    keys: [
      ["Ctrl + C / X / V", "Copy, cut, paste (works with Excel and Google Sheets)"],
      ["Drag the small square", "Fill a series or copy a formula"],
      ["Ctrl + D", "Fill down from the top row of the selection"],
      ["Ctrl + R", "Fill right from the first column"],
    ],
  },
  {
    title: "Formatting and more",
    keys: [
      ["Ctrl + B / I / U", "Bold, italic, underline"],
      ["Ctrl + Z / Ctrl + Y", "Undo / redo"],
      ["Ctrl + F", "Find"],
      ["Ctrl + H", "Find and replace"],
      ["Right-click", "Cell menu"],
      ["Ctrl + /", "This list"],
    ],
  },
];

export default function ShortcutsDialog() {
  const close = useUi((s) => s.close);
  return (
    <Modal title="Keyboard shortcuts" onClose={close} width="max-w-2xl">
      <p className="mb-4 text-xs text-slate-500">On a Mac, use Cmd instead of Ctrl.</p>
      <div className="grid gap-6 sm:grid-cols-2">
        {GROUPS.map((g) => (
          <section key={g.title}>
            <h3 className="mb-2 text-xs font-semibold tracking-wide text-slate-500 uppercase">{g.title}</h3>
            <dl className="space-y-1.5 text-sm">
              {g.keys.map(([k, v]) => (
                <div key={k} className="flex gap-3">
                  <dt className="w-40 shrink-0">
                    <kbd className="rounded border border-slate-300 bg-slate-50 px-1.5 py-0.5 font-mono text-[11px] text-slate-700">{k}</kbd>
                  </dt>
                  <dd className="text-slate-600">{v}</dd>
                </div>
              ))}
            </dl>
          </section>
        ))}
      </div>
    </Modal>
  );
}
