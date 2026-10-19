"use client";

import { useMemo, useState } from "react";
import { Calendar, CaseSensitive, Columns3, CopyMinus, Eraser, Hash, PaintBucket } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { useStore } from "@/lib/store";
import { applyCleanup, cleanupScope, columnsLabel, describeCleanup, type CaseMode, type CleanupOp, type FillMode } from "@/lib/cleanup";
import { columnLetter } from "@/lib/format";
import { useUi } from "@/lib/ui";

type Kind = CleanupOp["kind"];

const OPS: { kind: Kind; label: string; icon: typeof Eraser; help: string }[] = [
  { kind: "dedupe", label: "Remove duplicates", icon: CopyMinus, help: "Rows are duplicates when every chosen column matches (ignoring case and extra spaces). The first is kept." },
  { kind: "trim", label: "Trim spaces", icon: Eraser, help: "Removes leading and trailing spaces and collapses repeated spaces to one." },
  { kind: "case", label: "Change case", icon: CaseSensitive, help: "Changes the case of text cells." },
  { kind: "split", label: "Split column", icon: Columns3, help: "Splits text into new columns inserted to the right. Existing columns move right and formulas that point at them follow." },
  { kind: "dates", label: "Fix dates", icon: Calendar, help: "Turns text like 2024-03-05, 03/05/2024, 5 Mar 2024 or March 5 2024 into real dates and formats the column as dates." },
  { kind: "fill", label: "Fill blanks", icon: PaintBucket, help: "Fills empty cells. Values are not carried across total, subtotal or section rows." },
  { kind: "numbers", label: "Text to numbers", icon: Hash, help: 'Converts text that looks like a number ("1,200", "$5", "(300)", "12%") into a real number. Codes with leading zeros like 00123 are left alone.' },
];

const DELIMITERS = [
  { value: ",", label: "Comma" },
  { value: " ", label: "Space" },
  { value: ";", label: "Semicolon" },
  { value: "\t", label: "Tab" },
  { value: "custom", label: "Custom…" },
];

/** Tidy up the selected columns (or the whole sheet): duplicates, spaces, case, split, dates, blanks, numbers. */
export default function CleanupDialog() {
  const close = useUi((s) => s.close);
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const selection = useStore((s) => s.selection);
  const sheet = workbook.sheets[active];
  const scope = useMemo(() => cleanupScope(sheet, selection), [sheet, selection]);
  const dataCount = sheet.rows.filter((r) => (r.kind ?? "data") === "data").length;

  const [kind, setKind] = useState<Kind>("dedupe");
  const [caseMode, setCaseMode] = useState<CaseMode>("title");
  const [splitCol, setSplitCol] = useState(() => scope.cols[0] ?? 0);
  const [delim, setDelim] = useState(",");
  const [customDelim, setCustomDelim] = useState("");
  const [dayFirst, setDayFirst] = useState(false);
  const [fillMode, setFillMode] = useState<FillMode>("above");
  const [fillValue, setFillValue] = useState("");
  const [done, setDone] = useState<string | null>(null);

  const op = useMemo((): CleanupOp => {
    switch (kind) {
      case "case":
        return { kind, mode: caseMode };
      case "split":
        return { kind, column: scope.cols.includes(splitCol) ? splitCol : scope.cols[0] ?? 0, delimiter: delim === "custom" ? customDelim : delim };
      case "dates":
        return { kind, dayFirst };
      case "fill":
        return { kind, mode: fillMode, value: fillValue };
      case "dedupe":
      case "trim":
      case "numbers":
        return { kind };
    }
  }, [kind, caseMode, splitCol, delim, customDelim, dayFirst, fillMode, fillValue, scope.cols]);

  const invalid =
    (op.kind === "split" && !op.delimiter) || (op.kind === "fill" && op.mode === "value" && fillValue.trim() === "") || !scope.cols.length;

  // Preview: run the operation on a copy.
  const preview = useMemo(() => {
    if (invalid) return null;
    return applyCleanup(structuredClone(workbook), active, op, scope.cols);
  }, [workbook, active, op, scope.cols, invalid]);

  const apply = () => {
    if (!preview?.changed) return;
    let result = preview;
    useStore.getState().mutate((wb) => {
      result = applyCleanup(wb, active, op, scope.cols);
    });
    const msg = describeCleanup(op, result, true);
    setDone(msg);
    useUi.getState().notify(`${msg} Undo with Ctrl+Z.`);
  };

  const info = OPS.find((o) => o.kind === kind)!;
  const colName = (c: number) => `${columnLetter(c)}${sheet.columns[c]?.header ? ` – ${sheet.columns[c].header}` : ""}`;

  return (
    <Modal
      title="Clean up data"
      onClose={close}
      width="max-w-2xl"
      footer={
        <>
          <Button onClick={close}>Close</Button>
          <Button primary disabled={invalid || !preview?.changed} onClick={apply}>
            Apply
          </Button>
        </>
      }
    >
      <p className="mb-4 rounded-md bg-slate-50 px-3 py-2 text-sm text-slate-600">
        {scope.whole ? (
          <>
            Working on the <strong className="text-slate-800">whole sheet</strong> ({sheet.columns.length} columns, {dataCount} data rows). Select
            columns first to narrow it down.
          </>
        ) : (
          <>
            Working on the <strong className="text-slate-800">selected columns {columnsLabel(scope.cols)}</strong> across all {dataCount} data rows.
          </>
        )}{" "}
        Header, total, subtotal and section rows and formula cells are never changed.
      </p>

      <div className="grid gap-4 sm:grid-cols-[11rem_1fr]">
        <div role="radiogroup" aria-label="Clean-up operation" className="flex flex-row flex-wrap gap-1 sm:flex-col">
          {OPS.map((o) => {
            const Icon = o.icon;
            const on = o.kind === kind;
            return (
              <button
                key={o.kind}
                type="button"
                role="radio"
                aria-checked={on}
                onClick={() => {
                  setKind(o.kind);
                  setDone(null);
                }}
                className={`flex items-center gap-2 rounded-md px-3 py-2 text-left text-sm ${
                  on ? "bg-slate-100 font-medium text-slate-900 ring-1 ring-slate-300" : "text-slate-600 hover:bg-slate-50"
                }`}
              >
                <Icon className="h-4 w-4 shrink-0 text-slate-500" />
                {o.label}
              </button>
            );
          })}
        </div>

        <div className="min-w-0">
          <p className="mb-4 text-sm text-slate-600">{info.help}</p>

          {kind === "dedupe" && (
            <p className="mb-4 text-sm text-slate-700">
              Comparing {scope.cols.length === 1 ? "column" : "columns"} <span className="font-mono text-xs">{columnsLabel(scope.cols)}</span>.
            </p>
          )}

          {kind === "case" && (
            <div role="group" aria-label="Change to" className="mb-4">
              <span className="mb-1 block text-sm font-medium text-slate-700">Change to</span>
              <div className="flex gap-2">
                {(
                  [
                    ["upper", "UPPER"],
                    ["lower", "lower"],
                    ["title", "Title Case"],
                  ] as [CaseMode, string][]
                ).map(([m, label]) => (
                  <button
                    key={m}
                    type="button"
                    aria-pressed={caseMode === m}
                    onClick={() => setCaseMode(m)}
                    className={`rounded-md border px-3 py-1.5 text-sm ${caseMode === m ? "border-brand text-brand" : "border-slate-300 text-slate-700 hover:bg-slate-50"}`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>
          )}

          {kind === "split" && (
            <>
              <Field label="Column to split">
                <select aria-label="Column to split" className={inputClass} value={op.kind === "split" ? op.column : 0} onChange={(e) => setSplitCol(+e.target.value)}>
                  {scope.cols.map((c) => (
                    <option key={c} value={c}>
                      {colName(c)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label="Split at">
                <div className="flex gap-2">
                  <select aria-label="Delimiter" className={inputClass} value={delim} onChange={(e) => setDelim(e.target.value)}>
                    {DELIMITERS.map((d) => (
                      <option key={d.label} value={d.value}>
                        {d.label}
                      </option>
                    ))}
                  </select>
                  {delim === "custom" && (
                    <input aria-label="Custom delimiter" className={inputClass} value={customDelim} onChange={(e) => setCustomDelim(e.target.value)} placeholder="e.g. | or -" />
                  )}
                </div>
              </Field>
              <p className="mb-4 text-xs text-slate-500">
                New columns are named after the original (&ldquo;Name 2&rdquo;, &ldquo;Name 3&rdquo;). Formulas, charts, rules and merges that point at columns to the right
                move with them; formulas that point at the split column itself still see only its first part.
              </p>
            </>
          )}

          {kind === "dates" && (
            <Field label="Dates like 03/05/2024 mean" hint="Dates where one part is over 12 (25/12/2024) are read correctly either way.">
              <select aria-label="Date order" className={inputClass} value={dayFirst ? "dmy" : "mdy"} onChange={(e) => setDayFirst(e.target.value === "dmy")}>
                <option value="mdy">Month first: March 5 (US)</option>
                <option value="dmy">Day first: 3 May (UK / Europe)</option>
              </select>
            </Field>
          )}

          {kind === "fill" && (
            <div role="radiogroup" aria-label="Fill empty cells with" className="mb-4">
              <span className="mb-1 block text-sm font-medium text-slate-700">Fill empty cells with</span>
              <div className="space-y-2">
                {(
                  [
                    ["above", "The value above"],
                    ["value", "A fixed value"],
                    ["zero", "0"],
                  ] as [FillMode, string][]
                ).map(([m, label]) => (
                  <label key={m} className="flex items-center gap-2 text-sm text-slate-700">
                    <input type="radio" name="fill-mode" checked={fillMode === m} onChange={() => setFillMode(m)} />
                    {label}
                  </label>
                ))}
                {fillMode === "value" && (
                  <input aria-label="Fill value" className={inputClass} value={fillValue} onChange={(e) => setFillValue(e.target.value)} placeholder="e.g. Unknown" />
                )}
              </div>
            </div>
          )}

          <div aria-live="polite" className="rounded-md border border-slate-200 px-3 py-2 text-sm">
            {done && <p className="mb-1 text-green-700">Done: {done}</p>}
            {invalid || !preview ? (
              <span className="text-slate-500">{scope.cols.length ? "Fill in the options above to see a preview." : "Nothing selected."}</span>
            ) : (
              <span className={preview.changed ? "font-medium text-slate-800" : "text-slate-500"}>{describeCleanup(op, preview)}</span>
            )}
          </div>
        </div>
      </div>
    </Modal>
  );
}
