"use client";

import { useEffect, useState, type ReactNode } from "react";
import {
  BarChart3,
  Camera,
  Cable,
  GitCompare,
  Mail,
  Sparkles,
  CreditCard,
  History,
  Lock,
  MessageSquareText,
  Eraser,
  FlaskConical,
  LayoutTemplate,
  ShieldCheck,
  Wrench,
  Highlighter,
  MessageSquarePlus,
  Pin,
  Search,
  ListChecks,
  Sigma,
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Combine,
  WrapText,
  Columns3,
  Italic,
  PaintBucket,
  Redo2,
  Rows3,
  Square,
  SquareDashed,
  Trash2,
  Type,
  Underline,
  Undo2,
} from "lucide-react";
import { useShallow } from "zustand/react/shallow";
import { CURRENCIES } from "@/lib/format";
import { normalize, useStore } from "@/lib/store";
import type { NumberFormat, RowKind } from "@/lib/types";
import { useUi } from "@/lib/ui";

const FORMATS: { value: NumberFormat; label: string }[] = [
  { value: "text", label: "Text" },
  { value: "integer", label: "Number (1,234)" },
  { value: "decimal", label: "Decimal (1,234.00)" },
  { value: "currency", label: "Currency ($)" },
  { value: "percent", label: "Percent (%)" },
  { value: "date", label: "Date" },
];

const KINDS: { value: RowKind; label: string }[] = [
  { value: "data", label: "Data row" },
  { value: "subtotal", label: "Subtotal row" },
  { value: "total", label: "Total row" },
  { value: "section", label: "Section label" },
];

export default function FormatToolbar() {
  const st = useStore.getState();
  const canUndo = useStore((s) => s.past.length > 0);
  const canRedo = useStore((s) => s.future.length > 0);
  // The cell under the cursor drives the toolbar's pressed states.
  const info = useStore(useShallow((s) => {
    const sheet = s.workbook.sheets[s.active];
    const { r2, c2 } = s.selection;
    const cell = r2 > 0 ? sheet.rows[r2 - 1]?.cells[c2] : undefined;
    return {
      bold: !!cell?.bold,
      italic: !!cell?.italic,
      underline: !!cell?.underline,
      wrap: !!cell?.wrap,
      align: cell?.align ?? null,
      fontColor: cell?.font_color ?? "#1f2937",
      fillColor: cell?.fill_color ?? "#ffff00",
      format: (r2 > 0 ? cell?.format : null) ?? sheet.columns[c2]?.format ?? "text",
      currency: (r2 > 0 ? cell?.currency : null) ?? sheet.columns[c2]?.currency ?? "USD",
      kind: r2 > 0 ? (sheet.rows[r2 - 1]?.kind ?? "data") : null,
      header: normalize(s.selection).r1 === 0,
    };
  }));

  return (
    <div className="flex min-h-11 shrink-0 items-center gap-1 border-b border-slate-200 bg-slate-50 px-2 py-1.5 max-md:overflow-x-auto md:flex-wrap">
      <Btn label="Undo (Ctrl+Z)" onClick={st.undo} disabled={!canUndo}>
        <Undo2 />
      </Btn>
      <Btn label="Redo (Ctrl+Y)" onClick={st.redo} disabled={!canRedo}>
        <Redo2 />
      </Btn>
      <Sep />
      <Btn label="Dropdown list for selected cells" onClick={() => useUi.getState().open("dropdown")}>
        <ListChecks />
        <span className="text-xs">Dropdown</span>
      </Btn>
      <Btn label="Chart the selected cells" onClick={() => useUi.getState().open("chart")}>
        <BarChart3 />
        <span className="text-xs">Chart</span>
      </Btn>
      <Btn label="Summarise by category (pivot-style)" onClick={() => useUi.getState().open("pivot")}>
        <Sigma />
        <span className="text-xs">Summarise</span>
      </Btn>
      <Btn label="Conditional formatting: colour cells by their values" onClick={() => useUi.getState().open("conditional")}>
        <Highlighter />
        <span className="text-xs">Highlight</span>
      </Btn>
      <ToolsMenu />
      <Sep />
      <Btn label="Bold (Ctrl+B)" pressed={info.bold} onClick={() => st.toggleStyle("bold")}>
        <Bold />
      </Btn>
      <Btn label="Italic (Ctrl+I)" pressed={info.italic} onClick={() => st.toggleStyle("italic")}>
        <Italic />
      </Btn>
      <Btn label="Underline (Ctrl+U)" pressed={info.underline} onClick={() => st.toggleStyle("underline")}>
        <Underline />
      </Btn>
      <ColorBtn label="Text colour" icon={<Type />} value={info.fontColor} onChange={(v) => st.styleSelection({ font_color: v })} />
      <ColorBtn label="Fill colour" icon={<PaintBucket />} value={info.fillColor} onChange={(v) => st.styleSelection({ fill_color: v })} />
      <Btn label="Clear colours" onClick={() => st.styleSelection({ font_color: null, fill_color: null })}>
        <span className="text-[11px] font-medium">Reset</span>
      </Btn>
      <Sep />
      <Btn label="Align left" pressed={info.align === "left"} onClick={() => st.styleSelection({ align: "left" })}>
        <AlignLeft />
      </Btn>
      <Btn label="Align centre" pressed={info.align === "center"} onClick={() => st.styleSelection({ align: "center" })}>
        <AlignCenter />
      </Btn>
      <Btn label="Align right" pressed={info.align === "right"} onClick={() => st.styleSelection({ align: "right" })}>
        <AlignRight />
      </Btn>
      <Btn label="Wrap text" pressed={info.wrap} onClick={() => st.toggleStyle("wrap")}>
        <WrapText />
      </Btn>
      <Btn label="Merge or unmerge cells" onClick={st.toggleMerge}>
        <Combine />
      </Btn>
      <Sep />
      <Btn label="Thick border" onClick={() => st.styleSelection({ border: "thick" })}>
        <Square className="stroke-[3]" />
      </Btn>
      <Btn label="Thin border" onClick={() => st.styleSelection({ border: null })}>
        <Square />
      </Btn>
      <Btn label="No border" onClick={() => st.styleSelection({ border: "none" })}>
        <SquareDashed />
      </Btn>
      <Sep />
      <select
        aria-label="Number format"
        title={info.header ? "Number format for the whole column" : "Number format"}
        className="h-8 rounded-md border border-slate-300 bg-white px-2 text-xs"
        value={info.format}
        onChange={(e) => st.setNumberFormat(e.target.value as NumberFormat)}
      >
        {FORMATS.map((f) => (
          <option key={f.value} value={f.value}>
            {f.label}
          </option>
        ))}
      </select>
      <select
        aria-label="Currency"
        title="Currency"
        className="h-8 w-[72px] rounded-md border border-slate-300 bg-white px-1 text-xs"
        value={info.format === "currency" ? info.currency : ""}
        onChange={(e) => e.target.value && st.setCurrency(e.target.value)}
      >
        <option value="">¤ ...</option>
        {CURRENCIES.map((c) => (
          <option key={c.code} value={c.code} title={c.label}>
            {c.symbol.trim()} {c.code}
          </option>
        ))}
      </select>
      <Btn label="Fewer decimal places" onClick={() => st.changeDecimals(-1)}>
        <span className="font-mono text-[11px]">.0←</span>
      </Btn>
      <Btn label="More decimal places" onClick={() => st.changeDecimals(1)}>
        <span className="font-mono text-[11px]">.00→</span>
      </Btn>
      <select
        aria-label="Row type"
        className="h-8 rounded-md border border-slate-300 bg-white px-2 text-xs disabled:opacity-50"
        value={info.kind ?? "data"}
        disabled={info.kind === null}
        onChange={(e) => st.setRowKind(e.target.value as RowKind)}
      >
        {KINDS.map((k) => (
          <option key={k.value} value={k.value}>
            {k.label}
          </option>
        ))}
      </select>
      <Sep />
      <Btn label="Insert row above" onClick={() => st.insertRow(false)}>
        <Rows3 />
        <span className="text-[11px]">↑</span>
      </Btn>
      <Btn label="Insert row below" onClick={() => st.insertRow(true)}>
        <Rows3 />
        <span className="text-[11px]">↓</span>
      </Btn>
      <Btn label="Delete row(s)" onClick={st.deleteRows}>
        <Rows3 />
        <Trash2 className="!h-3 !w-3" />
      </Btn>
      <Btn label="Insert column left" onClick={() => st.insertColumn(false)}>
        <Columns3 />
        <span className="text-[11px]">←</span>
      </Btn>
      <Btn label="Insert column right" onClick={() => st.insertColumn(true)}>
        <Columns3 />
        <span className="text-[11px]">→</span>
      </Btn>
      <Btn label="Delete column(s)" onClick={st.deleteColumns}>
        <Columns3 />
        <Trash2 className="!h-3 !w-3" />
      </Btn>
      <Sep />
      <FreezeMenu />
      <Btn label="Add or edit a comment (Shift+F2)" disabled={info.kind === null} onClick={() => {
        const { r2, c2 } = useStore.getState().selection;
        useUi.setState({ noteEdit: { r: r2, c: c2 } });
      }}>
        <MessageSquarePlus />
      </Btn>
      <Btn label="Find and replace (Ctrl+F)" onClick={() => useUi.setState({ find: "replace" })}>
        <Search />
      </Btn>
    </div>
  );
}

function Btn({
  label,
  onClick,
  pressed,
  disabled,
  children,
}: {
  label: string;
  onClick: () => void;
  pressed?: boolean;
  disabled?: boolean;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      title={label}
      aria-label={label}
      aria-pressed={pressed}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} // keep grid focus and selection
      onClick={onClick}
      className={`flex h-8 min-w-8 shrink-0 items-center justify-center gap-0.5 rounded-md px-1.5 text-slate-700 [&_svg]:h-4 [&_svg]:w-4 hover:bg-slate-200 disabled:opacity-35 disabled:hover:bg-transparent ${
        pressed ? "bg-green-100 text-brand" : ""
      }`}
    >
      {children}
    </button>
  );
}

function ColorBtn({ label, icon, value, onChange }: { label: string; icon: ReactNode; value: string; onChange: (v: string) => void }) {
  return (
    <label
      title={label}
      className="relative flex h-8 w-8 shrink-0 cursor-pointer flex-col items-center justify-center rounded-md text-slate-700 hover:bg-slate-200 [&_svg]:h-4 [&_svg]:w-4"
    >
      {icon}
      <span className="mt-0.5 h-1 w-5 rounded-sm" style={{ background: value }} />
      <input
        type="color"
        aria-label={label}
        className="absolute inset-0 cursor-pointer opacity-0"
        value={/^#[0-9a-f]{6}$/i.test(value) ? value : "#000000"}
        onChange={(e) => onChange(e.target.value)}
      />
    </label>
  );
}

function Sep() {
  return <div className="mx-1 h-6 w-px shrink-0 bg-slate-300" />;
}

/** Keep the header row and/or leading columns in view while scrolling. */
function FreezeMenu() {
  const [open, setOpen] = useState(false);
  const { header, cols } = useStore(useShallow((s) => ({
    header: s.workbook.sheets[s.active].freeze_header !== false,
    cols: s.workbook.sheets[s.active].freeze_columns ?? 0,
  })));
  const selCol = useStore((s) => normalize(s.selection).c1);
  useEffect(() => {
    if (!open) return;
    const close = () => setOpen(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [open]);
  const set = useStore.getState().setFreeze;
  const opt = (label: string, on: boolean, fn: () => void) => (
    <button
      role="menuitemcheckbox"
      aria-checked={on}
      className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-xs hover:bg-slate-100"
      onClick={fn}
    >
      <span className="w-3 text-brand">{on ? "✓" : ""}</span>
      {label}
    </button>
  );
  return (
    <div className="relative">
      <Btn label="Freeze rows and columns" pressed={cols > 0} onClick={() => setTimeout(() => setOpen(!open))}>
        <Pin />
        <span className="text-xs">Freeze</span>
      </Btn>
      {open && (
        <div role="menu" className="absolute top-full right-0 z-50 mt-1 w-56 rounded-lg border border-slate-200 bg-white py-1 shadow-lg md:left-0">
          {opt("Header row", header, () => set(!header, cols))}
          <div className="my-1 h-px bg-slate-200" />
          {opt("No columns", cols === 0, () => set(header, 0))}
          {opt("First column", cols === 1, () => set(header, 1))}
          {opt("First 2 columns", cols === 2, () => set(header, 2))}
          {selCol > 1 && opt(`Up to column ${String.fromCharCode(64 + Math.min(selCol, 26))}`, cols === selCol, () => set(header, selCol))}
        </div>
      )}
    </div>
  );
}

/** Workbook-wide tools: checking, cleaning, what-if, templates. */
function ToolsMenu() {
  const [open, setOpen] = useState(false);
  useEffect(() => {
    if (!open) return;
    const close = () => setOpen(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [open]);
  const ui = useUi.getState();
  const item = (icon: ReactNode, label: string, hint: string, fn: () => void) => (
    <button role="menuitem" className="flex w-full items-start gap-2.5 px-3 py-2 text-left hover:bg-slate-100 [&_svg]:mt-0.5 [&_svg]:h-4 [&_svg]:w-4 [&_svg]:text-brand" onClick={fn}>
      {icon}
      <span>
        <span className="block text-xs font-medium text-slate-800">{label}</span>
        <span className="block text-[11px] text-slate-500">{hint}</span>
      </span>
    </button>
  );
  return (
    <div className="relative">
      <Btn label="Tools: check, clean up, what-if, templates" onClick={() => setTimeout(() => setOpen(!open))}>
        <Wrench />
        <span className="text-xs">Tools</span>
      </Btn>
      {open && (
        <div role="menu" aria-label="Tools" className="absolute top-full left-0 z-50 mt-1 max-h-[70vh] w-72 overflow-y-auto rounded-lg border border-slate-200 bg-white py-1 shadow-lg">
          {item(<Sparkles />, "Fill a column with AI", "Categorise, extract or translate row by row", () => ui.open("aiFill"))}
          {item(<ShieldCheck />, "Check my workbook", "Find broken formulas, typed-in totals and odd cells", () => ui.open("check"))}
          {item(<Eraser />, "Clean up data", "Duplicates, spaces, dates, splitting, blanks", () => ui.open("cleanup"))}
          {item(<FlaskConical />, "What-if and goal seek", "Compare scenarios, or find the input that hits a target", () => ui.open("whatif"))}
          {item(<Camera />, "Import from a photo or PDF", "Receipts, statements, invoices into rows", () => ui.open("importImage"))}
          {item(<Cable />, "Live data", "Pull in a CSV, Google Sheet or API and keep it fresh", () => ui.open("connections"))}
          {item(<Mail />, "Email reports", "Send this workbook on a schedule", () => ui.open("reports"))}
          {item(<GitCompare />, "Compare versions", "See what changed between two versions", () => ui.open("compare"))}
          {item(<LayoutTemplate />, "Save as template", "Reuse this layout and its formulas", () => ui.open("saveTemplate"))}
          <div className="my-1 h-px bg-slate-200" />
          {item(<MessageSquareText />, "Comments", "Discuss cells and @mention people", () => ui.open("comments"))}
          {item(<History />, "Who changed what", "Edit history for a cell or the workbook", () => ui.open("cellHistory"))}
          {item(<Lock />, "Protect cells", "Lock a sheet or ranges so only you can change them", () => ui.open("protect"))}
          {item(<CreditCard />, "Plan and usage", "AI requests left today, upgrade", () => ui.open("billing"))}
        </div>
      )}
    </div>
  );
}
