"use client";

import { useState } from "react";
import { Trash2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { selectionRange, useStore } from "@/lib/store";
import type { CondOperator, CondStyle, ConditionalFormatSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";

const RULES: { value: CondOperator; label: string }[] = [
  { value: "greater_than", label: "Greater than" },
  { value: "less_than", label: "Less than" },
  { value: "between", label: "Between" },
  { value: "equal", label: "Equal to" },
  { value: "negative", label: "Negative numbers" },
  { value: "contains", label: "Text contains" },
  { value: "color_scale", label: "Colour scale (heatmap)" },
  { value: "data_bar", label: "Data bars" },
];

const STYLES: { value: CondStyle; label: string; bg: string; fg: string }[] = [
  { value: "good", label: "Green", bg: "#C6EFCE", fg: "#006100" },
  { value: "bad", label: "Red", bg: "#FFC7CE", fg: "#9C0006" },
  { value: "warning", label: "Amber", bg: "#FFEB9C", fg: "#9C5700" },
  { value: "custom", label: "Custom", bg: "#DDEBF7", fg: "#1F4E78" },
];

function describe(cf: ConditionalFormatSpec) {
  const rule = RULES.find((r) => r.value === cf.operator)?.label ?? cf.operator;
  if (cf.operator === "between") return `${rule} ${cf.value1} and ${cf.value2}`;
  if (cf.operator === "contains") return `${rule} "${cf.text}"`;
  if (["greater_than", "less_than", "equal"].includes(cf.operator)) return `${rule} ${cf.value1}`;
  if (cf.operator === "color_scale") return `${rule}${cf.reverse ? ", low is green" : ""}`;
  return rule;
}

/** Colour cells by their values. Saved as real Excel conditional formatting. */
export default function ConditionalDialog() {
  const close = useUi((s) => s.close);
  const active = useStore((s) => s.active);
  const rules = useStore((s) => s.workbook.sheets[s.active].conditional_formats ?? []);
  const [range, setRange] = useState(() => selectionRange(useStore.getState().selection));
  const [op, setOp] = useState<CondOperator>("greater_than");
  const [v1, setV1] = useState("");
  const [v2, setV2] = useState("");
  const [text, setText] = useState("");
  const [style, setStyle] = useState<CondStyle>("good");
  const [fill, setFill] = useState("#DDEBF7");
  const [font, setFont] = useState("#1F4E78");
  const [reverse, setReverse] = useState(false);
  const [bar, setBar] = useState("#638EC6");

  const needsOne = ["greater_than", "less_than", "equal", "between"].includes(op);
  const highlight = !["color_scale", "data_bar"].includes(op);
  const validRange = /^\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/i.test(range.trim());
  const ready =
    validRange && (!needsOne || v1.trim() !== "" && !isNaN(+v1)) && (op !== "between" || (v2.trim() !== "" && !isNaN(+v2))) && (op !== "contains" || text.trim() !== "");

  const add = () => {
    const rule: ConditionalFormatSpec = { range: range.trim().toUpperCase(), operator: op };
    if (needsOne) rule.value1 = +v1;
    if (op === "between") rule.value2 = +v2;
    if (op === "contains") rule.text = text.trim();
    if (op === "color_scale") rule.reverse = reverse;
    if (op === "data_bar") rule.fill_color = bar;
    if (highlight) {
      rule.style = style;
      if (style === "custom") Object.assign(rule, { fill_color: fill, font_color: font });
    }
    useStore.getState().setConditionalFormats(active, [...rules, rule]);
  };

  return (
    <Modal
      title="Conditional formatting"
      onClose={close}
      width="max-w-xl"
      footer={
        <>
          <Button onClick={close}>Done</Button>
          <Button primary disabled={!ready} onClick={add}>
            Add rule
          </Button>
        </>
      }
    >
      <div className="grid gap-x-4 sm:grid-cols-2">
        <Field label="Cells" hint="Defaults to your selection">
          <input aria-label="Cells" className={inputClass} value={range} onChange={(e) => setRange(e.target.value)} />
        </Field>
        <Field label="Format cells when">
          <select aria-label="Rule" className={inputClass} value={op} onChange={(e) => setOp(e.target.value as CondOperator)}>
            {RULES.map((r) => (
              <option key={r.value} value={r.value}>
                {r.label}
              </option>
            ))}
          </select>
        </Field>
        {needsOne && (
          <Field label={op === "between" ? "From" : "Value"}>
            <input aria-label="Value" inputMode="decimal" className={inputClass} value={v1} onChange={(e) => setV1(e.target.value)} />
          </Field>
        )}
        {op === "between" && (
          <Field label="To">
            <input aria-label="Upper value" inputMode="decimal" className={inputClass} value={v2} onChange={(e) => setV2(e.target.value)} />
          </Field>
        )}
        {op === "contains" && (
          <Field label="Text">
            <input aria-label="Text" className={inputClass} value={text} onChange={(e) => setText(e.target.value)} placeholder="Overdue" />
          </Field>
        )}
        {op === "color_scale" && (
          <label className="mb-4 flex items-center gap-2 self-end text-sm text-slate-700">
            <input type="checkbox" checked={reverse} onChange={(e) => setReverse(e.target.checked)} />
            Low values are good (green)
          </label>
        )}
        {op === "data_bar" && (
          <Field label="Bar colour">
            <input aria-label="Bar colour" type="color" className="h-9 w-20 rounded border border-slate-300" value={bar} onChange={(e) => setBar(e.target.value)} />
          </Field>
        )}
      </div>
      {highlight && (
        <Field label="Style">
          <div className="flex flex-wrap gap-2">
            {STYLES.map((s) => (
              <button
                key={s.value}
                type="button"
                aria-pressed={style === s.value}
                onClick={() => setStyle(s.value)}
                className={`rounded-md px-3 py-1.5 text-sm ${style === s.value ? "ring-2 ring-brand ring-offset-1" : ""}`}
                style={{ background: s.value === "custom" ? fill : s.bg, color: s.value === "custom" ? font : s.fg }}
              >
                {s.label}
              </button>
            ))}
          </div>
          {style === "custom" && (
            <div className="mt-2 flex gap-4 text-xs text-slate-600">
              <label className="flex items-center gap-1.5">
                Fill <input aria-label="Fill colour" type="color" value={fill} onChange={(e) => setFill(e.target.value)} />
              </label>
              <label className="flex items-center gap-1.5">
                Text <input aria-label="Text colour" type="color" value={font} onChange={(e) => setFont(e.target.value)} />
              </label>
            </div>
          )}
        </Field>
      )}

      <h3 className="mt-2 mb-2 text-xs font-semibold tracking-wide text-slate-500 uppercase">Rules on this sheet</h3>
      {rules.length === 0 && <p className="text-sm text-slate-500">None yet.</p>}
      <ul className="divide-y divide-slate-100">
        {rules.map((cf, i) => {
          const sw = cf.operator === "color_scale"
            ? "linear-gradient(90deg,#F8696B,#FFEB84,#63BE7B)"
            : cf.operator === "data_bar"
              ? cf.fill_color ?? "#638EC6"
              : cf.style === "custom"
                ? cf.fill_color ?? "#DDEBF7"
                : STYLES.find((s) => s.value === (cf.style ?? "good"))!.bg;
          return (
            <li key={i} className="flex items-center gap-3 py-2 text-sm">
              <span className="h-4 w-8 shrink-0 rounded" style={{ background: sw }} />
              <span className="font-mono text-xs text-slate-500">{cf.range}</span>
              <span className="flex-1 text-slate-700">{describe(cf)}</span>
              <button
                aria-label={`Delete rule ${i + 1}`}
                className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-red-600"
                onClick={() => useStore.getState().setConditionalFormats(active, rules.filter((_, j) => j !== i))}
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </li>
          );
        })}
      </ul>
    </Modal>
  );
}
