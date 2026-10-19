"use client";

import { forwardRef, useEffect, useImperativeHandle, useRef, useState, type CSSProperties, type KeyboardEvent } from "react";
import { formulaContext, functionInfo, suggestFunctions } from "@/lib/functions";
import { useUi } from "@/lib/ui";

/**
 * A text input that helps with formulas: suggests function names as you type, shows the
 * arguments of the function you're in, and tells the grid which references to colour.
 */
const FormulaInput = forwardRef<HTMLInputElement, {
  value: string;
  onChange: (v: string) => void;
  onKeyDown?: (e: KeyboardEvent<HTMLInputElement>) => void;
  onBlur?: () => void;
  className?: string;
  style?: CSSProperties;
  autoFocus?: boolean;
  ariaLabel?: string;
  /** Where the suggestion list opens relative to the input. */
  popup?: "below" | "below-wide";
}>(function FormulaInput({ value, onChange, onKeyDown, onBlur, className, style, autoFocus, ariaLabel, popup = "below" }, ref) {
  const input = useRef<HTMLInputElement>(null);
  useImperativeHandle(ref, () => input.current!);
  const [caret, setCaret] = useState(value.length);
  const [pick, setPick] = useState(0);
  const [dismissed, setDismissed] = useState(false);
  const [focused, setFocused] = useState(!!autoFocus);

  const ctx = formulaContext(value, caret);
  const suggestions = focused && ctx.partial && !dismissed ? suggestFunctions(ctx.partial.text) : [];
  const call = focused && ctx.call ? functionInfo(ctx.call.name) : undefined;

  useEffect(() => setPick(0), [ctx.partial?.text]);
  useEffect(() => {
    if (!focused) return;
    useUi.setState({ editFormula: value.startsWith("=") ? value : null });
  }, [value, focused]);
  useEffect(() => () => useUi.setState({ editFormula: null }), []);

  const accept = (name: string) => {
    if (!ctx.partial) return;
    const start = ctx.partial.start;
    const rest = value.slice(caret);
    const hasParen = rest.startsWith("(");
    const next = value.slice(0, start) + name + (hasParen ? "" : "(") + rest;
    onChange(next);
    const pos = start + name.length + 1;
    requestAnimationFrame(() => {
      input.current?.setSelectionRange(pos, pos);
      setCaret(pos);
    });
  };

  const syncCaret = () => setCaret(input.current?.selectionStart ?? value.length);

  return (
    <>
      <input
        ref={input}
        autoFocus={autoFocus}
        aria-label={ariaLabel}
        aria-autocomplete="list"
        aria-expanded={suggestions.length > 0}
        className={className}
        style={style}
        value={value}
        spellCheck={false}
        onChange={(e) => {
          onChange(e.target.value);
          setCaret(e.target.selectionStart ?? e.target.value.length);
          setDismissed(false);
        }}
        onSelect={syncCaret}
        onFocus={() => {
          setFocused(true);
          syncCaret();
        }}
        onBlur={() => {
          setFocused(false);
          useUi.setState({ editFormula: null });
          onBlur?.();
        }}
        onKeyDown={(e) => {
          if (suggestions.length) {
            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
              e.preventDefault();
              e.stopPropagation();
              setPick((p) => (p + (e.key === "ArrowDown" ? 1 : suggestions.length - 1)) % suggestions.length);
              return;
            }
            if (e.key === "Tab" || (e.key === "Enter" && suggestions[pick].name !== ctx.partial?.text.toUpperCase())) {
              e.preventDefault();
              e.stopPropagation();
              accept(suggestions[pick].name);
              return;
            }
            if (e.key === "Escape") {
              e.preventDefault();
              e.stopPropagation();
              setDismissed(true);
              return;
            }
          }
          onKeyDown?.(e);
        }}
      />
      {(suggestions.length > 0 || call) && (
        <div
          className={`absolute top-full left-0 z-50 mt-0.5 rounded-md border border-slate-200 bg-white text-left font-normal text-slate-800 shadow-lg ${
            popup === "below-wide" ? "w-[420px]" : "w-80"
          }`}
          onMouseDown={(e) => e.preventDefault()}
        >
          {suggestions.length > 0 ? (
            <ul role="listbox" aria-label="Function suggestions" className="py-1">
              {suggestions.map((f, i) => (
                <li key={f.name}>
                  <button
                    role="option"
                    aria-selected={i === pick}
                    className={`block w-full px-3 py-1 text-left ${i === pick ? "bg-green-50" : "hover:bg-slate-50"}`}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      accept(f.name);
                    }}
                  >
                    <span className="font-mono text-[12px] font-semibold text-brand">{f.name}</span>
                    <span className="ml-2 text-[11px] text-slate-500">{f.desc}</span>
                  </button>
                </li>
              ))}
              <li className="border-t border-slate-100 px-3 pt-1 text-[10px] text-slate-400">Tab or Enter to insert, Esc to close</li>
            </ul>
          ) : (
            call && (
              <div className="px-3 py-1.5 font-mono text-[12px]" role="note">
                <span className="font-semibold">{call.name}</span>(
                {call.args.map((a, i) => (
                  <span key={i}>
                    {i > 0 && ", "}
                    <span className={i === Math.min(ctx.call!.arg, call.args.length - 1) ? "rounded bg-green-100 px-0.5 font-semibold text-brand" : "text-slate-600"}>
                      {a}
                    </span>
                  </span>
                ))}
                )<div className="mt-0.5 font-sans text-[11px] text-slate-500">{call.desc}</div>
              </div>
            )
          )}
        </div>
      )}
    </>
  );
});

export default FormulaInput;
