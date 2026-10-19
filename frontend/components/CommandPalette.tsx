"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { Clock, CornerDownLeft, Search } from "lucide-react";
import { buildCommands, formatShortcut, isMac, rememberCommand, searchCommands, type Command } from "@/lib/commands";
import { useUi } from "@/lib/ui";

/** Ctrl+K / Cmd+K: search and run any action by name. */
export default function CommandPalette() {
  const close = useUi((s) => s.close);
  const [query, setQuery] = useState("");
  const [index, setIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const commands = useMemo(() => buildCommands(), []);
  const results = useMemo(() => searchCommands(query, commands), [query, commands]);
  const mac = useMemo(() => isMac(), []);
  const restoreFocus = useRef<Element | null>(null);

  useEffect(() => {
    restoreFocus.current = document.activeElement;
    inputRef.current?.focus();
  }, []);
  useEffect(() => setIndex(0), [query]);
  useEffect(() => {
    listRef.current?.querySelector(`[data-index="${index}"]`)?.scrollIntoView({ block: "nearest" });
  }, [index]);

  const dismiss = () => {
    close();
    (restoreFocus.current as HTMLElement | null)?.focus?.();
  };

  const run = (c: Command) => {
    rememberCommand(c.id);
    // Close first: many commands open another dialog.
    close();
    (restoreFocus.current as HTMLElement | null)?.focus?.();
    void c.run();
  };

  const onKey = (e: React.KeyboardEvent) => {
    const n = results.length;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (n) setIndex((i) => (i + 1) % n);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (n) setIndex((i) => (i - 1 + n) % n);
    } else if (e.key === "Home" && !query) {
      e.preventDefault();
      setIndex(0);
    } else if (e.key === "End" && !query) {
      e.preventDefault();
      setIndex(Math.max(0, n - 1));
    } else if (e.key === "PageDown") {
      e.preventDefault();
      setIndex((i) => Math.min(n - 1, i + 8));
    } else if (e.key === "PageUp") {
      e.preventDefault();
      setIndex((i) => Math.max(0, i - 8));
    } else if (e.key === "Enter") {
      e.preventDefault();
      const hit = results[index];
      if (hit) run(hit.command);
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      dismiss();
    }
  };

  const activeId = results[index] ? `cmd-${index}` : undefined;
  let lastSection = "";

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center bg-black/40 p-4 pt-[12vh]" onMouseDown={(e) => e.target === e.currentTarget && dismiss()}>
      <div role="dialog" aria-modal="true" aria-label="Command palette" className="flex max-h-[70vh] w-full max-w-xl flex-col overflow-hidden rounded-xl bg-white shadow-2xl">
        <div className="flex items-center gap-2 border-b border-slate-200 px-4">
          <Search className="h-4 w-4 shrink-0 text-slate-400" aria-hidden="true" />
          <input
            ref={inputRef}
            role="combobox"
            aria-expanded="true"
            aria-controls="command-list"
            aria-activedescendant={activeId}
            aria-autocomplete="list"
            aria-label="Search commands, sheets, or type a cell like Summary!B4"
            className="h-12 min-w-0 flex-1 bg-transparent text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none"
            placeholder="Type a command, a sheet name, or a cell like B4…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={onKey}
          />
          <kbd className="hidden rounded border border-slate-300 bg-slate-50 px-1.5 py-0.5 font-mono text-[11px] text-slate-500 sm:inline">Esc</kbd>
        </div>
        <ul ref={listRef} id="command-list" role="listbox" aria-label="Commands" className="min-h-0 flex-1 overflow-y-auto py-1">
          {results.length === 0 && <li className="px-4 py-6 text-center text-sm text-slate-500">No matching commands.</li>}
          {results.map(({ command, recent }, i) => {
            const section = !query.trim() ? (recent ? "Recent" : "All commands") : "";
            const header = section && section !== lastSection ? section : "";
            lastSection = section || lastSection;
            return (
              <li key={`${command.id}-${i}`} role="presentation">
                {header && (
                  <div role="presentation" className="px-4 pt-2 pb-1 text-[11px] font-semibold tracking-wide text-slate-400 uppercase">
                    {header}
                  </div>
                )}
                <div
                  id={`cmd-${i}`}
                  role="option"
                  aria-selected={i === index}
                  data-index={i}
                  className={`mx-1 flex cursor-pointer items-center gap-3 rounded-md px-3 py-2 text-sm ${i === index ? "bg-slate-100 text-slate-900" : "text-slate-700"}`}
                  onMouseMove={() => i !== index && setIndex(i)}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => run(command)}
                >
                  {recent && !query.trim() ? <Clock className="h-3.5 w-3.5 shrink-0 text-slate-400" aria-hidden="true" /> : <span className="w-3.5 shrink-0" />}
                  <span className="min-w-0 flex-1 truncate">{command.title}</span>
                  <span className="shrink-0 text-xs text-slate-400">{command.group}</span>
                  {command.shortcut && (
                    <kbd className="shrink-0 rounded border border-slate-300 bg-slate-50 px-1.5 py-0.5 font-mono text-[11px] text-slate-600">
                      {formatShortcut(command.shortcut, mac)}
                    </kbd>
                  )}
                  {i === index && <CornerDownLeft className="h-3.5 w-3.5 shrink-0 text-slate-400" aria-hidden="true" />}
                </div>
              </li>
            );
          })}
        </ul>
        <div className="flex flex-wrap gap-x-4 gap-y-1 border-t border-slate-200 px-4 py-2 text-[11px] text-slate-500">
          <span>
            <kbd className="font-mono">↑</kbd> <kbd className="font-mono">↓</kbd> to choose
          </span>
          <span>
            <kbd className="font-mono">Enter</kbd> to run
          </span>
          <span>
            <kbd className="font-mono">Esc</kbd> to close
          </span>
          <span className="ml-auto">{formatShortcut(["Mod", "K"], mac)} opens this anytime</span>
        </div>
      </div>
    </div>
  );
}
