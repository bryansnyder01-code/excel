"use client";

import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { ClipboardPaste, Copy, Eraser, Highlighter, MessageSquarePlus, Scissors } from "lucide-react";
import { copySelection, pasteText } from "@/lib/clipboard";
import { normalize, useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

/** Right-click menu on the grid. */
export default function ContextMenu({ x, y, onClose }: { x: number; y: number; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ x, y });
  const st = useStore.getState();
  const s = normalize(useStore((s) => s.selection));
  const hasNote = useStore((st) => !!st.workbook.sheets[st.active].rows[s.r2 - 1]?.cells[s.c2]?.note);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    setPos({ x: Math.min(x, window.innerWidth - w - 8), y: Math.min(y, window.innerHeight - h - 8) });
  }, [x, y]);

  useEffect(() => {
    const close = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e.type === "mousedown" && ref.current?.contains(e.target as Node)) return;
      onClose();
    };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", close);
    window.addEventListener("blur", close);
    return () => {
      window.removeEventListener("mousedown", close);
      window.removeEventListener("keydown", close);
      window.removeEventListener("blur", close);
    };
  }, [onClose]);

  const act = (fn: () => void | Promise<void>) => async () => {
    onClose();
    await fn();
  };

  const clip = async (cut: boolean) => {
    const { text } = copySelection(cut);
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      useUi.getState().notify(`Your browser blocked the clipboard. Use ${cut ? "Ctrl+X" : "Ctrl+C"} instead.`, true);
    }
  };

  return (
    <div
      ref={ref}
      role="menu"
      className="fixed z-[60] w-56 rounded-lg border border-slate-200 bg-white py-1 text-sm text-slate-700 shadow-xl"
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => e.preventDefault()}
    >
      <Item icon={<Scissors />} label="Cut" hint="Ctrl+X" onClick={act(() => clip(true))} />
      <Item icon={<Copy />} label="Copy" hint="Ctrl+C" onClick={act(() => clip(false))} />
      <Item
        icon={<ClipboardPaste />}
        label="Paste"
        hint="Ctrl+V"
        onClick={act(async () => {
          try {
            pasteText(await navigator.clipboard.readText());
          } catch {
            useUi.getState().notify("Your browser blocked reading the clipboard. Use Ctrl+V instead.", true);
          }
        })}
      />
      <Sep />
      <Item label="Insert row above" onClick={act(() => st.insertRow(false))} />
      <Item label="Insert row below" onClick={act(() => st.insertRow(true))} />
      <Item label="Insert column left" onClick={act(() => st.insertColumn(false))} />
      <Item label="Insert column right" onClick={act(() => st.insertColumn(true))} />
      <Item label={s.r2 > s.r1 ? "Delete rows" : "Delete row"} onClick={act(st.deleteRows)} />
      <Item label={s.c2 > s.c1 ? "Delete columns" : "Delete column"} onClick={act(st.deleteColumns)} />
      <Sep />
      <Item
        icon={<MessageSquarePlus />}
        label={hasNote ? "Edit comment" : "Add comment"}
        hint="Shift+F2"
        disabled={s.r2 === 0}
        onClick={act(() => useUi.setState({ noteEdit: { r: s.r2, c: s.c2 } }))}
      />
      <Item icon={<Highlighter />} label="Conditional formatting..." onClick={act(() => useUi.getState().open("conditional"))} />
      <Item icon={<Eraser />} label="Clear contents" hint="Delete" onClick={act(st.clearSelection)} />
    </div>
  );
}

function Item({ icon, label, hint, onClick, disabled }: { icon?: ReactNode; label: string; hint?: string; onClick: () => void; disabled?: boolean }) {
  return (
    <button
      role="menuitem"
      disabled={disabled}
      onClick={onClick}
      className="flex w-full items-center gap-2.5 px-3 py-1.5 text-left hover:bg-slate-100 disabled:opacity-40 [&_svg]:h-4 [&_svg]:w-4 [&_svg]:text-slate-500"
    >
      <span className="flex w-4 justify-center">{icon}</span>
      <span className="flex-1">{label}</span>
      {hint && <span className="text-xs text-slate-400">{hint}</span>}
    </button>
  );
}

function Sep() {
  return <div className="my-1 h-px bg-slate-200" />;
}
