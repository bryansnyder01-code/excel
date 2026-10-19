"use client";

import { useEffect, useRef, useState } from "react";
import { AlertTriangle, ArrowUp, Check, CheckCircle2, HelpCircle, History, Loader2, MessageSquare, RotateCcw, Sparkles, Upload, Wand2, XCircle } from "lucide-react";
import TemplateCards from "@/components/dialogs/TemplateCards";
import { askData, sendPrompt, uploadFile } from "@/lib/actions";
import { useAccount } from "@/lib/account";
import { getUsage } from "@/lib/api";
import { parseRange } from "@/lib/cellStyle";
import { isBlank, useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";
import type { ChatMessage, Issue } from "@/lib/types";
import HistoryPanel from "./HistoryPanel";

const EXAMPLES = [
  "Create a monthly small business budget with income, expenses, and a net profit formula",
  "Build a loan amortization schedule for a $250,000 mortgage at 6.5% over 30 years",
  "Make a sales tracker for 5 reps by quarter with totals, averages, a top performer flag and a chart",
];

export const ACCEPT = ".xlsx,.xlsm,.csv,.tsv,.txt";

export default function ChatPanel() {
  const [tab, setTab] = useState<"chat" | "history">("chat");
  const versions = useStore((s) => s.checkpoints.length);
  return (
    <div className="flex h-full flex-col">
      <div className="flex shrink-0 border-b border-slate-200 text-sm">
        <TabButton active={tab === "chat"} onClick={() => setTab("chat")} icon={<MessageSquare className="h-4 w-4" />}>
          Chat
        </TabButton>
        <TabButton active={tab === "history"} onClick={() => setTab("history")} icon={<History className="h-4 w-4" />}>
          Versions{versions ? ` (${versions})` : ""}
        </TabButton>
      </div>
      {tab === "chat" ? <Chat /> : <HistoryPanel />}
    </div>
  );
}

function TabButton({ active, onClick, icon, children }: { active: boolean; onClick: () => void; icon: React.ReactNode; children: React.ReactNode }) {
  return (
    <button
      onClick={onClick}
      className={`flex flex-1 items-center justify-center gap-1.5 py-2.5 ${
        active ? "border-b-2 border-brand font-semibold text-brand" : "text-slate-500 hover:bg-slate-50"
      }`}
    >
      {icon}
      {children}
    </button>
  );
}

function Chat() {
  const messages = useStore((s) => s.messages);
  const busy = useStore((s) => s.busy);
  const progress = useStore((s) => s.progress);
  const reviewing = useStore((s) => !!s.review);
  const blank = useStore((s) => isBlank(s.workbook));
  const [draft, setDraft] = useState("");
  const [mode, setMode] = useState<"change" | "ask">("change");
  const [steps, setSteps] = useState<string[]>([]);
  const listRef = useRef<HTMLDivElement>(null);
  const asking = mode === "ask" && !blank;

  // Keep every step the AI reports during one request, so progress reads as a short checklist.
  useEffect(() => {
    if (!busy) setSteps([]);
    else if (progress) setSteps((s) => (s.at(-1) === progress ? s : [...s, progress]));
  }, [busy, progress]);

  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight, behavior: "smooth" });
  }, [messages.length, busy, steps.length, reviewing]);

  const send = (text: string) => {
    if (!text.trim() || busy) return;
    setDraft("");
    if (asking) askData(text.trim());
    else sendPrompt(text.trim());
  };

  return (
    <>
      <div ref={listRef} className="flex-1 space-y-4 overflow-y-auto p-4">
        {messages.length === 0 ? <Welcome onPick={send} /> : messages.map((m) => <Bubble key={m.id} message={m} />)}
        {reviewing && <ReviewNote />}
        {busy && (
          <div className="space-y-1 text-sm text-slate-500" aria-live="polite">
            {steps.slice(0, -1).map((s, i) => (
              <div key={i} className="flex items-center gap-2 text-slate-400">
                <Check className="h-3.5 w-3.5 text-brand" />
                {s}
              </div>
            ))}
            <div className="flex items-center gap-2">
              {steps.length ? <Loader2 className="h-3.5 w-3.5 animate-spin text-brand" /> : <Sparkles className="h-4 w-4 text-brand" />}
              {steps.at(-1) ?? "Working on it"}
              <span className="sc-dots">
                <span>.</span>
                <span>.</span>
                <span>.</span>
              </span>
            </div>
          </div>
        )}
      </div>
      <form
        className="border-t border-slate-200 p-3"
        onSubmit={(e) => {
          e.preventDefault();
          send(draft);
        }}
      >
        {!blank && (
          <div role="radiogroup" aria-label="What to do" className="mb-2 flex gap-1 rounded-lg bg-slate-100 p-0.5 text-xs">
            {(
              [
                ["change", "Change the workbook", <Wand2 key="w" className="h-3.5 w-3.5" />],
                ["ask", "Ask about the data", <HelpCircle key="h" className="h-3.5 w-3.5" />],
              ] as const
            ).map(([value, label, icon]) => (
              <button
                key={value}
                type="button"
                role="radio"
                aria-checked={mode === value}
                onClick={() => setMode(value)}
                className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 ${
                  mode === value ? "bg-white font-semibold text-brand shadow-sm" : "text-slate-500 hover:text-slate-700"
                }`}
              >
                {icon}
                {label}
              </button>
            ))}
          </div>
        )}
        <div className="flex items-end gap-2 rounded-xl border border-slate-300 bg-white p-2 focus-within:border-brand focus-within:ring-2 focus-within:ring-brand/15">
          <textarea
            aria-label={asking ? "Ask a question about your data" : "Describe your spreadsheet"}
            className="max-h-40 min-h-[44px] flex-1 resize-none bg-transparent px-1 text-sm focus:outline-none"
            placeholder={
              asking
                ? "e.g. Which month had the highest costs?"
                : messages.length
                  ? "Ask for a change, e.g. add a 15% tax column"
                  : "Describe the spreadsheet you need"
            }
            value={draft}
            rows={2}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                send(draft);
              }
            }}
          />
          <button
            type="submit"
            aria-label={asking ? "Ask" : "Send"}
            disabled={busy || !draft.trim()}
            className="flex h-9 w-9 items-center justify-center rounded-lg bg-brand text-white disabled:bg-slate-300"
          >
            <ArrowUp className="h-4 w-4" />
          </button>
        </div>
        <p className="mt-1.5 px-1 text-[11px] text-slate-400">
          {asking ? "Answers use your current numbers and never change the workbook." : "Enter to send, Shift+Enter for a new line. Drop a file anywhere to open it."}
          <Usage busy={busy} />
        </p>
      </form>
    </>
  );
}

/** Shown while an AI change waits for Accept or Discard. */
function ReviewNote() {
  const review = useStore((s) => s.review)!;
  return (
    <div className="mr-4 rounded-2xl rounded-bl-sm border border-amber-300 bg-amber-50 px-3.5 py-2.5 text-sm text-slate-800">
      <p>{review.reply.text}</p>
      <p className="mt-1.5 text-xs text-amber-800">Changed cells are highlighted in the sheet. Keep or discard the change above the grid.</p>
      <div className="mt-2 flex gap-2">
        <button className="rounded-md bg-brand px-3 py-1 text-xs font-semibold text-white hover:bg-brand-dark" onClick={() => useStore.getState().acceptReview()}>
          Keep
        </button>
        <button className="rounded-md border border-slate-300 bg-white px-3 py-1 text-xs text-slate-700 hover:bg-slate-50" onClick={() => useStore.getState().discardReview()}>
          Discard
        </button>
      </div>
    </div>
  );
}

/** Jump to cells an answer refers to and outline them. */
function showReference(ref: { sheet: string; range: string }) {
  const st = useStore.getState();
  const si = st.workbook.sheets.findIndex((s) => s.name.toLowerCase() === ref.sheet.toLowerCase());
  const r = parseRange(ref.range);
  if (si < 0 || !r) return;
  if (si !== st.active) st.setActive(si);
  st.select({ r1: r.r1 - 1, c1: r.c1 - 1, r2: r.r2 - 1, c2: r.c2 - 1 });
  useUi.setState({ highlight: [ref], mobilePane: "sheet" });
}

function Welcome({ onPick }: { onPick: (t: string) => void }) {
  const fileRef = useRef<HTMLInputElement>(null);
  return (
    <div className="pt-2">
      <h2 className="text-lg font-semibold text-slate-800">What should we build?</h2>
      <p className="mt-1 text-sm text-slate-500">Describe a spreadsheet, start from a template, or open a file you already have.</p>

      <button
        onClick={() => fileRef.current?.click()}
        className="mt-4 flex w-full items-center gap-3 rounded-lg border-2 border-dashed border-slate-300 px-3 py-3 text-left text-sm text-slate-600 hover:border-brand hover:bg-green-50"
      >
        <Upload className="h-5 w-5 text-brand" />
        <span>
          <span className="block font-medium text-slate-800">Open an Excel or CSV file</span>
          <span className="text-xs">Then ask the AI to clean it up, analyse it or add formulas</span>
        </span>
      </button>
      <input
        ref={fileRef}
        type="file"
        accept={ACCEPT}
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) uploadFile(f);
          e.target.value = "";
        }}
      />

      <h3 className="mt-5 mb-2 text-xs font-semibold tracking-wide text-slate-500 uppercase">Templates</h3>
      <TemplateCards compact />

      <h3 className="mt-5 mb-2 text-xs font-semibold tracking-wide text-slate-500 uppercase">Or try asking</h3>
      <div className="space-y-2">
        {EXAMPLES.map((ex) => (
          <button
            key={ex}
            onClick={() => onPick(ex)}
            className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-left text-sm text-slate-700 hover:border-brand hover:bg-green-50"
          >
            {ex}
          </button>
        ))}
      </div>
    </div>
  );
}

function Bubble({ message }: { message: ChatMessage }) {
  const canRestore = useStore((s) => !!message.checkpointId && s.checkpoints.some((c) => c.id === message.checkpointId));
  const isLatest = useStore((s) => s.checkpoints.at(-1)?.id === message.checkpointId);
  if (message.role === "user") {
    return (
      <div className={`ml-8 rounded-2xl rounded-br-sm px-3.5 py-2.5 text-sm text-white ${message.kind === "question" ? "bg-[#6d28d9]" : "bg-accent"}`}>
        {message.kind === "question" && <HelpCircle className="mr-1.5 mb-0.5 inline h-3.5 w-3.5 opacity-80" />}
        {message.text}
      </div>
    );
  }
  const issues = message.issues ?? [];
  const fixed = issues.filter((i) => i.fixed);
  const open = issues.filter((i) => !i.fixed);
  return (
    <div className="mr-4">
      <div
        className={`rounded-2xl rounded-bl-sm px-3.5 py-2.5 text-sm whitespace-pre-line ${
          message.error ? "bg-red-50 text-red-700" : message.kind === "answer" ? "border border-violet-200 bg-violet-50 text-slate-800" : "bg-slate-100 text-slate-800"
        }`}
      >
        {message.text}
      </div>
      {!!message.references?.length && (
        <div className="mt-1.5 flex flex-wrap gap-1.5">
          {message.references.map((ref, i) => (
            <button
              key={i}
              className="rounded-full border border-violet-200 bg-white px-2 py-0.5 font-mono text-[11px] text-violet-700 hover:bg-violet-50"
              title="Show these cells"
              onClick={() => showReference(ref)}
            >
              {ref.sheet}!{ref.range}
            </button>
          ))}
        </div>
      )}
      {open.length > 0 && (
        <ul className="mt-2 space-y-1">
          {open.map((i, n) => (
            <IssueChip key={n} issue={i} />
          ))}
        </ul>
      )}
      {fixed.length > 0 && (
        <details className="mt-2 text-xs text-green-700">
          <summary className="cursor-pointer select-none">
            Auto-fixed {fixed.length} {fixed.length === 1 ? "formula" : "formulas"}
          </summary>
          <ul className="mt-1 space-y-1">
            {fixed.map((i, n) => (
              <IssueChip key={n} issue={i} />
            ))}
          </ul>
        </details>
      )}
      {canRestore && !isLatest && (
        <button
          className="mt-1.5 flex items-center gap-1 text-xs text-slate-500 hover:text-brand"
          onClick={() => useStore.getState().restoreCheckpoint(message.checkpointId!)}
        >
          <RotateCcw className="h-3 w-3" /> Restore this version
        </button>
      )}
    </div>
  );
}

function IssueChip({ issue }: { issue: Issue }) {
  const where = [issue.sheet, issue.cell].filter(Boolean).join(" ");
  const Icon = issue.fixed ? CheckCircle2 : issue.severity === "error" ? XCircle : AlertTriangle;
  const color = issue.fixed ? "text-green-700" : issue.severity === "error" ? "text-red-600" : "text-amber-700";
  return (
    <li className={`flex gap-1.5 text-xs ${color}`}>
      <Icon className="mt-0.5 h-3.5 w-3.5 shrink-0" />
      <span>
        {where && <strong className="font-medium">{where}: </strong>}
        {issue.message}
      </span>
    </li>
  );
}

/** How many AI requests are left today, once that's worth knowing (a quarter or less left). */
function Usage({ busy }: { busy: boolean }) {
  const user = useAccount((s) => s.user);
  const [usage, setUsage] = useState<{ used: number; limit: number; signed_in: boolean } | null>(null);
  useEffect(() => {
    if (busy) return;
    getUsage()
      .then(setUsage)
      .catch(() => setUsage(null));
  }, [busy, user]);
  if (!usage || usage.limit <= 0) return null;
  const left = Math.max(0, usage.limit - usage.used);
  if (left > usage.limit / 4) return null;
  return (
    <span className={`block ${left === 0 ? "text-red-600" : "text-amber-600"}`}>
      {left === 0 ? "No AI requests left today." : `${left} AI ${left === 1 ? "request" : "requests"} left today.`}
      {!usage.signed_in && " Sign in for more."}
    </span>
  );
}
