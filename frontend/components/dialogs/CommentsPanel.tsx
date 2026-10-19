"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { Check, Loader2, MessageSquareText, RotateCcw, Trash2, X } from "lucide-react";
import { useAccount } from "@/lib/account";
import { cloud, type Member } from "@/lib/api";
import { addComment, deleteComment, loadComments, setResolved, useComments, type Thread } from "@/lib/comments";
import { columnLetter } from "@/lib/format";
import { useLive } from "@/lib/live";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

const when = (ms: number) =>
  new Date(ms).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });

/** Comment threads on cells, with replies, resolving, and @mentions that email people. */
export default function CommentsPanel() {
  const close = useUi((s) => s.close);
  const user = useAccount((s) => s.user);
  const docId = useStore((s) => s.docId);
  const sheet = useStore((s) => s.workbook.sheets[s.active]);
  const selection = useStore((s) => s.selection);
  const live = useLive((s) => s.status);
  const { threads, loading, error } = useComments();
  const [show, setShow] = useState<"open" | "resolved" | "cell">("open");
  const [members, setMembers] = useState<Member[]>([]);

  const cell = `${columnLetter(selection.c2)}${selection.r2 + 1}`;
  useEffect(() => {
    if (!user) return;
    loadComments(docId);
    cloud
      .members(docId)
      .then(setMembers)
      .catch(() => setMembers([]));
  }, [user, docId]);

  const list = useMemo(
    () =>
      threads
        .filter((t) => (show === "cell" ? t.sheet === sheet.name && t.cell === cell : show === "resolved" ? t.resolved : !t.resolved))
        .sort((a, b) => (b.comments.at(-1)?.created ?? 0) - (a.comments.at(-1)?.created ?? 0)),
    [threads, show, sheet.name, cell],
  );

  const go = (t: Thread) => {
    const st = useStore.getState();
    const i = st.workbook.sheets.findIndex((s) => s.name === t.sheet);
    const m = /^([A-Z]+)(\d+)$/.exec(t.cell);
    if (i < 0 || !m) return;
    const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const r = Number(m[2]) - 1;
    if (i !== st.active) st.setActive(i);
    st.select({ r1: r, c1: c, r2: r, c2: c });
    useUi.setState({ view: "sheet" });
  };

  return (
    <aside
      role="complementary"
      aria-label="Comments"
      className="fixed top-16 right-3 bottom-14 z-40 flex w-[min(380px,calc(100vw-24px))] flex-col rounded-xl border border-slate-200 bg-white shadow-2xl"
    >
      <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3">
        <MessageSquareText className="h-5 w-5 text-brand" />
        <h2 className="flex-1 text-sm font-semibold text-slate-800">Comments</h2>
        <button aria-label="Close" className="rounded p-1 text-slate-500 hover:bg-slate-100" onClick={close}>
          <X className="h-4 w-4" />
        </button>
      </div>
      {!user ? (
        <div className="p-4 text-sm text-slate-600">
          <p className="mb-3">Sign in to comment on cells and talk with the people you share this workbook with.</p>
          <button className="rounded-md bg-brand px-3 py-1.5 text-xs font-semibold text-white" onClick={() => useUi.getState().open("account")}>
            Sign in
          </button>
        </div>
      ) : (
        <>
          <div role="tablist" aria-label="Which comments" className="flex gap-1 border-b border-slate-100 px-3 py-2 text-xs">
            {(
              [
                ["open", "Open"],
                ["cell", `On ${cell}`],
                ["resolved", "Resolved"],
              ] as const
            ).map(([k, label]) => (
              <button
                key={k}
                role="tab"
                aria-selected={show === k}
                className={`rounded-full px-2.5 py-1 ${show === k ? "bg-brand text-white" : "text-slate-600 hover:bg-slate-100"}`}
                onClick={() => setShow(k)}
              >
                {label}
              </button>
            ))}
          </div>
          <div className="min-h-0 flex-1 space-y-3 overflow-y-auto px-3 py-3">
            {loading && !threads.length && (
              <p className="flex items-center gap-2 text-sm text-slate-500">
                <Loader2 className="h-4 w-4 animate-spin" /> Loading comments...
              </p>
            )}
            {error && <p className="text-sm text-slate-500">{/not found/i.test(error) ? "Comments start once this workbook is saved to your account. Make a change and try again in a moment." : error}</p>}
            {!loading && !error && !list.length && (
              <p className="text-sm text-slate-400">{show === "cell" ? `No comments on ${sheet.name}!${cell} yet.` : show === "resolved" ? "Nothing resolved yet." : "No open comments."}</p>
            )}
            {list.map((t) => (
              <ThreadCard key={t.id} docId={docId} thread={t} members={members} me={user.id} onGo={() => go(t)} />
            ))}
          </div>
          <div className="border-t border-slate-200 p-3">
            <p className="mb-1.5 text-xs text-slate-500">
              New comment on <strong>{sheet.name}!{cell}</strong>
              {live !== "live" && " (shows for others once they open the workbook)"}
            </p>
            <Composer
              members={members}
              placeholder="Add a comment. Type @ to mention someone."
              onSend={async (body) => {
                const existing = threads.find((t) => !t.resolved && t.sheet === sheet.name && t.cell === cell);
                await addComment(docId, sheet.name, cell, body, existing?.id);
              }}
            />
          </div>
        </>
      )}
    </aside>
  );
}

function ThreadCard({ docId, thread, members, me, onGo }: { docId: string; thread: Thread; members: Member[]; me: string; onGo: () => void }) {
  const [replying, setReplying] = useState(false);
  return (
    <div className={`rounded-lg border p-2.5 ${thread.resolved ? "border-slate-200 opacity-75" : "border-violet-200"}`}>
      <div className="mb-1.5 flex items-center gap-2">
        <button className="text-xs font-semibold text-brand hover:underline" onClick={onGo}>
          {thread.sheet}!{thread.cell}
        </button>
        <span className="flex-1" />
        <button
          className="flex items-center gap-1 rounded px-1.5 py-0.5 text-xs text-slate-600 hover:bg-slate-100"
          onClick={() => setResolved(docId, thread, !thread.resolved).catch((e) => useUi.getState().notify(e.message, true))}
        >
          {thread.resolved ? <RotateCcw className="h-3 w-3" /> : <Check className="h-3 w-3" />} {thread.resolved ? "Reopen" : "Resolve"}
        </button>
      </div>
      <ul className="space-y-2">
        {thread.comments.map((c) => (
          <li key={c.id} className="group text-sm">
            <div className="flex items-baseline gap-2">
              <span className="font-medium text-slate-800">{c.author.email.split("@")[0]}</span>
              <span className="text-[11px] text-slate-400">{when(c.created)}</span>
              {c.author.id === me && (
                <button
                  aria-label="Delete comment"
                  className="ml-auto rounded p-0.5 text-slate-400 opacity-0 group-hover:opacity-100 hover:text-red-600 focus:opacity-100"
                  onClick={() => deleteComment(docId, thread, c.id).catch((e) => useUi.getState().notify(e.message, true))}
                >
                  <Trash2 className="h-3 w-3" />
                </button>
              )}
            </div>
            <p className="whitespace-pre-wrap text-slate-700">
              {c.body.split(/(@[^\s@]+@[^\s@]+\.[^\s@,;:!?]+)/).map((part, i) =>
                part.startsWith("@") ? (
                  <span key={i} className="rounded bg-violet-100 px-0.5 text-violet-800">
                    {part}
                  </span>
                ) : (
                  part
                ),
              )}
            </p>
          </li>
        ))}
      </ul>
      {!thread.resolved &&
        (replying ? (
          <div className="mt-2">
            <Composer
              members={members}
              autoFocus
              placeholder="Reply"
              onSend={async (body) => {
                await addComment(docId, thread.sheet, thread.cell, body, thread.id);
                setReplying(false);
              }}
            />
          </div>
        ) : (
          <button className="mt-1.5 text-xs text-brand hover:underline" onClick={() => setReplying(true)}>
            Reply
          </button>
        ))}
    </div>
  );
}

/** A comment box that suggests people to @mention as you type. */
function Composer({ members, placeholder, autoFocus, onSend }: { members: Member[]; placeholder: string; autoFocus?: boolean; onSend: (body: string) => Promise<void> }) {
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);
  const [pick, setPick] = useState(0);
  const ref = useRef<HTMLTextAreaElement>(null);
  const at = /(^|\s)@([^\s@]*)$/.exec(text.slice(0, ref.current?.selectionStart ?? text.length));
  const matches = at ? members.filter((m) => m.email.toLowerCase().includes(at[2].toLowerCase())).slice(0, 5) : [];

  const insert = (m: Member) => {
    const caret = ref.current?.selectionStart ?? text.length;
    const before = text.slice(0, caret).replace(/@([^\s@]*)$/, `@${m.email} `);
    setText(before + text.slice(caret));
    setPick(0);
    requestAnimationFrame(() => ref.current?.focus());
  };

  const submit = async () => {
    if (!text.trim() || busy) return;
    setBusy(true);
    try {
      await onSend(text.trim());
      setText("");
    } catch (e) {
      useUi.getState().notify(e instanceof Error ? e.message : "The comment wasn't saved.", true);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="relative">
      {matches.length > 0 && (
        <ul role="listbox" aria-label="Mention someone" className="absolute bottom-full left-0 z-10 mb-1 w-full rounded-md border border-slate-200 bg-white py-1 shadow-lg">
          {matches.map((m, i) => (
            <li key={m.id}>
              <button
                role="option"
                aria-selected={i === pick}
                className={`w-full px-3 py-1 text-left text-xs ${i === pick ? "bg-slate-100" : ""}`}
                onMouseDown={(e) => {
                  e.preventDefault();
                  insert(m);
                }}
              >
                {m.email}
              </button>
            </li>
          ))}
        </ul>
      )}
      <textarea
        ref={ref}
        autoFocus={autoFocus}
        aria-label={placeholder}
        rows={2}
        className="w-full resize-none rounded-md border border-slate-300 bg-white px-2 py-1.5 text-sm text-slate-800 focus:border-brand focus:outline-none"
        placeholder={placeholder}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (matches.length && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
            e.preventDefault();
            setPick((pick + (e.key === "ArrowDown" ? 1 : matches.length - 1)) % matches.length);
          } else if (matches.length && (e.key === "Enter" || e.key === "Tab")) {
            e.preventDefault();
            insert(matches[pick] ?? matches[0]);
          } else if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            submit();
          }
        }}
      />
      <div className="mt-1 flex justify-end">
        <button className="flex items-center gap-1 rounded-md bg-brand px-3 py-1 text-xs font-semibold text-white disabled:opacity-50" disabled={!text.trim() || busy} onClick={submit}>
          {busy && <Loader2 className="h-3 w-3 animate-spin" />} Comment
        </button>
      </div>
    </div>
  );
}
