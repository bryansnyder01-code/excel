"use client";

import { useCallback, useEffect, useState, type KeyboardEvent } from "react";
import { CalendarClock, Loader2, Mail, Plus, Send, Trash2, X } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { saveCurrent } from "@/components/Workspace";
import { useAccount } from "@/lib/account";
import {
  describeSchedule,
  errorText,
  hourLabel,
  reports as reportsApi,
  scheduleToUtc,
  WEEKDAYS,
  whenText,
  type Report,
  type ReportAttach,
  type ReportFrequency,
} from "@/lib/aiApi";
import { ApiError, cloud } from "@/lib/api";
import { getDoc } from "@/lib/persist";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

const EMAIL = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;
const NOT_SAVED = "This workbook needs to be saved to your account before reports can be scheduled. Make a change so it syncs, then try again.";

interface Draft {
  frequency: ReportFrequency;
  weekday: number;
  day: number;
  hour: number;
  recipients: string[];
  sheet: string;
  attach: ReportAttach;
  summary: boolean;
}

/** Upload the open workbook to the account, so the server can find it. */
async function pushDoc(docId: string) {
  saveCurrent();
  const doc = getDoc(docId);
  if (!doc) throw new Error("Add something to the workbook before scheduling reports.");
  await cloud.put(doc.id, doc.title, doc.updated, doc);
}

/** Run a reports call; if the server doesn't know this workbook yet, save it to the account and retry once. */
async function withDoc<T>(docId: string, call: () => Promise<T>): Promise<T> {
  try {
    return await call();
  } catch (e) {
    if (!(e instanceof ApiError && e.status === 404)) throw e;
  }
  await pushDoc(docId);
  try {
    return await call();
  } catch (e) {
    if (e instanceof ApiError && e.status === 404) throw new Error(NOT_SAVED);
    throw e;
  }
}

/** Email the workbook (or one sheet) on a schedule, with an optional AI summary of what changed. */
export default function ReportsDialog() {
  const close = useUi((s) => s.close);
  const user = useAccount((s) => s.user);
  const docId = useStore((s) => s.docId);
  const sheets = useStore((s) => s.workbook.sheets);
  const [list, setList] = useState<Report[] | null>(null);
  const [draft, setDraft] = useState<Draft | null>(null);
  const [email, setEmail] = useState("");
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setError(null);
    try {
      setList(await withDoc(docId, () => reportsApi.list(docId)));
    } catch (e) {
      setList([]);
      setError(errorText(e, "Couldn't load scheduled reports."));
    }
  }, [docId]);

  useEffect(() => {
    if (user) void load();
  }, [user, load]);

  if (!user) {
    return (
      <Modal title="Scheduled reports" onClose={close} width="max-w-md">
        <p className="text-sm text-slate-600">
          Sign in to have this workbook emailed to you or your team every day, week or month, with an optional AI summary of what changed.
        </p>
        <div className="mt-4 flex justify-end">
          <Button primary onClick={() => useUi.getState().open("account")}>
            Sign in or create an account
          </Button>
        </div>
      </Modal>
    );
  }

  const startNew = () => {
    setError(null);
    setEmail("");
    setDraft({ frequency: "weekly", weekday: 0, day: 1, hour: 9, recipients: [user.email], sheet: "", attach: "xlsx", summary: true });
  };

  const addEmail = (raw: string) => {
    if (!draft) return true;
    const parts = raw.split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean);
    const bad = parts.filter((p) => !EMAIL.test(p));
    const good = parts.filter((p) => EMAIL.test(p) && !draft.recipients.some((r) => r.toLowerCase() === p.toLowerCase()));
    if (good.length) setDraft({ ...draft, recipients: [...draft.recipients, ...good] });
    if (bad.length) {
      setError(`"${bad[0]}" doesn't look like an email address.`);
      setEmail(bad.join(" "));
      return false;
    }
    setError(null);
    setEmail("");
    return true;
  };

  const onEmailKey = (e: KeyboardEvent<HTMLInputElement>) => {
    if (!draft) return;
    if ((e.key === "Enter" || e.key === "," || e.key === ";" || e.key === " ") && email.trim()) {
      e.preventDefault();
      addEmail(email);
    } else if (e.key === "Backspace" && !email && draft.recipients.length) {
      setDraft({ ...draft, recipients: draft.recipients.slice(0, -1) });
    }
  };

  const create = async () => {
    if (!draft) return;
    if (email.trim() && !addEmail(email)) return;
    const typed = email.split(/[\s,;]+/).filter((p) => EMAIL.test(p));
    const recipients = [...draft.recipients, ...typed];
    if (!recipients.length) return setError("Add at least one email address.");
    const utc = scheduleToUtc({
      frequency: draft.frequency,
      weekday: draft.frequency === "weekly" ? draft.weekday : null,
      day: draft.frequency === "monthly" ? draft.day : null,
      hour: draft.hour,
    });
    setBusy("create");
    setError(null);
    try {
      const made = await withDoc(docId, () =>
        reportsApi.create(docId, {
          ...utc,
          recipients: [...new Map(recipients.map((r) => [r.toLowerCase(), r])).values()],
          sheet: draft.sheet || null,
          attach: draft.attach,
          include_summary: draft.summary,
        }),
      );
      setList((l) => [...(l ?? []), made]);
      setDraft(null);
      useUi.getState().notify("Report scheduled.");
    } catch (e) {
      setError(errorText(e, "Couldn't schedule the report."));
    } finally {
      setBusy(null);
    }
  };

  const sendNow = async (r: Report) => {
    setBusy(`send:${r.id}`);
    setError(null);
    try {
      await withDoc(docId, async () => {
        // Send the latest version, not the last synced one.
        await pushDoc(docId);
        return reportsApi.sendNow(docId, r.id);
      });
      useUi.getState().notify(`Report sent to ${r.recipients.join(", ")}.`);
      void load();
    } catch (e) {
      setError(errorText(e, "Couldn't send the report."));
    } finally {
      setBusy(null);
    }
  };

  const remove = async (r: Report) => {
    setBusy(`del:${r.id}`);
    setError(null);
    try {
      await reportsApi.remove(docId, r.id);
      setList((l) => (l ?? []).filter((x) => x.id !== r.id));
    } catch (e) {
      setError(errorText(e, "Couldn't delete the report."));
    } finally {
      setBusy(null);
    }
  };

  const attachLabel = (a: ReportAttach) => (a === "xlsx" ? "Excel file attached" : a === "pdf" ? "PDF attached" : "No attachment");
  const tz = (() => {
    try {
      return Intl.DateTimeFormat().resolvedOptions().timeZone;
    } catch {
      return "your time zone";
    }
  })();
  const select = `${inputClass} w-auto`;

  return (
    <Modal
      title="Scheduled reports"
      onClose={close}
      width="max-w-2xl"
      footer={
        draft ? (
          <>
            <Button onClick={() => (setDraft(null), setError(null))}>Back</Button>
            <Button primary disabled={busy === "create"} onClick={create}>
              {busy === "create" ? "Scheduling..." : "Schedule report"}
            </Button>
          </>
        ) : (
          <>
            <Button onClick={close}>Close</Button>
            <Button primary onClick={startNew}>
              <span className="flex items-center gap-1.5">
                <Plus className="h-4 w-4" /> New report
              </span>
            </Button>
          </>
        )
      }
    >
      {error && <p className="mb-3 text-sm text-red-600">{error}</p>}

      {!draft &&
        (list === null ? (
          <p className="flex items-center gap-2 text-sm text-slate-500">
            <Loader2 className="h-4 w-4 animate-spin" /> Loading reports...
          </p>
        ) : list.length === 0 ? (
          <div className="flex flex-col items-center gap-2 rounded-lg border border-dashed border-slate-300 py-8 text-center text-sm text-slate-500">
            <CalendarClock className="h-6 w-6 text-slate-400" />
            No scheduled reports yet. Create one to have this workbook emailed on a schedule.
          </div>
        ) : (
          <ul className="divide-y divide-slate-100 rounded-md border border-slate-200" aria-label="Scheduled reports">
            {list.map((r) => {
              const next = whenText(r.next_send);
              const last = whenText(r.last_sent);
              return (
                <li key={r.id} className="flex flex-wrap items-start gap-2 px-3 py-2">
                  <div className="min-w-0 flex-1">
                    <div className="text-sm font-medium text-slate-800">{describeSchedule(r)}</div>
                    <div className="truncate text-xs text-slate-600">
                      <Mail className="mr-1 inline h-3 w-3" />
                      {r.recipients.join(", ")}
                    </div>
                    <div className="text-xs text-slate-500">
                      {r.sheet ? `Sheet "${r.sheet}"` : "Whole workbook"} · {attachLabel(r.attach)}
                      {r.include_summary ? " · AI summary" : ""}
                    </div>
                    <div className="text-xs text-slate-400">
                      {next ? `Next: ${next}` : ""}
                      {next && last ? " · " : ""}
                      {last ? `Last sent: ${last}` : next ? "" : "Not sent yet"}
                    </div>
                  </div>
                  <button
                    className="flex items-center gap-1 rounded px-2 py-1 text-xs text-slate-600 hover:bg-slate-100 disabled:opacity-50"
                    disabled={!!busy}
                    onClick={() => sendNow(r)}
                  >
                    {busy === `send:${r.id}` ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Send className="h-3.5 w-3.5" />} Send now
                  </button>
                  <button
                    aria-label="Delete report"
                    className="rounded p-1 text-red-600 hover:bg-red-50 disabled:opacity-50"
                    disabled={!!busy}
                    onClick={() => remove(r)}
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </li>
              );
            })}
          </ul>
        ))}

      {draft && (
        <div>
          <div className="mb-4">
            <span className="mb-1 block text-sm font-medium text-slate-700">When</span>
            <div className="flex flex-wrap items-center gap-2 text-sm text-slate-700">
              Every
              <select
                aria-label="How often"
                className={select}
                value={draft.frequency}
                onChange={(e) => setDraft({ ...draft, frequency: e.target.value as ReportFrequency })}
              >
                <option value="daily">day</option>
                <option value="weekly">week</option>
                <option value="monthly">month</option>
              </select>
              {draft.frequency === "weekly" && (
                <>
                  on
                  <select aria-label="Day of the week" className={select} value={draft.weekday} onChange={(e) => setDraft({ ...draft, weekday: Number(e.target.value) })}>
                    {WEEKDAYS.map((d, i) => (
                      <option key={d} value={i}>
                        {d}
                      </option>
                    ))}
                  </select>
                </>
              )}
              {draft.frequency === "monthly" && (
                <>
                  on day
                  <select aria-label="Day of the month" className={select} value={draft.day} onChange={(e) => setDraft({ ...draft, day: Number(e.target.value) })}>
                    {Array.from({ length: 28 }, (_, i) => (
                      <option key={i} value={i + 1}>
                        {i + 1}
                      </option>
                    ))}
                  </select>
                </>
              )}
              at
              <select aria-label="Time" className={select} value={draft.hour} onChange={(e) => setDraft({ ...draft, hour: Number(e.target.value) })}>
                {Array.from({ length: 24 }, (_, h) => (
                  <option key={h} value={h}>
                    {hourLabel(h)}
                  </option>
                ))}
              </select>
            </div>
            <span className="mt-1 block text-xs text-slate-500">Times are in your time zone ({tz}).</span>
          </div>

          <div className="mb-4">
            <span className="mb-1 block text-sm font-medium text-slate-700">Send to</span>
            <div className="flex flex-wrap items-center gap-1.5 rounded-md border border-slate-300 px-2 py-1.5 focus-within:border-brand">
              {draft.recipients.map((r) => (
                <span key={r} className="flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs text-slate-700">
                  {r}
                  <button
                    aria-label={`Remove ${r}`}
                    className="text-slate-400 hover:text-slate-700"
                    onClick={() => setDraft({ ...draft, recipients: draft.recipients.filter((x) => x !== r) })}
                  >
                    <X className="h-3 w-3" />
                  </button>
                </span>
              ))}
              <input
                aria-label="Add an email address"
                className="min-w-[10rem] flex-1 bg-transparent px-1 py-0.5 text-sm text-slate-800 outline-none"
                placeholder={draft.recipients.length ? "Add another..." : "name@example.com"}
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                onKeyDown={onEmailKey}
                onBlur={() => email.trim() && addEmail(email)}
              />
            </div>
            <span className="mt-1 block text-xs text-slate-500">Press Enter or comma after each address.</span>
          </div>

          <div className="grid gap-x-4 sm:grid-cols-2">
            <Field label="What to send">
              <select className={inputClass} value={draft.sheet} onChange={(e) => setDraft({ ...draft, sheet: e.target.value })}>
                <option value="">The whole workbook</option>
                {sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    Sheet: {s.name}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Attachment">
              <select className={inputClass} value={draft.attach} onChange={(e) => setDraft({ ...draft, attach: e.target.value as ReportAttach })}>
                <option value="xlsx">Excel file (.xlsx)</option>
                <option value="pdf">PDF</option>
                <option value="none">No attachment, just the email</option>
              </select>
            </Field>
          </div>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" className="accent-brand" checked={draft.summary} onChange={(e) => setDraft({ ...draft, summary: e.target.checked })} />
            Include an AI summary of the numbers and what changed
          </label>
        </div>
      )}
    </Modal>
  );
}
