"use client";

import { useState } from "react";
import { Database, Loader2, Pencil, Plus, RefreshCw, Trash2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { errorText, fetchData, type ConnectionKind, type FetchedTable } from "@/lib/aiApi";
import {
  cleanSheetName,
  deleteConnection,
  detectKind,
  KIND_LABELS,
  newConnection,
  refreshConnection,
  refreshedAt,
  REFRESH_CHOICES,
  saveConnection,
} from "@/lib/connections";
import { useStore } from "@/lib/store";
import type { DataConnection } from "@/lib/types";
import { useUi } from "@/lib/ui";

interface Draft {
  id: string | null;
  name: string;
  url: string;
  kind: ConnectionKind | "auto";
  path: string;
  sheet: string;
  refresh: number | null;
}

const blank = (): Draft => ({ id: null, name: "", url: "", kind: "auto", path: "", sheet: "", refresh: 60 });

function ago(ms: number | null): string {
  if (!ms) return "never";
  const s = Math.round((Date.now() - ms) / 1000);
  if (s < 60) return "just now";
  if (s < 3600) return `${Math.round(s / 60)} min ago`;
  if (s < 86400) return `${Math.round(s / 3600)} h ago`;
  return new Date(ms).toLocaleDateString();
}

/** Pull a CSV, Google Sheet or JSON feed into a sheet, refreshed by hand or on a schedule. */
export default function ConnectionsDialog() {
  const close = useUi((s) => s.close);
  const connections = useStore((s) => s.workbook.connections ?? []);
  const [draft, setDraft] = useState<Draft | null>(null);
  const [test, setTest] = useState<FetchedTable | null>(null);
  const [testing, setTesting] = useState(false);
  const [busyIds, setBusyIds] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);

  const kindOf = (d: Draft): ConnectionKind => (d.kind === "auto" ? detectKind(d.url) : d.kind);
  const update = (patch: Partial<Draft>) => {
    setDraft((d) => (d ? { ...d, ...patch } : d));
    if ("url" in patch || "kind" in patch || "path" in patch) setTest(null);
  };

  const refresh = async (c: DataConnection) => {
    setError(null);
    setBusyIds((ids) => [...ids, c.id]);
    try {
      const table = await refreshConnection(c.id);
      if (table) useUi.getState().notify(`Loaded ${table.rows.length} rows into "${c.sheet}". Undo with Ctrl+Z.`);
    } catch (e) {
      setError(`Couldn't refresh "${c.name}": ${errorText(e)}`);
    } finally {
      setBusyIds((ids) => ids.filter((x) => x !== c.id));
    }
  };

  const runTest = async () => {
    if (!draft) return;
    setTesting(true);
    setError(null);
    setTest(null);
    try {
      setTest(await fetchData(kindOf(draft), draft.url.trim(), draft.path.trim() || null));
    } catch (e) {
      setError(errorText(e, "Couldn't fetch that URL."));
    } finally {
      setTesting(false);
    }
  };

  const save = async () => {
    if (!draft) return;
    const kind = kindOf(draft);
    const fields = {
      name: draft.name.trim() || draft.sheet.trim() || "Live data",
      kind,
      url: draft.url.trim(),
      path: kind === "json" ? draft.path.trim() || null : null,
      sheet: cleanSheetName(draft.sheet || draft.name || "Live data"),
      refresh_minutes: draft.refresh,
    };
    const prev = draft.id ? connections.find((c) => c.id === draft.id) : undefined;
    const conn: DataConnection = prev ? { ...prev, ...fields } : newConnection({ ...fields, last_refreshed: null });
    saveConnection(conn);
    setDraft(null);
    setTest(null);
    await refresh(conn);
  };

  const validUrl = (() => {
    try {
      const u = new URL(draft?.url.trim() ?? "");
      return u.protocol === "https:" || u.protocol === "http:";
    } catch {
      return false;
    }
  })();

  return (
    <Modal
      title="Live data connections"
      onClose={close}
      width="max-w-2xl"
      footer={
        draft ? (
          <>
            <Button onClick={() => (setDraft(null), setTest(null), setError(null))}>Back</Button>
            <Button disabled={!validUrl || testing} onClick={runTest}>
              {testing ? "Testing..." : "Test"}
            </Button>
            <Button primary disabled={!validUrl} onClick={save}>
              {draft.id ? "Save and refresh" : "Add and load data"}
            </Button>
          </>
        ) : (
          <>
            <Button onClick={close}>Close</Button>
            <Button primary onClick={() => (setDraft(blank()), setError(null))}>
              <span className="flex items-center gap-1.5">
                <Plus className="h-4 w-4" /> Add a connection
              </span>
            </Button>
          </>
        )
      }
    >
      {error && <p className="mb-3 text-sm text-red-600">{error}</p>}

      {!draft && (
        <>
          <p className="mb-3 text-sm text-slate-600">
            Bring in data from a CSV link, a Google Sheet (shared as &quot;anyone with the link&quot;) or a JSON API. Each refresh replaces the rows of
            its sheet; formulas and charts on other sheets that use it update too.
          </p>
          {connections.length === 0 ? (
            <div className="flex flex-col items-center gap-2 rounded-lg border border-dashed border-slate-300 py-8 text-sm text-slate-500">
              <Database className="h-6 w-6 text-slate-400" />
              No connections yet.
            </div>
          ) : (
            <ul className="divide-y divide-slate-100 rounded-md border border-slate-200" aria-label="Connections">
              {connections.map((c) => {
                const busy = busyIds.includes(c.id);
                return (
                  <li key={c.id} className="flex flex-wrap items-center gap-2 px-3 py-2">
                    <div className="min-w-0 flex-1">
                      <div className="truncate text-sm font-medium text-slate-800">{c.name}</div>
                      <div className="truncate text-xs text-slate-500">
                        {KIND_LABELS[c.kind]} into <strong className="font-medium text-slate-600">{c.sheet}</strong> ·{" "}
                        {REFRESH_CHOICES.find((x) => x.minutes === (c.refresh_minutes ?? null))?.label ?? `Every ${c.refresh_minutes} minutes`} · refreshed{" "}
                        {ago(refreshedAt(c))}
                      </div>
                      <div className="truncate text-xs text-slate-400" title={c.url}>
                        {c.url}
                      </div>
                    </div>
                    <button
                      className="flex items-center gap-1 rounded px-2 py-1 text-xs text-slate-600 hover:bg-slate-100 disabled:opacity-50"
                      disabled={busy}
                      onClick={() => refresh(c)}
                    >
                      {busy ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <RefreshCw className="h-3.5 w-3.5" />} Refresh now
                    </button>
                    <button
                      aria-label={`Edit ${c.name}`}
                      className="rounded p-1 text-slate-500 hover:bg-slate-100"
                      onClick={() =>
                        setDraft({ id: c.id, name: c.name, url: c.url, kind: c.kind, path: c.path ?? "", sheet: c.sheet, refresh: c.refresh_minutes ?? null })
                      }
                    >
                      <Pencil className="h-3.5 w-3.5" />
                    </button>
                    <button
                      aria-label={`Delete ${c.name}`}
                      className="rounded p-1 text-red-600 hover:bg-red-50"
                      onClick={() => {
                        deleteConnection(c.id);
                        useUi.getState().notify(`Removed "${c.name}". Its sheet and data stay.`);
                      }}
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
          <p className="mt-3 text-xs text-slate-500">
            Scheduled refreshes run while this workbook is open, and on the server for workbooks saved to your account.
          </p>
        </>
      )}

      {draft && (
        <div>
          <Field label="URL" hint="A link to a .csv file, a Google Sheet, or a JSON API that returns a list of records.">
            <input
              className={inputClass}
              value={draft.url}
              autoFocus
              placeholder="https://..."
              inputMode="url"
              onChange={(e) => update({ url: e.target.value })}
            />
          </Field>
          <div className="grid gap-x-4 sm:grid-cols-2">
            <Field label="Type">
              <select className={inputClass} value={draft.kind} onChange={(e) => update({ kind: e.target.value as Draft["kind"] })}>
                <option value="auto">Detect automatically{draft.url ? ` (${KIND_LABELS[detectKind(draft.url)]})` : ""}</option>
                {(Object.keys(KIND_LABELS) as ConnectionKind[]).map((k) => (
                  <option key={k} value={k}>
                    {KIND_LABELS[k]}
                  </option>
                ))}
              </select>
            </Field>
            {kindOf(draft) === "json" ? (
              <Field label="Path to the records (optional)" hint='Dotted path to the list, e.g. "data.items".'>
                <input className={inputClass} value={draft.path} onChange={(e) => update({ path: e.target.value })} placeholder="data.items" />
              </Field>
            ) : (
              <div />
            )}
            <Field label="Name">
              <input className={inputClass} value={draft.name} placeholder="Daily sales" onChange={(e) => update({ name: e.target.value })} />
            </Field>
            <Field label="Sheet to fill" hint="Created if it doesn't exist. Its rows are replaced on every refresh.">
              <input className={inputClass} value={draft.sheet} placeholder={cleanSheetName(draft.name || "Live data")} onChange={(e) => update({ sheet: e.target.value })} />
            </Field>
            <Field label="Refresh">
              <select
                className={inputClass}
                value={draft.refresh ?? ""}
                onChange={(e) => update({ refresh: e.target.value === "" ? null : Number(e.target.value) })}
              >
                {REFRESH_CHOICES.map((r) => (
                  <option key={r.label} value={r.minutes ?? ""}>
                    {r.label}
                  </option>
                ))}
              </select>
            </Field>
          </div>
          {testing && (
            <p className="flex items-center gap-2 text-sm text-slate-600">
              <Loader2 className="h-4 w-4 animate-spin" /> Fetching...
            </p>
          )}
          {test && (
            <div className="rounded-md bg-slate-50 px-3 py-2 text-sm text-slate-700" role="status">
              <p className="font-medium text-emerald-700">
                Found {test.rows.length} {test.rows.length === 1 ? "row" : "rows"} and {test.columns.length} {test.columns.length === 1 ? "column" : "columns"}.
              </p>
              <p className="mt-1 text-xs text-slate-600">{test.columns.join(", ") || "(no column names)"}</p>
            </div>
          )}
        </div>
      )}
    </Modal>
  );
}
