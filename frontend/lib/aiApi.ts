// Server calls for the AI cell features, file extraction, live data and scheduled reports.

import { request, send } from "./api";
import type { WorkbookResponse } from "./types";

export const AI_FILL_MAX_ROWS = 200;
export const AI_EVAL_MAX_ITEMS = 100;

export type AiFillOutput = "text" | "number";

/** Fill one value per row from an instruction and that row's input columns. */
export function aiFill(instruction: string, headers: string[], rows: string[][], output: AiFillOutput) {
  return send<{ values: (string | number | null)[] }>("POST", "/api/workbooks/ai-fill", { instruction, headers, rows, output });
}

export interface AiEvalItem {
  key: string;
  prompt: string;
  inputs: string[];
}

/** Answer =AI(...) formulas, keyed so the answers can be matched back. */
export function aiEval(items: AiEvalItem[]) {
  return send<{ results: Record<string, string> }>("POST", "/api/workbooks/ai-eval", { items });
}

/** Read a table out of a photo, screenshot or PDF into a new workbook. */
export async function extractFile(file: File, hint?: string): Promise<WorkbookResponse> {
  const form = new FormData();
  form.append("file", file);
  if (hint?.trim()) form.append("hint", hint.trim());
  const res = await request("/api/workbooks/extract", { method: "POST", body: form });
  return res.json();
}

export type ConnectionKind = "csv" | "gsheet" | "json";

export interface FetchedTable {
  columns: string[];
  rows: (string | number | null)[][];
  fetched_at?: string | number | null;
}

export function fetchData(kind: ConnectionKind, url: string, path?: string | null) {
  return send<FetchedTable>("POST", "/api/data/fetch", { kind, url, path: path || null });
}

export type ReportFrequency = "daily" | "weekly" | "monthly";
export type ReportAttach = "xlsx" | "pdf" | "none";

export interface ReportInput {
  frequency: ReportFrequency;
  /** 0 = Monday ... 6 = Sunday (Python's weekday()), weekly only. */
  weekday: number | null;
  /** Day of the month, 1-28, monthly only. */
  day: number | null;
  /** Hour of the day in UTC. */
  hour: number;
  recipients: string[];
  /** One sheet's name, or null for the whole workbook. */
  sheet: string | null;
  attach: ReportAttach;
  include_summary: boolean;
}

export interface Report extends ReportInput {
  id: string;
  last_sent?: string | number | null;
  next_send?: string | number | null;
}

const reportsPath = (docId: string) => `/api/docs/${encodeURIComponent(docId)}/reports`;

export const reports = {
  list: (docId: string) => send<Report[]>("GET", reportsPath(docId)),
  create: (docId: string, body: ReportInput) => send<Report>("POST", reportsPath(docId), body),
  remove: (docId: string, id: string) => send<void>("DELETE", `${reportsPath(docId)}/${encodeURIComponent(id)}`),
  sendNow: (docId: string, id: string) => send<unknown>("POST", `${reportsPath(docId)}/${encodeURIComponent(id)}/send`),
};

export const errorText = (e: unknown, fallback = "Something went wrong.") => (e instanceof Error && e.message ? e.message : fallback);

// Report times are stored in UTC; people pick them in their own time zone.
// Weekdays are 0 = Monday ... 6 = Sunday. Times use today's UTC offset.

export interface Schedule {
  frequency: ReportFrequency;
  weekday: number | null;
  day: number | null;
  hour: number;
}

/** This week's Monday and this month, as reference dates, so the conversion uses today's daylight-saving offset. */
function reference(local: boolean) {
  const now = new Date();
  const y = local ? now.getFullYear() : now.getUTCFullYear();
  const m = local ? now.getMonth() : now.getUTCMonth();
  const date = local ? now.getDate() : now.getUTCDate();
  const dow = ((local ? now.getDay() : now.getUTCDay()) + 6) % 7;
  return { y, m, monday: date - dow };
}

function convert(s: Schedule, toUtc: boolean): Schedule {
  const ref = reference(toUtc);
  const date = s.frequency === "monthly" ? s.day ?? 1 : ref.monday + (s.frequency === "weekly" ? s.weekday ?? 0 : 0);
  let d = toUtc ? new Date(ref.y, ref.m, date, s.hour) : new Date(Date.UTC(ref.y, ref.m, date, s.hour));
  // Half-hour time zones: the server keeps whole UTC hours, so round to the nearest one.
  if (toUtc && d.getUTCMinutes()) d = new Date(d.getTime() + ((d.getUTCMinutes() >= 30 ? 60 : 0) - d.getUTCMinutes()) * 60_000);
  const get = toUtc
    ? { hour: d.getUTCHours(), dow: d.getUTCDay(), month: d.getUTCMonth(), date: d.getUTCDate() }
    : { hour: d.getHours(), dow: d.getDay(), month: d.getMonth(), date: d.getDate() };
  const out: Schedule = { frequency: s.frequency, weekday: null, day: null, hour: get.hour };
  if (s.frequency === "weekly") out.weekday = (get.dow + 6) % 7;
  if (s.frequency === "monthly") {
    const day = get.month === ref.m ? get.date : 0;
    // A day that would cross into another month (the 1st, or past the 28th) keeps its number.
    out.day = day >= 1 && day <= 28 ? day : s.day ?? 1;
  }
  return out;
}

/** A schedule picked in local time, as the server wants it (UTC). */
export const scheduleToUtc = (s: Schedule) => convert(s, true);
/** A server (UTC) schedule in local time, for showing. */
export const scheduleToLocal = (s: Schedule) => convert(s, false);

export const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

export const hourLabel = (h: number) => new Date(2026, 0, 1, h).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

const ordinal = (n: number) => `${n}${n % 10 === 1 && n !== 11 ? "st" : n % 10 === 2 && n !== 12 ? "nd" : n % 10 === 3 && n !== 13 ? "rd" : "th"}`;

/** "Every week on Monday at 9:00 AM", in local time, from a server (UTC) schedule. */
export function describeSchedule(utc: Schedule): string {
  const s = scheduleToLocal(utc);
  const at = `at ${hourLabel(s.hour)}`;
  if (s.frequency === "weekly") return `Every week on ${WEEKDAYS[s.weekday ?? 0]} ${at}`;
  if (s.frequency === "monthly") return `Every month on the ${ordinal(s.day ?? 1)} ${at}`;
  return `Every day ${at}`;
}

/** A server timestamp (seconds, milliseconds or ISO text) as local date and time. */
export function whenText(v: string | number | null | undefined): string | null {
  if (v === null || v === undefined || v === "") return null;
  const d = typeof v === "number" ? new Date(v < 1e12 ? v * 1000 : v) : new Date(v);
  return Number.isNaN(d.getTime()) ? null : d.toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
}
