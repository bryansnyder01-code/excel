import type {
  AskResult,
  ExplainResult,
  Finding,
  Issue,
  PivotOptions,
  TemplateInfo,
  WorkbookResponse,
  WorkbookSpec,
} from "./types";

export const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";
const OFFLINE = `Can't reach the SheetCraft server at ${API_URL}. Is the backend running?`;

let authToken: string | null = null;
/** Signed-in requests carry this token (set by lib/account.ts). */
export function setAuthToken(token: string | null) {
  authToken = token;
}

export class ApiError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
  }
}

export async function request(path: string, init?: RequestInit): Promise<Response> {
  let res: Response;
  const headers = new Headers(init?.headers);
  if (authToken && !headers.has("Authorization")) headers.set("Authorization", `Bearer ${authToken}`);
  try {
    res = await fetch(`${API_URL}${path}`, { ...init, headers });
  } catch {
    throw new Error(OFFLINE);
  }
  if (!res.ok) {
    let detail = `Request failed (${res.status}).`;
    try {
      const data = await res.json();
      if (typeof data.detail === "string") detail = data.detail;
    } catch {}
    throw new ApiError(detail, res.status);
  }
  return res;
}

export async function send<T>(method: string, path: string, body?: unknown): Promise<T> {
  const res = await request(path, {
    method,
    headers: body === undefined ? undefined : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (res.status === 204) return undefined as T;
  return res.json() as Promise<T>;
}

const post = <T,>(path: string, body: unknown) => send<T>("POST", path, body);

/**
 * Generate or edit with live progress. The server streams `progress` events while the AI works,
 * then one `result` (or `error`) event.
 */
export async function streamWorkbook(
  kind: "generate" | "edit",
  body: { prompt: string; workbook?: WorkbookSpec },
  onProgress: (text: string) => void,
): Promise<WorkbookResponse> {
  const res = await request(`/api/workbooks/${kind}/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
    body: JSON.stringify(body),
  });
  if (!res.body) throw new Error("The server's answer was cut off. Please try again.");
  const reader = res.body.pipeThrough(new TextDecoderStream()).getReader();
  let buffer = "";
  for (;;) {
    const { value, done } = await reader.read();
    if (value) buffer += value;
    let cut: number;
    while ((cut = buffer.indexOf("\n\n")) >= 0) {
      const chunk = buffer.slice(0, cut);
      buffer = buffer.slice(cut + 2);
      const data = chunk
        .split("\n")
        .filter((l) => l.startsWith("data:"))
        .map((l) => l.slice(5).trimStart())
        .join("\n");
      if (!data) continue;
      const event = JSON.parse(data);
      if (event.type === "progress") onProgress(event.text);
      else if (event.type === "result") return event as WorkbookResponse;
      else if (event.type === "error") throw new Error(event.message);
    }
    if (done) break;
  }
  throw new Error("The server's answer was cut off. Please try again.");
}

export function askQuestion(workbook: WorkbookSpec, question: string, values: Record<string, string[][]>) {
  return post<AskResult>("/api/workbooks/ask", { workbook, question, values });
}

// Accounts and cloud sync

export interface User {
  id: string;
  email: string;
}

export interface CloudDocInfo {
  id: string;
  title: string;
  updated: number;
  sheets: number;
  shared: boolean;
  /** "owner" for your own workbooks, "editor" for ones shared with you by invite. */
  role?: "owner" | "editor" | "viewer";
  owner_email?: string;
}

export interface Member {
  id: string;
  email: string;
  role: "owner" | "editor" | "viewer";
}

export const auth = {
  signup: (email: string, password: string) => post<{ token: string; user: User }>("/api/auth/signup", { email, password }),
  login: (email: string, password: string) => post<{ token: string; user: User }>("/api/auth/login", { email, password }),
  logout: () => send<void>("POST", "/api/auth/logout"),
  me: () => send<User>("GET", "/api/auth/me"),
  forgot: (email: string) => post<void>("/api/auth/forgot", { email }),
  reset: (token: string, password: string) => post<{ token: string; user: User }>("/api/auth/reset", { token, password }),
  google: (credential: string) => post<{ token: string; user: User }>("/api/auth/google", { credential }),
  config: () => send<{ google_client_id: string | null; password_reset: boolean }>("GET", "/api/auth/config"),
};

export const cloud = {
  list: () => send<CloudDocInfo[]>("GET", "/api/docs"),
  get: (id: string) => send<{ id: string; title: string; updated: number; data: unknown }>("GET", `/api/docs/${encodeURIComponent(id)}`),
  put: (id: string, title: string, updated: number, data: unknown) =>
    send<{ id: string; updated: number }>("PUT", `/api/docs/${encodeURIComponent(id)}`, { title, updated, data }),
  remove: (id: string) => send<void>("DELETE", `/api/docs/${encodeURIComponent(id)}`),
  share: (id: string) => post<{ token: string }>(`/api/docs/${encodeURIComponent(id)}/share`, {}),
  unshare: (id: string) => send<void>("DELETE", `/api/docs/${encodeURIComponent(id)}/share`),
  shared: (token: string) => send<{ title: string; workbook: WorkbookSpec }>("GET", `/api/shared/${encodeURIComponent(token)}`),
  invite: (id: string, role: "editor" | "viewer" = "editor") => post<{ token: string; role: string }>(`/api/docs/${encodeURIComponent(id)}/invite`, { role }),
  uninvite: (id: string, role?: "editor" | "viewer") =>
    send<void>("DELETE", `/api/docs/${encodeURIComponent(id)}/invite${role ? `?role=${role}` : ""}`),
  invites: (id: string) => send<{ editor: string | null; viewer: string | null }>("GET", `/api/docs/${encodeURIComponent(id)}/invites`),
  setRole: (id: string, userId: string, role: "editor" | "viewer") =>
    send<void>("PATCH", `/api/docs/${encodeURIComponent(id)}/members/${encodeURIComponent(userId)}`, { role }),
  accept: (token: string) => post<{ id: string; title: string; role?: Member["role"] }>(`/api/invites/${encodeURIComponent(token)}/accept`, {}),
  members: (id: string) => send<Member[]>("GET", `/api/docs/${encodeURIComponent(id)}/members`),
  removeMember: (id: string, userId: string) => send<void>("DELETE", `/api/docs/${encodeURIComponent(id)}/members/${encodeURIComponent(userId)}`),
};

export interface UserTemplateInfo {
  id: string;
  title: string;
  description: string;
  sheets: number;
  created: number;
}

export const myTemplates = {
  list: () => send<UserTemplateInfo[]>("GET", "/api/templates/mine"),
  get: (id: string) => send<WorkbookSpec>("GET", `/api/templates/mine/${encodeURIComponent(id)}`),
  save: (title: string, description: string, workbook: WorkbookSpec) =>
    post<{ id: string }>("/api/templates/mine", { title, description, workbook }),
  remove: (id: string) => send<void>("DELETE", `/api/templates/mine/${encodeURIComponent(id)}`),
};

/** "Check my workbook": rule checks always, plus an AI review when `useAi`. */
export function checkWorkbook(workbook: WorkbookSpec, values: Record<string, string[][]>, useAi: boolean) {
  return post<{ summary: string; findings: Finding[] }>("/api/workbooks/check", { workbook, values, use_ai: useAi });
}

export function getUsage() {
  return send<{ used: number; limit: number; signed_in: boolean }>("GET", "/api/workbooks/usage");
}

export function generateWorkbook(prompt: string) {
  return post<WorkbookResponse>("/api/workbooks/generate", { prompt });
}

export function editWorkbook(prompt: string, workbook: WorkbookSpec) {
  return post<WorkbookResponse>("/api/workbooks/edit", { prompt, workbook });
}

export function validateWorkbook(workbook: WorkbookSpec) {
  return post<{ workbook: WorkbookSpec; issues: Issue[] }>("/api/workbooks/validate", workbook);
}

export function explainCell(workbook: WorkbookSpec, sheet: string, cell: string, computed: string, question?: string) {
  return post<ExplainResult>("/api/workbooks/explain", { workbook, sheet, cell, computed, question: question || null });
}

export function buildPivot(workbook: WorkbookSpec, options: PivotOptions) {
  return post<WorkbookResponse>("/api/workbooks/pivot", { workbook, ...options });
}

export async function importFile(file: File): Promise<WorkbookResponse> {
  const form = new FormData();
  form.append("file", file);
  const res = await request("/api/workbooks/import", { method: "POST", body: form });
  return res.json();
}

export async function listTemplates(): Promise<TemplateInfo[]> {
  return (await request("/api/templates")).json();
}

export async function getTemplate(id: string): Promise<WorkbookSpec> {
  return (await request(`/api/templates/${encodeURIComponent(id)}`)).json();
}

export async function getCapabilities(): Promise<{ pdf: boolean }> {
  return (await request("/api/workbooks/capabilities")).json();
}

/** Compile the workbook on the server and return the file. */
export async function exportWorkbook(workbook: WorkbookSpec, format: "xlsx" | "pdf" = "xlsx") {
  const res = await request(`/api/workbooks/export?format=${format}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(workbook),
  });
  const disposition = res.headers.get("Content-Disposition") ?? "";
  const name = /filename="([^"]+)"/.exec(disposition)?.[1] ?? `workbook.${format}`;
  return { blob: await res.blob(), name };
}

export function saveBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // Not right away: the desktop app's Save dialog reads the file only after you pick where it goes.
  setTimeout(() => URL.revokeObjectURL(url), 5 * 60_000);
}

export async function downloadWorkbook(workbook: WorkbookSpec, format: "xlsx" | "pdf" = "xlsx") {
  const { blob, name } = await exportWorkbook(workbook, format);
  saveBlob(blob, name);
}
