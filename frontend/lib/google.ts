// "Open in Google Sheets": upload the compiled .xlsx to the user's Google Drive, converted to a
// Google Sheet, then open it. Needs NEXT_PUBLIC_GOOGLE_CLIENT_ID (an OAuth client ID for a web app
// with this site's origin authorised). Uses the drive.file scope, so the app can only see files it creates.

import { exportWorkbook } from "./api";
import type { WorkbookSpec } from "./types";

export const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "";
const SCOPE = "https://www.googleapis.com/auth/drive.file";

interface TokenResponse {
  access_token?: string;
  error?: string;
}

declare global {
  interface Window {
    google?: {
      accounts: {
        oauth2: {
          initTokenClient: (cfg: {
            client_id: string;
            scope: string;
            callback: (r: TokenResponse) => void;
            error_callback?: (e: { type: string }) => void;
          }) => { requestAccessToken: () => void };
        };
        id: {
          initialize: (o: { client_id: string; callback: (r: { credential: string }) => void }) => void;
          renderButton: (el: HTMLElement, o: Record<string, unknown>) => void;
        };
      };
    };
  }
}

let scriptPromise: Promise<void> | null = null;
export function loadGis(): Promise<void> {
  if (window.google?.accounts) return Promise.resolve();
  scriptPromise ??= new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Couldn't load Google sign-in."));
    document.head.appendChild(s);
  });
  return scriptPromise;
}

let cachedToken: { token: string; expires: number } | null = null;

async function getToken(): Promise<string> {
  if (cachedToken && cachedToken.expires > Date.now()) return cachedToken.token;
  await loadGis();
  return new Promise((resolve, reject) => {
    const client = window.google!.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: SCOPE,
      callback: (r) => {
        if (!r.access_token) return reject(new Error("Google sign-in was cancelled."));
        cachedToken = { token: r.access_token, expires: Date.now() + 50 * 60 * 1000 };
        resolve(r.access_token);
      },
      error_callback: () => reject(new Error("Google sign-in was cancelled.")),
    });
    client.requestAccessToken();
  });
}

/** Returns the URL of the new Google Sheet. Call from a click handler so the sign-in popup is allowed. */
export async function openInGoogleSheets(workbook: WorkbookSpec): Promise<string> {
  if (!GOOGLE_CLIENT_ID) throw new Error("Google Sheets export isn't set up on this site yet.");
  // Ask for the token first, while we're still inside the user's click.
  const tokenPromise = getToken();
  const [{ blob }, token] = await Promise.all([exportWorkbook(workbook, "xlsx"), tokenPromise]);

  const metadata = { name: workbook.title || "SheetCraft workbook", mimeType: "application/vnd.google-apps.spreadsheet" };
  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", blob);
  const res = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: form,
  });
  if (!res.ok) {
    if (res.status === 401) cachedToken = null;
    throw new Error(`Google Drive upload failed (${res.status}).`);
  }
  const data = (await res.json()) as { id: string; webViewLink?: string };
  return data.webViewLink ?? `https://docs.google.com/spreadsheets/d/${data.id}/edit`;
}
