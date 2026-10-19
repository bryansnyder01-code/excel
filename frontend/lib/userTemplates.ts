"use client";

// "Save as template": your own starting points. Signed in, they live in the account (and follow
// you to other devices); otherwise in this browser.

import { myTemplates, type UserTemplateInfo } from "./api";
import { useAccount } from "./account";
import type { WorkbookSpec } from "./types";

const KEY = "sheetcraft:templates";
const MAX = 50;

interface LocalTemplate extends UserTemplateInfo {
  workbook: WorkbookSpec;
}

function readLocal(): LocalTemplate[] {
  try {
    return JSON.parse(localStorage.getItem(KEY) ?? "[]");
  } catch {
    return [];
  }
}

function writeLocal(list: LocalTemplate[]) {
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    throw new Error("This browser's storage is full, so the template wasn't saved.");
  }
}

const signedIn = () => !!useAccount.getState().user;

/** A template keeps the layout, formulas and formatting; `keepData: false` empties typed values in data rows. */
export function templateFrom(wb: WorkbookSpec, keepData: boolean): WorkbookSpec {
  const copy: WorkbookSpec = structuredClone(wb);
  if (!keepData)
    for (const sheet of copy.sheets)
      for (const row of sheet.rows) {
        if ((row.kind ?? "data") !== "data") continue;
        for (const cell of row.cells) if (!(typeof cell.value === "string" && cell.value.startsWith("="))) cell.value = null;
      }
  return copy;
}

export async function listUserTemplates(): Promise<UserTemplateInfo[]> {
  if (signedIn()) return myTemplates.list();
  return readLocal().map(({ workbook: _w, ...info }) => info);
}

export async function saveUserTemplate(title: string, description: string, workbook: WorkbookSpec) {
  if (signedIn()) return void (await myTemplates.save(title, description, workbook));
  const list = readLocal();
  if (list.length >= MAX) throw new Error(`You can keep up to ${MAX} templates.`);
  list.unshift({
    id: `local-${Date.now().toString(36)}`,
    title: title.trim() || workbook.title || "My template",
    description: description.trim(),
    sheets: workbook.sheets.length,
    created: Date.now(),
    workbook,
  });
  writeLocal(list);
}

export async function getUserTemplate(id: string): Promise<WorkbookSpec> {
  if (!id.startsWith("local-")) return myTemplates.get(id);
  const t = readLocal().find((x) => x.id === id);
  if (!t) throw new Error("Template not found.");
  return t.workbook;
}

export async function removeUserTemplate(id: string) {
  if (!id.startsWith("local-")) return myTemplates.remove(id);
  writeLocal(readLocal().filter((x) => x.id !== id));
}
