"use client";

// Things several parts of the UI can trigger: chat prompts, uploads, templates and summaries.

import { getUserTemplate } from "./userTemplates";
import { askQuestion, buildPivot, getTemplate, importFile, streamWorkbook } from "./api";
import { evaluateCached } from "./formula";
import { formatValue, numberStyleAt } from "./format";
import { isBlank, useStore } from "./store";
import type { PivotOptions, WorkbookSpec } from "./types";

const errText = (e: unknown) => (e instanceof Error ? e.message : "Something went wrong.");

/** Templates and uploads open as a new document unless the current one is still empty. */
function startDocFor(): void {
  if (!isBlank(useStore.getState().workbook)) useStore.getState().newWorkbook();
}

async function run(userText: string, task: () => Promise<void>, kind?: "question") {
  const st = useStore.getState();
  if (st.busy) return;
  st.addMessage({ role: "user", text: userText, kind });
  st.setBusy(true);
  try {
    await task();
  } catch (e) {
    useStore.getState().addMessage({ role: "assistant", text: errText(e), error: true });
  } finally {
    useStore.getState().setBusy(false);
  }
}

const AUTO_KEY = "sheetcraft:autoApply";

/** Whether AI changes apply straight away instead of waiting for Accept / Discard. */
export function getAutoApply(): boolean {
  try {
    return localStorage.getItem(AUTO_KEY) === "1";
  } catch {
    return false;
  }
}

export function setAutoApply(on: boolean) {
  try {
    localStorage.setItem(AUTO_KEY, on ? "1" : "0");
  } catch {}
}

export function sendPrompt(prompt: string) {
  // A new request builds on whatever is on screen, so a pending change is kept.
  useStore.getState().acceptReview();
  return run(prompt, async () => {
    const { workbook } = useStore.getState();
    const blank = isBlank(workbook);
    const progress = (text: string) => useStore.getState().setProgress(text);
    try {
      const res = blank
        ? await streamWorkbook("generate", { prompt }, progress)
        : await streamWorkbook("edit", { prompt, workbook }, progress);
      const reply = { text: res.summary, issues: res.issues };
      // Changes to an existing workbook are shown for review first; a brand-new one has nothing to compare.
      if (blank || getAutoApply()) useStore.getState().applyResult(res.workbook, shortLabel(prompt), reply);
      else useStore.getState().startReview(res.workbook, shortLabel(prompt), reply);
    } finally {
      useStore.getState().setProgress(null);
    }
  });
}

/** What every sheet shows, formulas calculated, for answering questions about the data. */
export function displayedValues(): Record<string, string[][]> {
  const { workbook } = useStore.getState();
  const computed = evaluateCached(workbook);
  const out: Record<string, string[][]> = {};
  workbook.sheets.forEach((sheet, si) => {
    const width = Math.max(sheet.columns.length, ...sheet.rows.map((r) => r.cells.length));
    const grid = [Array.from({ length: width }, (_, c) => sheet.columns[c]?.header ?? "")];
    sheet.rows.forEach((row, ri) => {
      grid.push(
        Array.from({ length: width }, (_, c) => {
          const v = row.cells[c]?.value;
          const shown = typeof v === "string" && v.startsWith("=") ? computed.get(`${si}:${ri + 2}:${c + 1}`) : v;
          return formatValue(shown, numberStyleAt(sheet, ri + 1, c)).text;
        }),
      );
    });
    out[sheet.name] = grid;
  });
  return out;
}

/** Answer a question about the workbook's data in the chat, without changing anything. */
export function askData(question: string) {
  return run(question, async () => {
    useStore.getState().setProgress("Reading your data");
    try {
      const res = await askQuestion(useStore.getState().workbook, question, displayedValues());
      useStore.getState().addMessage({ role: "assistant", text: res.answer, references: res.references, kind: "answer" });
    } finally {
      useStore.getState().setProgress(null);
    }
  }, "question");
}

export async function uploadFile(file: File) {
  if (useStore.getState().busy) return;
  startDocFor();
  return run(`Uploaded ${file.name}`, async () => {
    const res = await importFile(file);
    useStore.getState().applyResult(res.workbook, `Imported ${file.name}`, { text: res.summary, issues: res.issues });
  });
}

export async function openTemplate(id: string, title: string) {
  if (useStore.getState().busy) return;
  startDocFor();
  return run(`Opened the ${title} template`, async () => {
    const wb: WorkbookSpec = await getTemplate(id);
    useStore.getState().applyResult(wb, `${title} template`, {
      text: `Here's the ${title.toLowerCase()} template. Edit any cell, or tell me how to tailor it.`,
    });
  });
}

/** Open one of your own saved templates as a new workbook. */
export async function openUserTemplate(id: string, title: string) {
  if (useStore.getState().busy) return;
  startDocFor();
  return run(`Opened the ${title} template`, async () => {
    const wb = await getUserTemplate(id);
    useStore.getState().applyResult({ ...wb, title }, `${title} template`, {
      text: `Here's your ${title} template. Edit any cell, or tell me how to tailor it.`,
    });
  });
}

export function runPivot(options: PivotOptions) {
  const label = `Summarise ${options.value_field ?? "rows"} by ${options.row_field}${options.column_field ? ` and ${options.column_field}` : ""}`;
  return run(label, async () => {
    const res = await buildPivot(useStore.getState().workbook, options);
    useStore.getState().applyResult(res.workbook, label, { text: res.summary, issues: res.issues }, true);
  });
}

function shortLabel(prompt: string) {
  const p = prompt.trim().replace(/\s+/g, " ");
  return p.length > 60 ? `${p.slice(0, 57)}...` : p;
}
