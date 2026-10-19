"use client";

// Every action the command palette (and anything else, e.g. the shortcuts dialog) can list and run.

import { downloadWorkbook } from "./api";
import { formatCellRef, parseCellRef } from "./dashboard";
import { normalize, useStore } from "./store";
import { applyTheme, isDark } from "./theme";
import { useUi, type Dialog } from "./ui";

export type CommandGroup = "Go to" | "View" | "Format" | "Rows and columns" | "Data" | "Workbook" | "Tools" | "Help";

export interface Command {
  id: string;
  title: string;
  group: CommandGroup;
  /** Extra words the search matches, e.g. "pivot" for "Summarise by category". */
  keywords?: string;
  /** Keys as written on the keyboard; "Mod" is Ctrl, or Cmd on a Mac. */
  shortcut?: string[];
  run: () => void | Promise<void>;
}

const st = () => useStore.getState();
const ui = () => useUi.getState();

/** Grid actions need the grid on screen. */
const onSheet = (fn: () => void) => () => {
  if (ui().view !== "sheet") useUi.setState({ view: "sheet" });
  useUi.setState({ mobilePane: "sheet" });
  fn();
};

const DIALOGS: { dialog: Exclude<Dialog, null | "palette">; title: string; group: CommandGroup; keywords?: string; shortcut?: string[] }[] = [
  { dialog: "templates", title: "Browse templates", group: "Workbook", keywords: "new start gallery" },
  { dialog: "docs", title: "Open my workbooks", group: "Workbook", keywords: "documents files recent open" },
  { dialog: "saveTemplate", title: "Save as template", group: "Workbook", keywords: "reuse" },
  { dialog: "share", title: "Share workbook", group: "Workbook", keywords: "invite collaborate link" },
  { dialog: "compare", title: "Compare versions", group: "Workbook", keywords: "diff history changes checkpoint" },
  { dialog: "reports", title: "Scheduled reports", group: "Workbook", keywords: "email schedule send" },
  { dialog: "protect", title: "Protect sheet or cells", group: "Workbook", keywords: "lock read-only permissions" },
  { dialog: "pivot", title: "Summarise by category (pivot table)", group: "Data", keywords: "pivot group totals aggregate" },
  { dialog: "dropdown", title: "Add a dropdown list", group: "Data", keywords: "validation options list" },
  { dialog: "chart", title: "Insert chart", group: "Data", keywords: "graph plot visualise" },
  { dialog: "conditional", title: "Conditional formatting", group: "Format", keywords: "highlight colour scale data bar rules" },
  { dialog: "cleanup", title: "Clean up data", group: "Data", keywords: "duplicates trim tidy fix" },
  { dialog: "connections", title: "Data connections", group: "Data", keywords: "import csv json google sheet refresh live" },
  { dialog: "importImage", title: "Import a table from an image", group: "Data", keywords: "photo screenshot ocr picture" },
  { dialog: "aiFill", title: "Fill cells with AI", group: "Tools", keywords: "generate complete autofill" },
  { dialog: "whatif", title: "What-if scenarios", group: "Tools", keywords: "scenario goal seek inputs" },
  { dialog: "check", title: "Check workbook for errors", group: "Tools", keywords: "audit validate problems" },
  { dialog: "comments", title: "Comments", group: "Tools", keywords: "notes discussion threads" },
  { dialog: "cellHistory", title: "Cell history", group: "Tools", keywords: "changes who edited audit trail" },
  { dialog: "account", title: "Account and sign in", group: "Help", keywords: "login profile user" },
  { dialog: "billing", title: "Plan and billing", group: "Help", keywords: "subscription upgrade pay" },
  { dialog: "shortcuts", title: "Keyboard shortcuts", group: "Help", keywords: "keys help", shortcut: ["Mod", "/"] },
];

/** The static commands plus one "Go to sheet" per sheet, built from the current state. */
export function buildCommands(): Command[] {
  const out: Command[] = [];
  const add = (c: Command) => out.push(c);

  // Go to
  st().workbook.sheets.forEach((s, i) =>
    add({ id: `sheet:${s.name}`, title: `Go to sheet: ${s.name}`, group: "Go to", keywords: "tab switch", run: onSheet(() => st().setActive(i)) }),
  );

  // View
  add({ id: "view:dashboard", title: "Show the dashboard", group: "View", keywords: "kpi tiles overview", run: () => useUi.setState({ view: "dashboard", mobilePane: "sheet" }) });
  add({ id: "view:sheet", title: "Show the sheet", group: "View", keywords: "grid cells", run: onSheet(() => {}) });
  add({
    id: "view:dark",
    title: "Toggle dark mode",
    group: "View",
    keywords: "theme light night",
    run: () => applyTheme(isDark() ? "light" : "dark"),
  });
  add({ id: "view:find", title: "Find", group: "View", keywords: "search", shortcut: ["Mod", "F"], run: onSheet(() => useUi.setState({ find: "find" })) });
  add({ id: "view:replace", title: "Find and replace", group: "View", keywords: "search substitute", shortcut: ["Mod", "H"], run: onSheet(() => useUi.setState({ find: "replace" })) });

  // Format
  add({ id: "fmt:bold", title: "Bold", group: "Format", shortcut: ["Mod", "B"], run: onSheet(() => st().toggleStyle("bold")) });
  add({ id: "fmt:italic", title: "Italic", group: "Format", shortcut: ["Mod", "I"], run: onSheet(() => st().toggleStyle("italic")) });
  add({ id: "fmt:underline", title: "Underline", group: "Format", shortcut: ["Mod", "U"], run: onSheet(() => st().toggleStyle("underline")) });
  add({ id: "fmt:wrap", title: "Wrap text", group: "Format", run: onSheet(() => st().toggleStyle("wrap")) });
  add({ id: "fmt:merge", title: "Merge or unmerge cells", group: "Format", keywords: "combine", run: onSheet(() => st().toggleMerge()) });
  add({ id: "fmt:left", title: "Align left", group: "Format", run: onSheet(() => st().styleSelection({ align: "left" })) });
  add({ id: "fmt:center", title: "Align centre", group: "Format", keywords: "center middle", run: onSheet(() => st().styleSelection({ align: "center" })) });
  add({ id: "fmt:right", title: "Align right", group: "Format", run: onSheet(() => st().styleSelection({ align: "right" })) });

  // Rows and columns
  add({ id: "row:above", title: "Insert row above", group: "Rows and columns", keywords: "add", run: onSheet(() => st().insertRow(false)) });
  add({ id: "row:below", title: "Insert row below", group: "Rows and columns", keywords: "add", run: onSheet(() => st().insertRow(true)) });
  add({ id: "row:delete", title: "Delete row(s)", group: "Rows and columns", keywords: "remove", run: onSheet(() => st().deleteRows()) });
  add({ id: "col:left", title: "Insert column left", group: "Rows and columns", keywords: "add", run: onSheet(() => st().insertColumn(false)) });
  add({ id: "col:right", title: "Insert column right", group: "Rows and columns", keywords: "add", run: onSheet(() => st().insertColumn(true)) });
  add({ id: "col:delete", title: "Delete column(s)", group: "Rows and columns", keywords: "remove", run: onSheet(() => st().deleteColumns()) });
  const sheet = () => st().workbook.sheets[st().active];
  add({
    id: "freeze:header",
    title: "Freeze or unfreeze the header row",
    group: "Rows and columns",
    keywords: "pin lock top",
    run: onSheet(() => st().setFreeze(!sheet().freeze_header, sheet().freeze_columns ?? 0)),
  });
  add({ id: "freeze:col", title: "Freeze the first column", group: "Rows and columns", keywords: "pin lock", run: onSheet(() => st().setFreeze(!!sheet().freeze_header, 1)) });
  add({
    id: "freeze:upto",
    title: "Freeze columns up to the selection",
    group: "Rows and columns",
    keywords: "pin lock",
    run: onSheet(() => st().setFreeze(!!sheet().freeze_header, normalize(st().selection).c1)),
  });
  add({ id: "freeze:none", title: "Unfreeze columns", group: "Rows and columns", keywords: "unpin", run: onSheet(() => st().setFreeze(!!sheet().freeze_header, 0)) });

  // Data
  add({ id: "sort:asc", title: "Sort A → Z by the selected column", group: "Data", keywords: "ascending order smallest", run: onSheet(() => st().sortBy(normalize(st().selection).c1, false)) });
  add({ id: "sort:desc", title: "Sort Z → A by the selected column", group: "Data", keywords: "descending order largest", run: onSheet(() => st().sortBy(normalize(st().selection).c1, true)) });

  // Workbook
  add({ id: "edit:undo", title: "Undo", group: "Workbook", shortcut: ["Mod", "Z"], run: () => st().undo() });
  add({ id: "edit:redo", title: "Redo", group: "Workbook", shortcut: ["Mod", "Y"], run: () => st().redo() });
  add({ id: "wb:new", title: "New workbook", group: "Workbook", keywords: "blank create", run: () => st().newWorkbook() });
  add({ id: "wb:sheet", title: "Add a sheet", group: "Workbook", keywords: "tab new", run: onSheet(() => st().addSheet()) });
  add({
    id: "wb:download",
    title: "Download as Excel (.xlsx)",
    group: "Workbook",
    keywords: "export save xlsx file",
    run: async () => {
      try {
        await downloadWorkbook(st().workbook);
      } catch (e) {
        ui().notify(e instanceof Error ? e.message : "The download didn't work.", true);
      }
    },
  });

  for (const d of DIALOGS) add({ id: `dialog:${d.dialog}`, title: d.title, group: d.group, keywords: d.keywords, shortcut: d.shortcut, run: () => ui().open(d.dialog) });
  return out;
}

/** "Go to Summary!B4" for a query that reads as a cell reference, else null. */
export function goToCellCommand(query: string): Command | null {
  const q = query.trim().replace(/^(go\s*to|goto|jump\s*to)\s+/i, "");
  if (!/\d$/.test(q)) return null;
  const { workbook, active } = st();
  const at = parseCellRef(workbook, q, active);
  if (!at) return null;
  const label = formatCellRef(workbook.sheets[at.sheet].name, at.r, at.c);
  return {
    id: "goto-cell",
    title: `Go to cell ${label}`,
    group: "Go to",
    run: onSheet(() => {
      st().setActive(at.sheet);
      st().select({ r1: at.r, c1: at.c, r2: at.r, c2: at.c });
    }),
  };
}

// ---------- Search ----------

/**
 * How well `query` matches `text`: higher is better, null for no match. Contiguous matches and
 * matches at word starts score highest; otherwise the query's letters must appear in order.
 */
export function fuzzyScore(query: string, text: string): number | null {
  const q = query.toLowerCase().trim();
  const t = text.toLowerCase();
  if (!q) return 0;
  const idx = t.indexOf(q);
  if (idx >= 0) return 1000 - idx + (idx === 0 || /\W/.test(t[idx - 1]) ? 200 : 0);
  // Every word of the query appears somewhere.
  const words = q.split(/\s+/);
  if (words.length > 1 && words.every((w) => t.includes(w))) return 600 - words.reduce((n, w) => n + t.indexOf(w), 0);
  // Letters in order, rewarding runs and word starts.
  let score = 0;
  let ti = 0;
  let run = 0;
  for (const ch of q.replace(/\s+/g, "")) {
    const found = t.indexOf(ch, ti);
    if (found < 0) return null;
    run = found === ti ? run + 1 : 0;
    score += 10 + run * 5 + (found === 0 || /\W/.test(t[found - 1]) ? 15 : 0) - Math.min(found - ti, 10);
    ti = found + 1;
  }
  return score;
}

const RECENT_KEY = "sheetcraft:recentCommands";
const RECENT_MAX = 6;

export function recentCommandIds(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]");
    return Array.isArray(raw) ? raw.filter((x): x is string => typeof x === "string") : [];
  } catch {
    return [];
  }
}

export function rememberCommand(id: string) {
  if (id === "goto-cell") return;
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify([id, ...recentCommandIds().filter((x) => x !== id)].slice(0, RECENT_MAX)));
  } catch {}
}

/** Commands matching `query`, best first. With no query: recent commands, then everything in list order. */
export function searchCommands(query: string, commands: Command[] = buildCommands()): { command: Command; recent: boolean }[] {
  const recent = recentCommandIds();
  const rank = (id: string) => recent.indexOf(id);
  const cell = goToCellCommand(query);
  if (!query.trim()) {
    const byId = new Map(commands.map((c) => [c.id, c]));
    const first = recent.map((id) => byId.get(id)).filter((c): c is Command => !!c);
    return [...first.map((command) => ({ command, recent: true })), ...commands.filter((c) => !first.includes(c)).map((command) => ({ command, recent: false }))];
  }
  const scored = commands
    .map((command) => {
      const title = fuzzyScore(query, command.title);
      const extra = fuzzyScore(query, `${command.keywords ?? ""} ${command.group}`);
      const best = Math.max(title ?? -Infinity, extra === null ? -Infinity : extra * 0.6);
      if (!Number.isFinite(best)) return null;
      const r = rank(command.id);
      return { command, recent: r >= 0, score: best + (r >= 0 ? 150 - r * 10 : 0) };
    })
    .filter((x): x is { command: Command; recent: boolean; score: number } => !!x)
    .sort((a, b) => b.score - a.score);
  return [...(cell ? [{ command: cell, recent: false }] : []), ...scored.map(({ command, recent }) => ({ command, recent }))];
}

/** Keys for display: "Ctrl+B", or "⌘B" on a Mac. */
export function formatShortcut(keys: string[], mac = isMac()): string {
  return mac ? keys.map((k) => (k === "Mod" ? "⌘" : k === "Shift" ? "⇧" : k === "Alt" ? "⌥" : k)).join("") : keys.map((k) => (k === "Mod" ? "Ctrl" : k)).join("+");
}

export function isMac(): boolean {
  return typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
}
