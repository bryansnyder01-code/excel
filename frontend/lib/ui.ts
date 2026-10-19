"use client";

import { create } from "zustand";

export type Dialog = "templates" | "docs" | "pivot" | "dropdown" | "chart" | "shortcuts" | "conditional" | "account" | "share" | "cleanup" | "whatif" | "check" | "saveTemplate" | "aiFill" | "importImage" | "connections" | "reports" | "compare" | "palette" | "comments" | "cellHistory" | "billing" | "protect" | null;

interface UiState {
  dialog: Dialog;
  /** Find and replace bar: hidden, find only, or with replace. */
  find: null | "find" | "replace";
  /** The formula being typed in the cell editor or formula bar, so the grid can colour its references. */
  editFormula: string | null;
  /** Grid cell (r, c) whose comment is being edited. */
  noteEdit: { r: number; c: number } | null;
  /** Cells an answer pointed at, outlined in the grid until the selection moves. */
  highlight: { sheet: string; range: string }[] | null;
  /** Phone layout: which pane fills the screen. */
  mobilePane: "chat" | "sheet";
  /** What the main area shows: the grid, or the workbook's dashboard page. */
  view: "sheet" | "dashboard";
  toast: { text: string; error?: boolean } | null;
  open: (d: Dialog) => void;
  close: () => void;
  notify: (text: string, error?: boolean) => void;
}

let toastTimer: ReturnType<typeof setTimeout> | undefined;

export const useUi = create<UiState>((set) => ({
  dialog: null,
  find: null,
  editFormula: null,
  noteEdit: null,
  highlight: null,
  mobilePane: "chat",
  view: "sheet",
  toast: null,
  open: (dialog) => set({ dialog }),
  close: () => set({ dialog: null }),
  notify: (text, error) => {
    clearTimeout(toastTimer);
    set({ toast: { text, error } });
    toastTimer = setTimeout(() => set({ toast: null }), 5000);
  },
}));
