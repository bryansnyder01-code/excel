"use client";

import { create } from "zustand";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

/**
 * Which existing chart the chart dialog is editing. Kept here, next to the charts, so the
 * shared dialog state in lib/ui stays a plain "which dialog is open".
 */
interface ChartEditState {
  /** Sheet index and chart index being edited, or null when the dialog adds a new chart. */
  target: { sheet: number; index: number } | null;
}

export const useChartEdit = create<ChartEditState>(() => ({ target: null }));

/** Open the chart dialog pre-filled with chart `index` of the active sheet. */
export function editChart(index: number) {
  useChartEdit.setState({ target: { sheet: useStore.getState().active, index } });
  useUi.getState().open("chart");
}

/** Forget the edit target (the dialog calls this when it closes, so the next open adds a chart). */
export function clearChartEdit() {
  useChartEdit.setState({ target: null });
}
