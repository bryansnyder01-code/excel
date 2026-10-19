"use client";

import { Fragment, useEffect, useMemo, useRef, useState, type CSSProperties } from "react";
import { ChevronDown, Filter } from "lucide-react";
import { conditionalPainter, dataCellStyle, parseRange, THEME } from "@/lib/cellStyle";
import { copySelection, pasteText } from "@/lib/clipboard";
import { diffWorkbooks } from "@/lib/diff";
import { evaluateCached } from "@/lib/formula";
import { columnLetter, formatValue, numberStyleAt, parseInput } from "@/lib/format";
import { filteredRows, mergeInfo, rowAt, rowLayout, wrappedHeight } from "@/lib/layout";
import { formulaRefs } from "@/lib/refs";
import { isLocked } from "@/lib/access";
import { openThreadCells, useComments } from "@/lib/comments";
import { useLive } from "@/lib/live";
import { normalize, useStore } from "@/lib/store";
import type { Selection } from "@/lib/types";
import { useUi } from "@/lib/ui";
import ContextMenu from "./ContextMenu";
import FilterMenu from "./FilterMenu";
import FormulaInput from "./FormulaInput";
import NoteEditor from "./NoteEditor";

const MIN_ROWS = 30;
const MIN_COLS = 8;
const HEAD_H = 24;
/** Extra rows drawn above and below the visible ones, so fast scrolling doesn't flash. */
const OVERSCAN = 8;
const ROWNUM_W = 48;
/** Colours for the references of the formula being edited, as in Excel. */
const REF_COLORS = ["#2563eb", "#dc2626", "#7c3aed", "#059669", "#d97706", "#db2777"];

interface Editing {
  r: number;
  c: number;
  value: string;
}

const pxToChars = (px: number) => Math.max(3, (px - 10) / 7.5);
const charsToPx = (w: number) => Math.round(w * 7.5 + 10);

export default function SpreadsheetGrid() {
  const workbook = useStore((s) => s.workbook);
  const active = useStore((s) => s.active);
  const selection = useStore((s) => s.selection);
  const review = useStore((s) => s.review);
  const select = useStore((s) => s.select);
  const setCellValue = useStore((s) => s.setCellValue);
  const editFormula = useUi((s) => s.editFormula);
  const noteEdit = useUi((s) => s.noteEdit);
  const highlight = useUi((s) => s.highlight);

  const sheet = workbook.sheets[active];
  const computed = useMemo(() => evaluateCached(workbook), [workbook]);
  const [editing, setEditing] = useState<Editing | null>(null);
  const dragging = useRef(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const [fillTarget, setFillTarget] = useState<Selection | null>(null);
  const fillRef = useRef<Selection | null>(null);
  const [resize, setResize] = useState<{ c: number; startX: number; startW: number; w: number } | null>(null);
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number } | null>(null);
  const [filterMenu, setFilterMenu] = useState<number | null>(null);
  const [rowResize, setRowResize] = useState<{ r: number; startY: number; startH: number; h: number } | null>(null);
  const [view, setView] = useState({ top: 0, height: 800 });

  const nRows = Math.max(sheet.rows.length + 1 + 10, MIN_ROWS); // includes header row
  const nCols = Math.max(sheet.columns.length + 2, MIN_COLS);
  const sel = normalize(selection);
  const [menu, setMenu] = useState<{ r: number; c: number } | null>(null);
  const freezeHeader = sheet.freeze_header !== false;
  const freezeCols = Math.min(sheet.freeze_columns ?? 0, nCols - 1);

  // Dropdown lists (data validation) on this sheet, looked up by grid cell.
  const lists = useMemo(
    () =>
      (sheet.validations ?? [])
        .filter((v) => v.type === "list" && v.options?.length)
        .map((v) => ({ range: parseRange(v.range), options: v.options! }))
        .filter((v) => v.range),
    [sheet.validations],
  );
  const listAt = (r: number, c: number) => {
    if (r === 0) return null;
    const hit = lists.find(({ range }) => range && r + 1 >= range.r1 && r + 1 <= range.r2 && c + 1 >= range.c1 && c + 1 <= range.c2);
    return hit?.options ?? null;
  };
  useEffect(() => setMenu(null), [selection.r2, selection.c2, active]);
  // Answer highlights stay until the selection moves away from where they were shown.
  const shownAt = useRef<string | null>(null);
  useEffect(() => {
    shownAt.current = highlight ? JSON.stringify([selection, active]) : null;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [highlight]);
  useEffect(() => {
    if (shownAt.current && shownAt.current !== JSON.stringify([selection, active])) useUi.setState({ highlight: null });
  }, [selection, active]);

  const rawOf = (r: number, c: number): string => {
    if (r === 0) return sheet.columns[c]?.header ?? "";
    const v = sheet.rows[r - 1]?.cells[c]?.value;
    if (v === null || v === undefined) return "";
    return String(v);
  };

  const valueOf = (r: number, c: number) => {
    const raw = sheet.rows[r - 1]?.cells[c]?.value;
    return typeof raw === "string" && raw.startsWith("=") ? computed.get(`${active}:${r + 1}:${c + 1}`) : raw;
  };

  const displayOf = (r: number, c: number) => {
    if (r === 0) return formatValue(sheet.columns[c]?.header, "text");
    const value = valueOf(r, c);
    return { value, ...formatValue(value, numberStyleAt(sheet, r, c)) };
  };

  const paint = useMemo(
    () => conditionalPainter(sheet, (row, col) => valueOf(row - 1, col - 1) ?? undefined),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet, computed],
  );

  const merges = useMemo(() => mergeInfo(sheet), [sheet]);

  // Column widths follow content, like the builder's auto-fit, unless set by dragging.
  const autoWidths = useMemo(() => {
    const out: number[] = [];
    for (let c = 0; c < nCols; c++) {
      const col = sheet.columns[c];
      if (col?.width) {
        out.push(charsToPx(col.width));
        continue;
      }
      let longest = (col?.header.length ?? 0) + 2;
      sheet.rows.forEach((row, ri) => {
        const cell = row.cells[c];
        if (cell && !cell.wrap && !merges.anchors.has(`${ri + 1}:${c}`)) longest = Math.max(longest, displayOf(ri + 1, c).text.length);
      });
      out.push(Math.min(360, Math.max(90, longest * 7.5 + 20)));
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, computed, nCols]);
  const widths = resize ? autoWidths.map((w, i) => (i === resize.c ? resize.w : w)) : autoWidths;
  const lefts = useMemo(() => {
    const out: number[] = [];
    let x = ROWNUM_W;
    for (const w of widths) {
      out.push(x);
      x += w;
    }
    return out;
  }, [widths]);

  // Which rows to draw and where: filters hide rows, wrapped text and set heights change them.
  const hiddenRows = useMemo(
    () => filteredRows(sheet, (r, c) => displayOf(r, c).text),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet, computed],
  );
  const layout = useMemo(() => {
    const spanWidth = (r: number, c: number) => {
      const m = merges.anchors.get(`${r}:${c}`);
      if (!m) return widths[c] ?? 90;
      let w = 0;
      for (let x = m.c1; x <= m.c2; x++) w += widths[x] ?? 0;
      return w;
    };
    return rowLayout(sheet, nRows, hiddenRows, (r) => {
      const cells = r === 0 ? [] : (sheet.rows[r - 1]?.cells ?? []);
      let h: number | null = null;
      cells.forEach((cell, c) => {
        if (!cell.wrap || merges.covered.has(`${r}:${c}`)) return;
        const text = displayOf(r, c).text;
        if (text) h = Math.max(h ?? 0, wrappedHeight(text, spanWidth(r, c)));
      });
      return h;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, nRows, hiddenRows, widths, merges, computed]);
  const heights = rowResize ? layout.height.map((h, i) => (layout.visible[i] === rowResize.r ? rowResize.h : h)) : layout.height;
  const indexOf = useMemo(() => new Map(layout.visible.map((r, i) => [r, i])), [layout]);

  // Only the rows in view (plus a margin) are drawn, so big sheets stay fast.
  const win = useMemo(() => {
    const n = layout.visible.length;
    const y = view.top - HEAD_H;
    let start = Math.max(1, rowAt(layout, Math.max(0, y)) - OVERSCAN);
    const end = Math.min(n, rowAt(layout, Math.max(0, y + view.height)) + 1 + OVERSCAN);
    // A merged cell that starts above the window must still be drawn.
    for (const m of merges.anchors.values()) {
      const first = indexOf.get(m.r1);
      const last = indexOf.get(m.r2) ?? n;
      if (first !== undefined && first > 0 && first < start && last >= start) start = first;
    }
    return { start, end };
  }, [layout, view, merges, indexOf]);

  useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const update = () => setView({ top: el.scrollTop, height: el.clientHeight });
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Cells the AI just changed, while its change is waiting for review.
  const changed = useMemo(() => {
    if (!review) return null;
    return diffWorkbooks(review.before, workbook).find((d) => d.name === sheet.name)?.cells ?? null;
  }, [review, workbook, sheet.name]);

  // References of the formula being typed, coloured on this sheet.
  const refBoxes = useMemo(() => {
    if (!editFormula) return [];
    return formulaRefs(editFormula)
      .map((ref, i) => ({ ...ref, color: REF_COLORS[i % REF_COLORS.length] }))
      .filter((ref) => !ref.sheet || ref.sheet.toLowerCase() === sheet.name.toLowerCase());
  }, [editFormula, sheet.name]);

  // Viewers can't type at all; editors can't type into the owner's locked cells.
  const role = useLive((s) => (s.status === "live" ? s.role : null));
  const canEditAt = (r: number, c: number) => {
    const why = role === "viewer" ? "You can view this workbook but not change it." : role === "editor" && isLocked(sheet, r, c) ? "That cell is locked by the workbook's owner." : null;
    if (why) useUi.getState().notify(why, true);
    return !why;
  };

  // Cells with open comment threads get a purple corner.
  const threads = useComments((s) => s.threads);
  const threadCells = useMemo(() => openThreadCells(threads, sheet.name), [threads, sheet.name]);

  // Where the other people editing this workbook have their selection.
  const peers = useLive((s) => s.peers);
  const peerBoxes = useMemo(
    () =>
      Object.values(peers)
        .filter((p) => p.sel && p.sel.sheet === sheet.name)
        .map((p) => {
          const n = normalize(p.sel!);
          return { r1: n.r1 + 1, c1: n.c1 + 1, r2: n.r2 + 1, c2: n.c2 + 1, color: p.color, name: p.email.split("@")[0] };
        }),
    [peers, sheet.name],
  );

  const highlightBoxes = useMemo(
    () =>
      (highlight ?? [])
        .filter((h) => h.sheet.toLowerCase() === sheet.name.toLowerCase())
        .map((h) => parseRange(h.range))
        .filter((r): r is NonNullable<typeof r> => !!r),
    [highlight, sheet.name],
  );

  const commit = (e: Editing, move?: [number, number]) => {
    const value = e.r === 0 ? e.value : parseInput(e.value);
    if (rawOf(e.r, e.c) !== e.value) setCellValue(e.r, e.c, value, e.value);
    setEditing(null);
    if (move) {
      const r = Math.min(nRows - 1, Math.max(0, e.r + move[0]));
      const c = Math.min(nCols - 1, Math.max(0, e.c + move[1]));
      select({ r1: r, c1: c, r2: r, c2: c });
    }
    rootRef.current?.focus();
  };

  // Keep the active cell in view while moving with the keyboard.
  useEffect(() => {
    const el = rootRef.current;
    const i = indexOf.get(selection.r2);
    if (!el || i === undefined) return;
    // Rows under the frozen header row can't be seen, so keep the cursor below it.
    const above = HEAD_H + (freezeHeader && i > 0 ? layout.height[0] : 0);
    const y = layout.top[i] + HEAD_H;
    if (y - above < el.scrollTop) el.scrollTop = y - above;
    else if (y + layout.height[i] > el.scrollTop + el.clientHeight) el.scrollTop = y + layout.height[i] - el.clientHeight;
    const c = selection.c2;
    const frozenW = freezeCols > 0 && c >= freezeCols ? lefts[freezeCols] - ROWNUM_W : 0;
    const x = (lefts[c] ?? 0) - ROWNUM_W;
    if (x - frozenW < el.scrollLeft) el.scrollLeft = x - frozenW;
    else if ((lefts[c] ?? 0) + (widths[c] ?? 0) > el.scrollLeft + el.clientWidth) el.scrollLeft = (lefts[c] ?? 0) + (widths[c] ?? 0) - el.clientWidth;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.r2, selection.c2]);

  // Dragging a row's bottom edge.
  useEffect(() => {
    if (!rowResize) return;
    const move = (e: MouseEvent) => setRowResize((r) => r && { ...r, h: Math.max(14, r.startH + e.clientY - r.startY) });
    const up = () => {
      setRowResize((r) => {
        if (r && r.h !== r.startH) useStore.getState().setRowHeight(r.r, (r.h * 3) / 4);
        return null;
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    return () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
    };
  }, [rowResize]);

  useEffect(() => {
    const up = () => {
      dragging.current = false;
      if (fillRef.current) useStore.getState().fill(fillRef.current);
      fillRef.current = null;
      setFillTarget(null);
    };
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  }, []);

  // Dragging a column edge.
  useEffect(() => {
    if (!resize) return;
    const move = (e: MouseEvent) => setResize((r) => r && { ...r, w: Math.max(30, r.startW + e.clientX - r.startX) });
    const up = () => {
      setResize((r) => {
        if (r && r.w !== r.startW) useStore.getState().setColumnWidth(r.c, pxToChars(r.w));
        return null;
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    return () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
    };
  }, [resize]);

  // The fill handle extends the selection along one axis only, whichever the pointer moved further on.
  const fillOver = (r: number, c: number) => {
    const s = normalize(useStore.getState().selection);
    const dr = r < s.r1 ? r - s.r1 : r > s.r2 ? r - s.r2 : 0;
    const dc = c < s.c1 ? c - s.c1 : c > s.c2 ? c - s.c2 : 0;
    let t: Selection | null = null;
    if (Math.abs(dr) >= Math.abs(dc) && dr) t = dr > 0 ? { ...s, r2: r } : { ...s, r1: Math.max(1, r) };
    else if (dc) t = dc > 0 ? { ...s, c2: c } : { ...s, c1: c };
    fillRef.current = t;
    setFillTarget(t);
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (editing || noteEdit) return;
    const st = useStore.getState();
    const mod = e.metaKey || e.ctrlKey;
    const { r1, c1, r2, c2 } = selection;
    // Rows move through what's visible, skipping rows a filter hides.
    const go = (r: number, c: number) => select(e.shiftKey ? { r1, c1, r2: r, c2: c } : { r1: r, c1: c, r2: r, c2: c });
    const move = (dr: number, dc: number) => {
      const i = indexOf.get(r2) ?? 0;
      const r = layout.visible[Math.min(layout.visible.length - 1, Math.max(0, i + dr))];
      go(r, Math.min(nCols - 1, Math.max(0, c2 + dc)));
    };
    const page = Math.max(1, Math.floor(view.height / 28) - 2);
    const lastRow = layout.visible[Math.max(0, indexOf.get(sheet.rows.length) ?? layout.visible.length - 1)];
    const key = e.key.toLowerCase();

    if (mod && key === "z") (e.shiftKey ? st.redo : st.undo)();
    else if (mod && key === "y") st.redo();
    else if (mod && key === "b") st.toggleStyle("bold");
    else if (mod && key === "i") st.toggleStyle("italic");
    else if (mod && key === "u") st.toggleStyle("underline");
    else if (mod && key === "d") st.fillFromEdge("row");
    else if (mod && key === "r") st.fillFromEdge("col");
    else if (mod && key === "a") select({ r1: 0, c1: 0, r2: Math.max(1, sheet.rows.length), c2: Math.max(0, sheet.columns.length - 1) });
    else if (mod && key === "f") useUi.setState({ find: "find" });
    else if (mod && key === "h") useUi.setState({ find: "replace" });
    else if (mod && (key === "/" || key === "?")) useUi.getState().open("shortcuts");
    else if (mod && e.key === "ArrowDown") go(lastRow, c2);
    else if (mod && e.key === "ArrowUp") go(0, c2);
    else if (mod && e.key === "ArrowRight") go(r2, Math.max(0, sheet.columns.length - 1));
    else if (mod && e.key === "ArrowLeft") go(r2, 0);
    else if (mod && e.key === "Home") go(1, 0);
    else if (mod && e.key === "End") go(lastRow, Math.max(0, sheet.columns.length - 1));
    else if (mod) return; // let the browser handle copy, cut and paste
    else if (e.shiftKey && e.key === "F2" && r2 > 0) useUi.setState({ noteEdit: { r: r2, c: c2 } });
    else if (e.key === "ArrowUp") move(-1, 0);
    else if (e.key === "ArrowDown" && e.altKey && listAt(r2, c2)) setMenu({ r: r2, c: c2 });
    else if (e.key === "ArrowDown") move(1, 0);
    else if (e.key === "ArrowLeft") move(0, -1);
    else if (e.key === "ArrowRight") move(0, 1);
    else if (e.key === "Tab") {
      const c = Math.min(nCols - 1, Math.max(0, c2 + (e.shiftKey ? -1 : 1)));
      select({ r1: r2, c1: c, r2, c2: c });
    }
    else if (e.key === "PageDown") move(page, 0);
    else if (e.key === "PageUp") move(-page, 0);
    else if (e.key === "Escape" && menu) setMenu(null);
    else if ((e.key === "Enter" || e.key === "F2") && canEditAt(r2, c2)) setEditing({ r: r2, c: c2, value: rawOf(r2, c2) });
    else if (e.key === "Delete" || e.key === "Backspace") st.clearSelection();
    else if (e.key.length === 1 && canEditAt(r2, c2)) setEditing({ r: r2, c: c2, value: e.key });
    else return;
    e.preventDefault();
  };

  // Sticky positioning for the frozen header row and columns.
  const stickyFor = (r: number, c: number): CSSProperties | undefined => {
    const top = r === 0 && freezeHeader;
    const left = c < freezeCols;
    if (!top && !left) return undefined;
    return {
      position: "sticky",
      top: top ? HEAD_H : undefined,
      left: left ? lefts[c] : undefined,
      zIndex: top && left ? 16 : top ? 15 : 12,
    };
  };

  const inBox = (r: number, c: number, b: { r1: number; r2: number; c1: number; c2: number }) =>
    r + 1 >= b.r1 && r + 1 <= b.r2 && c + 1 >= b.c1 && c + 1 <= b.c2;

  return (
    <div
      ref={rootRef}
      tabIndex={0}
      aria-activedescendant={`sc-cell-${selection.r2}-${selection.c2}`}
      onKeyDown={onKeyDown}
      onScroll={(e) => {
        const el = e.currentTarget;
        // Re-render only when the drawn window would change.
        if (Math.abs(el.scrollTop - view.top) > 4 || el.clientHeight !== view.height) setView({ top: el.scrollTop, height: el.clientHeight });
      }}
      onCopy={(e) => {
        if (editing || (e.target as HTMLElement).tagName === "INPUT") return;
        const { text, html } = copySelection();
        e.clipboardData.setData("text/plain", text);
        e.clipboardData.setData("text/html", html);
        e.preventDefault();
      }}
      onCut={(e) => {
        if (editing || (e.target as HTMLElement).tagName === "INPUT") return;
        const { text, html } = copySelection(true);
        e.clipboardData.setData("text/plain", text);
        e.clipboardData.setData("text/html", html);
        e.preventDefault();
      }}
      onPaste={(e) => {
        if (editing || (e.target as HTMLElement).tagName === "INPUT") return;
        if (pasteText(e.clipboardData.getData("text/plain"))) e.preventDefault();
      }}
      onContextMenu={(e) => {
        const td = (e.target as HTMLElement).closest<HTMLElement>("[data-cell]");
        if (!td) return;
        e.preventDefault();
        const [r, c] = td.dataset.cell!.split(":").map(Number);
        if (!(r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2)) select({ r1: r, c1: c, r2: r, c2: c });
        setCtxMenu({ x: e.clientX, y: e.clientY });
      }}
      className={`sc-grid relative min-h-0 flex-1 overflow-auto focus:outline-none ${resize ? "cursor-col-resize select-none" : ""} ${fillTarget ? "cursor-crosshair" : ""}`}
      aria-label="Spreadsheet grid"
    >
      <table
        role="grid"
        aria-label={`Sheet ${sheet.name}`}
        aria-rowcount={nRows + 1}
        aria-colcount={widths.length + 1}
        aria-multiselectable
        aria-readonly={role === "viewer" || undefined}
        className={`border-separate border-spacing-0 text-[13px] ${sheet.show_gridlines === false ? "sc-gridless" : ""}`}
        style={{ tableLayout: "fixed" }}
      >
        <colgroup>
          <col style={{ width: ROWNUM_W }} />
          {widths.map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
        <thead>
          <tr aria-rowindex={1}>
            <th className="sc-head sticky top-0 left-0 z-30" style={{ height: HEAD_H }} aria-label="Select all" />
            {widths.map((_, c) => (
              <th
                key={c}
                role="columnheader"
                aria-colindex={c + 2}
                aria-label={`Column ${columnLetter(c)}`}
                className={`sc-head sticky top-0 cursor-pointer select-none text-xs font-medium ${
                  c >= sel.c1 && c <= sel.c2 ? "sc-head-on" : ""
                } ${c < freezeCols ? "z-[25]" : "z-20"} ${c === freezeCols - 1 ? "sc-freeze-right" : ""}`}
                style={c < freezeCols ? { left: lefts[c] } : undefined}
                onMouseDown={(e) => {
                  e.preventDefault();
                  rootRef.current?.focus();
                  select(e.shiftKey ? { ...selection, r2: nRows - 1, c2: c } : { r1: 0, c1: c, r2: nRows - 1, c2: c });
                }}
              >
                {columnLetter(c)}
                <span
                  role="separator"
                  aria-orientation="vertical"
                  aria-label={`Resize column ${columnLetter(c)}`}
                  title="Drag to resize, double-click to fit"
                  className="absolute top-0 right-0 z-10 h-full w-1.5 cursor-col-resize hover:bg-brand/50"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    setResize({ c, startX: e.clientX, startW: widths[c], w: widths[c] });
                  }}
                  onDoubleClick={(e) => {
                    e.stopPropagation();
                    useStore.getState().setColumnWidth(c, null);
                  }}
                />
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {[0, ...layout.visible.slice(win.start, win.end)].map((r, n) => {
            const row = r > 0 ? sheet.rows[r - 1] : undefined;
            const kind = row?.kind ?? "data";
            const banded = kind === "data" && r > 0 && r % 2 === 0 && r <= sheet.rows.length;
            const li = n === 0 ? 0 : win.start + n - 1;
            const gapAbove = n === 1 ? layout.top[win.start] - (layout.top[0] + heights[0]) : 0;
            const gapBelow = n === win.end - win.start ? layout.total - (layout.top[li] + layout.height[li]) : 0;
            return (
              <Fragment key={r}>
              {gapAbove > 0 && (
                <tr aria-hidden className="sc-spacer" style={{ height: gapAbove }}>
                  <td colSpan={widths.length + 1} />
                </tr>
              )}
              <tr style={{ height: heights[li] }} aria-rowindex={r + 2}>
                <th
                  role="rowheader"
                  aria-label={`Row ${r + 1}`}
                  className={`sc-head sticky left-0 cursor-pointer select-none text-xs font-medium ${
                    r >= sel.r1 && r <= sel.r2 ? "sc-head-on" : ""
                  } ${r === 0 && freezeHeader ? "z-[17] sc-freeze-bottom" : "z-10"}`}
                  style={r === 0 && freezeHeader ? { top: HEAD_H } : undefined}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    rootRef.current?.focus();
                    select(e.shiftKey ? { ...selection, r2: r, c2: nCols - 1 } : { r1: r, c1: 0, r2: r, c2: nCols - 1 });
                  }}
                >
                  {r + 1}
                  <span
                    role="separator"
                    aria-orientation="horizontal"
                    aria-label={`Resize row ${r + 1}`}
                    title="Drag to resize, double-click to fit"
                    className="absolute right-0 bottom-0 left-0 z-10 h-1.5 cursor-row-resize hover:bg-brand/50"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      setRowResize({ r, startY: e.clientY, startH: heights[li], h: heights[li] });
                    }}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      useStore.getState().setRowHeight(r, null);
                    }}
                  />
                </th>
                {widths.map((_, c) => {
                  if (merges.covered.has(`${r}:${c}`)) return null;
                  const merge = merges.anchors.get(`${r}:${c}`);
                  // Rows of a merge that are drawn right now (filtered or off-screen rows don't count).
                  const rowSpan = merge
                    ? layout.visible.slice(win.start, win.end).filter((x) => x >= merge.r1 && x <= merge.r2).length + (r === 0 ? 1 : 0)
                    : 1;
                  const colSpan = merge ? Math.min(merge.c2, widths.length - 1) - c + 1 : 1;
                  const inSel = r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
                  const isCursor = r === selection.r2 && c === selection.c2;
                  const isEditing = editing?.r === r && editing?.c === c;
                  const isHeader = r === 0;
                  const inSheet = c < sheet.columns.length;
                  const d = displayOf(r, c);
                  const cell = row?.cells[c];

                  const style: CSSProperties = isHeader
                    ? inSheet
                      ? { background: THEME.headerFill, color: "#fff", fontWeight: 600, textAlign: "center" as const }
                      : {}
                    : dataCellStyle(cell, inSheet ? kind : "data", banded && inSheet, d.numeric, paint(r + 1, c + 1, "value" in d ? (d.value ?? undefined) : undefined));
                  if (!isHeader && d.negative && numberStyleAt(sheet, r, c).fmt === "currency" && !cell?.font_color) style.color = "#C00000";
                  if (d.error) style.color = "#C00000";
                  Object.assign(style, stickyFor(r, c));
                  const options = listAt(r, c);
                  const offList = !!options && d.text !== "" && !options.includes(rawOf(r, c));
                  const isFillCorner = !editing && r === sel.r2 && c === sel.c2;
                  const inFill = fillTarget && !inSel && r >= fillTarget.r1 && r <= fillTarget.r2 && c >= fillTarget.c1 && c <= fillTarget.c2;
                  const refs = [...refBoxes, ...peerBoxes].filter((b) => inBox(r, c, b));
                  const peerTag = peerBoxes.find((b) => b.r1 === r + 1 && b.c1 === c + 1);
                  const lit = highlightBoxes.some((b) => inBox(r, c, b));
                  const wrap = !!cell?.wrap;
                  const filterOn = isHeader && inSheet && (sheet.filters ?? []).some((f) => f.column === c);

                  return (
                    <td
                      key={c}
                      id={`sc-cell-${r}-${c}`}
                      role="gridcell"
                      aria-colindex={c + 2}
                      aria-selected={inSel}
                      aria-label={`${columnLetter(c)}${r + 1}${d.text ? `, ${d.text}` : ", empty"}${cell?.note ? ", has note" : ""}`}
                      data-cell={`${r}:${c}`}
                      rowSpan={rowSpan > 1 ? rowSpan : undefined}
                      colSpan={colSpan > 1 ? colSpan : undefined}
                      className={`sc-cell relative px-1.5 ${wrap ? "whitespace-pre-wrap break-words" : "whitespace-nowrap"} ${c === freezeCols - 1 ? "sc-freeze-right" : ""} ${
                        r === 0 && freezeHeader ? "sc-freeze-bottom" : ""
                      }`}
                      style={style}
                      title={cell?.note ? undefined : typeof cell?.value === "string" && cell.value.startsWith("=") ? cell.value : undefined}
                      onMouseDown={(e) => {
                        if (isEditing || e.button !== 0) return;
                        e.preventDefault();
                        rootRef.current?.focus();
                        dragging.current = true;
                        select(e.shiftKey ? { ...selection, r2: r, c2: c } : { r1: r, c1: c, r2: r, c2: c });
                      }}
                      onMouseEnter={() => {
                        if (fillRef.current !== null || fillTarget) fillOver(r, c);
                        else if (dragging.current) select({ ...useStore.getState().selection, r2: r, c2: c });
                      }}
                      onDoubleClick={() => canEditAt(r, c) && setEditing({ r, c, value: rawOf(r, c) })}
                    >
                      {inSel && !isEditing && (
                        <div className={`pointer-events-none absolute inset-0 ${isCursor ? "border-2 border-brand" : "bg-green-600/10"}`} />
                      )}
                      {changed?.has(`${r}:${c}`) && !inSel && (
                        <div data-changed className="sc-changed pointer-events-none absolute inset-0" />
                      )}
                      {inFill && <div className="pointer-events-none absolute inset-0 border border-dashed border-slate-500 bg-slate-500/5" />}
                      {lit && <div className="pointer-events-none absolute inset-0 border-2 border-violet-500 bg-violet-500/10" />}
                      {refs.map((b, i) => (
                        <div
                          key={i}
                          className="pointer-events-none absolute inset-0"
                          style={{
                            background: `${b.color}14`,
                            borderColor: b.color,
                            borderStyle: "solid",
                            borderTopWidth: r + 1 === b.r1 ? 2 : 0,
                            borderBottomWidth: r + 1 === b.r2 ? 2 : 0,
                            borderLeftWidth: c + 1 === b.c1 ? 2 : 0,
                            borderRightWidth: c + 1 === b.c2 ? 2 : 0,
                          }}
                        />
                      ))}
                      {threadCells.has(`${r}:${c}`) && (
                        <button
                          aria-label="Open comments on this cell"
                          title="Comments"
                          className="absolute top-0 left-0 z-[5] h-0 w-0 border-t-[7px] border-r-[7px] border-t-violet-600 border-r-transparent"
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            select({ r1: r, c1: c, r2: r, c2: c });
                            useUi.getState().open("comments");
                          }}
                        />
                      )}
                      {peerTag && (
                        <span
                          className="pointer-events-none absolute -top-3.5 left-0 z-20 rounded-t px-1 text-[10px] leading-[14px] font-medium whitespace-nowrap text-white"
                          style={{ background: peerTag.color }}
                        >
                          {peerTag.name}
                        </span>
                      )}
                      {isEditing ? (
                        <CellEditor
                          editing={editing!}
                          setEditing={setEditing}
                          commit={commit}
                          cancel={() => {
                            setEditing(null);
                            rootRef.current?.focus();
                          }}
                        />
                      ) : (
                        <span
                          className={`relative block ${wrap ? "" : "overflow-hidden text-ellipsis"} ${(options && isCursor) || (isHeader && inSheet) ? "pr-5" : ""}`}
                        >
                          {d.text}
                        </span>
                      )}
                      {isHeader && inSheet && !isEditing && (
                        <button
                          aria-label={`Sort and filter ${sheet.columns[c]?.header || columnLetter(c)}`}
                          title="Sort and filter"
                          className={`absolute top-1/2 right-0.5 z-10 -translate-y-1/2 rounded p-0.5 ${
                            filterOn ? "bg-white text-[#1F4E78]" : "text-white/70 hover:bg-white/20 hover:text-white"
                          }`}
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            e.preventDefault();
                          }}
                          onClick={() => setFilterMenu(filterMenu === c ? null : c)}
                        >
                          {filterOn ? <Filter className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                        </button>
                      )}
                      {filterMenu === c && isHeader && (
                        <FilterMenu column={c} onClose={() => {
                          setFilterMenu(null);
                          rootRef.current?.focus();
                        }} />
                      )}
                      {cell?.note && (
                        <span className="group/note absolute top-0 right-0 z-[5] h-2.5 w-2.5" onMouseDown={(e) => e.stopPropagation()}>
                          <span className="absolute top-0 right-0 border-t-[6px] border-l-[6px] border-t-red-500 border-l-transparent" />
                          <span
                            role="tooltip"
                            className="pointer-events-none absolute top-0 left-3 z-40 hidden w-56 rounded-md border border-amber-200 bg-amber-50 px-2.5 py-1.5 text-left text-xs font-normal whitespace-pre-wrap text-slate-800 shadow-md group-hover/note:block"
                          >
                            {cell.note}
                          </span>
                        </span>
                      )}
                      {offList && (
                        <span
                          title="Not one of this cell's dropdown choices"
                          className="absolute top-0 left-0 border-t-[6px] border-r-[6px] border-t-amber-500 border-r-transparent"
                        />
                      )}
                      {isFillCorner && (
                        <span
                          aria-label="Fill handle"
                          title="Drag to fill"
                          className="absolute -right-[4px] -bottom-[4px] z-20 h-[8px] w-[8px] cursor-crosshair border border-white bg-brand"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            fillRef.current = null;
                            setFillTarget({ ...sel });
                          }}
                        />
                      )}
                      {options && isCursor && !isEditing && (
                        <button
                          aria-label="Show choices"
                          className="absolute top-1/2 right-0.5 z-10 -translate-y-1/2 rounded border border-slate-300 bg-white p-0.5 text-slate-600 shadow-sm hover:bg-slate-100"
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            e.preventDefault();
                            setMenu(menu ? null : { r, c });
                          }}
                        >
                          <ChevronDown className="h-3 w-3" />
                        </button>
                      )}
                      {options && menu?.r === r && menu?.c === c && (
                        <ul
                          role="listbox"
                          className="absolute top-full left-0 z-40 mt-0.5 max-h-56 min-w-full overflow-y-auto rounded-md border border-slate-200 bg-white py-1 text-left font-normal text-slate-800 shadow-lg"
                          onMouseDown={(e) => e.stopPropagation()}
                        >
                          {options.map((o) => (
                            <li key={o}>
                              <button
                                role="option"
                                aria-selected={rawOf(r, c) === o}
                                className={`block w-full px-3 py-1 text-left text-[13px] hover:bg-green-50 ${rawOf(r, c) === o ? "font-semibold text-brand" : ""}`}
                                onMouseDown={(e) => {
                                  e.preventDefault();
                                  setCellValue(r, c, o);
                                  setMenu(null);
                                  rootRef.current?.focus();
                                }}
                              >
                                {o}
                              </button>
                            </li>
                          ))}
                        </ul>
                      )}
                      {noteEdit?.r === r && noteEdit?.c === c && (
                        <NoteEditor
                          r={r}
                          c={c}
                          initial={cell?.note ?? ""}
                          onDone={() => {
                            useUi.setState({ noteEdit: null });
                            rootRef.current?.focus();
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
              {gapBelow > 0 && (
                <tr aria-hidden className="sc-spacer" style={{ height: gapBelow }}>
                  <td colSpan={widths.length + 1} />
                </tr>
              )}
              </Fragment>
            );
          })}
        </tbody>
      </table>
      {ctxMenu && (
        <ContextMenu
          x={ctxMenu.x}
          y={ctxMenu.y}
          onClose={() => {
            setCtxMenu(null);
            rootRef.current?.focus();
          }}
        />
      )}
    </div>
  );
}

function CellEditor({
  editing,
  setEditing,
  commit,
  cancel,
}: {
  editing: Editing;
  setEditing: (e: Editing) => void;
  commit: (e: Editing, move?: [number, number]) => void;
  cancel: () => void;
}) {
  // Enter/Tab/Escape finish the edit themselves; ignore the blur that follows.
  const done = useRef(false);
  const finish = (fn: () => void) => {
    if (done.current) return;
    done.current = true;
    fn();
  };
  return (
    <FormulaInput
      autoFocus
      ariaLabel="Edit cell"
      className="absolute inset-0 z-30 w-full border-2 border-brand bg-white px-1.5 text-[13px] text-slate-900 outline-none"
      style={{ fontWeight: 400, textAlign: "left" }}
      value={editing.value}
      onChange={(v) => setEditing({ ...editing, value: v })}
      onBlur={() => finish(() => commit(editing))}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          finish(() => commit(editing, [e.shiftKey ? -1 : 1, 0]));
        } else if (e.key === "Tab") {
          e.preventDefault();
          finish(() => commit(editing, [0, e.shiftKey ? -1 : 1]));
        } else if (e.key === "Escape") {
          e.preventDefault();
          finish(cancel);
        }
        e.stopPropagation();
      }}
    />
  );
}
