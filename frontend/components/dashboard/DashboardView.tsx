"use client";

import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import { Check, GripVertical, LayoutDashboard, Pencil, Plus, Settings2, Sparkles, Trash2 } from "lucide-react";
import MiniChart from "@/components/charts/MiniChart";
import {
  GRID_COLS,
  clampTile,
  defaultSize,
  freeSpot,
  kpiValue,
  readingOrder,
  rowsUsed,
  settle,
  suggestDashboard,
  tileChart,
  type CellRef,
} from "@/lib/dashboard";
import { newId, useStore } from "@/lib/store";
import type { DashboardSpec, DashboardTile, WorkbookSpec } from "@/lib/types";
import { useUi } from "@/lib/ui";
import { chartData } from "./chartData";
import TileForm, { type TileDraft } from "./DashboardEditor";
import KpiTile from "./KpiTile";

const ROW_H = 72;
const GAP = 12;
const NARROW = 640;

/** Apply a change to the workbook's dashboard as one undoable step. */
function editDashboard(fn: (d: DashboardSpec) => void) {
  useStore.getState().mutate((wb) => {
    wb.dashboard ??= { title: `${wb.title} dashboard`, tiles: [] };
    fn(wb.dashboard);
  });
}

function jumpTo(at: CellRef) {
  const st = useStore.getState();
  useUi.setState({ view: "sheet", mobilePane: "sheet" });
  st.setActive(at.sheet);
  st.select({ r1: at.r, c1: at.c, r2: at.r, c2: at.c });
}

interface Drag {
  id: string;
  mode: "move" | "resize";
  sx: number;
  sy: number;
  orig: DashboardTile;
  cur: DashboardTile;
}

/** The workbook's dashboard page: KPI, chart and text tiles on a 12-column grid. */
export default function DashboardView() {
  const workbook = useStore((s) => s.workbook);
  const dashboard = workbook.dashboard;
  const tiles = dashboard?.tiles ?? [];
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState<null | { tile?: DashboardTile }>(null);
  const [drag, setDrag] = useState<Drag | null>(null);
  const [announce, setAnnounce] = useState("");
  const sectionRef = useRef<HTMLElement>(null);
  const [width, setWidth] = useState(1024);

  // The grid stacks to one column when the page area (not the window) is narrow, e.g. with a wide chat panel.
  useEffect(() => {
    const el = sectionRef.current;
    if (!el) return;
    const measure = () => {
      const cs = getComputedStyle(el);
      setWidth(el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight));
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const narrow = width < NARROW;
  const colW = (width - GAP * (GRID_COLS - 1)) / GRID_COLS;

  // While dragging, show the dragged tile at its new spot and the others making room.
  const shown = drag ? settle(tiles.map((t) => (t.id === drag.id ? drag.cur : t)), drag.id) : tiles;
  const ordered = narrow ? readingOrder(shown) : shown;

  const suggest = () => {
    const had = tiles.length > 0;
    const next = suggestDashboard(useStore.getState().workbook, newId);
    useStore.getState().mutate((wb) => {
      wb.dashboard = next;
    });
    useUi.getState().notify(had ? "Replaced the dashboard with a suggested one. Undo (Ctrl+Z) brings yours back." : "Built a dashboard from your totals and charts.");
  };

  const placeTile = (t: DashboardTile, label: string) => {
    editDashboard((d) => {
      d.tiles = settle(d.tiles.map((x) => (x.id === t.id ? t : x)), t.id);
    });
    setAnnounce(label);
  };

  const removeTile = (t: DashboardTile) => {
    editDashboard((d) => {
      d.tiles = d.tiles.filter((x) => x.id !== t.id);
    });
    setAnnounce(`Removed ${t.title || "tile"}.`);
  };

  const saveForm = (draft: TileDraft) => {
    const editingTile = form?.tile;
    editDashboard((d) => {
      if (editingTile) {
        d.tiles = d.tiles.map((x) => (x.id === editingTile.id ? { ...x, ...draft } : x));
      } else {
        const size = defaultSize(draft.kind);
        const spot = freeSpot(d.tiles, size.w, size.h);
        d.tiles.push({ id: newId(), ...draft, ...size, ...spot });
      }
    });
    setForm(null);
    setEditing(true);
  };

  // ---- Pointer drag (move by the grip, resize by the corner) ----
  const startDrag = (e: ReactPointerEvent<HTMLElement>, tile: DashboardTile, mode: Drag["mode"]) => {
    if (narrow || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    setDrag({ id: tile.id, mode, sx: e.clientX, sy: e.clientY, orig: tile, cur: tile });
  };
  const moveDrag = (e: ReactPointerEvent<HTMLElement>) => {
    if (!drag) return;
    const dx = Math.round((e.clientX - drag.sx) / (colW + GAP));
    const dy = Math.round((e.clientY - drag.sy) / (ROW_H + GAP));
    const o = drag.orig;
    const cur = clampTile(drag.mode === "move" ? { ...o, x: o.x + dx, y: o.y + dy } : { ...o, w: o.w + dx, h: o.h + dy });
    if (cur.x !== drag.cur.x || cur.y !== drag.cur.y || cur.w !== drag.cur.w || cur.h !== drag.cur.h) setDrag({ ...drag, cur });
  };
  const endDrag = (commit: boolean) => {
    if (!drag) return;
    const { orig, cur } = drag;
    setDrag(null);
    if (commit && (cur.x !== orig.x || cur.y !== orig.y || cur.w !== orig.w || cur.h !== orig.h))
      placeTile(cur, describe(cur, drag.mode));
  };

  // ---- Keyboard: arrows move, Shift + arrows resize, Delete removes ----
  const onTileKey = (e: ReactKeyboardEvent<HTMLElement>, tile: DashboardTile) => {
    if (e.target !== e.currentTarget) return;
    const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (d) {
      e.preventDefault();
      const next = clampTile(e.shiftKey ? { ...tile, w: tile.w + d[0], h: tile.h + d[1] } : { ...tile, x: tile.x + d[0], y: tile.y + d[1] });
      if (next.x !== tile.x || next.y !== tile.y || next.w !== tile.w || next.h !== tile.h) placeTile(next, describe(next, e.shiftKey ? "resize" : "move"));
    } else if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      removeTile(tile);
    } else if (e.key === "Enter") {
      e.preventDefault();
      setForm({ tile });
    }
  };

  const toolbarBtn = "flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-sm text-slate-700 hover:bg-slate-100";

  return (
    <section ref={sectionRef} aria-label="Dashboard" className="min-h-0 flex-1 overflow-y-auto bg-slate-50 p-4 md:p-6">
      <div className="mb-4 flex flex-wrap items-center gap-2">
        <LayoutDashboard className="h-5 w-5 text-brand" aria-hidden="true" />
        {editing && dashboard ? (
          <input
            key={dashboard.title}
            aria-label="Dashboard title"
            className="min-w-0 flex-1 rounded-md border border-slate-300 bg-white px-2 py-1 text-lg font-semibold text-slate-800 focus:border-brand focus:outline-none"
            defaultValue={dashboard.title}
            onBlur={(e) => {
              const v = e.target.value.trim();
              if (v && v !== dashboard.title) editDashboard((d) => void (d.title = v));
            }}
            onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
          />
        ) : (
          <h2 className="min-w-0 flex-1 truncate text-lg font-semibold text-slate-800">{dashboard?.title || "Dashboard"}</h2>
        )}
        <button type="button" className={toolbarBtn} onClick={suggest} title="Build a dashboard from your total rows and charts">
          <Sparkles className="h-4 w-4 text-brand" aria-hidden="true" /> Suggest a dashboard
        </button>
        {tiles.length > 0 && (
          <>
            <button type="button" className={toolbarBtn} onClick={() => setForm({})}>
              <Plus className="h-4 w-4" aria-hidden="true" /> Add tile
            </button>
            <button
              type="button"
              aria-pressed={editing}
              className={editing ? "flex items-center gap-1.5 rounded-md bg-brand px-3 py-1.5 text-sm font-medium text-white hover:bg-brand-dark" : toolbarBtn}
              onClick={() => {
                setEditing(!editing);
                setForm(null);
              }}
            >
              {editing ? <Check className="h-4 w-4" aria-hidden="true" /> : <Pencil className="h-4 w-4" aria-hidden="true" />}
              {editing ? "Done" : "Edit"}
            </button>
          </>
        )}
      </div>

      {form && <TileForm key={form.tile?.id ?? "new"} initial={form.tile} onSave={saveForm} onCancel={() => setForm(null)} />}

      {editing && tiles.length > 0 && (
        <p className="mb-3 text-xs text-slate-500">
          {narrow
            ? "Focus a tile and use the arrow keys to move it, Shift + arrows to resize. Enter edits it, Delete removes it."
            : "Drag a tile by its grip to move it and by its corner to resize. Or focus a tile: arrows move, Shift + arrows resize, Enter edits, Delete removes."}
        </p>
      )}
      <div aria-live="polite" className="sr-only">
        {announce}
      </div>

      {tiles.length === 0 && !form ? (
        <EmptyState onSuggest={suggest} onAdd={() => setForm({})} />
      ) : (
        <div
          className={narrow ? "flex flex-col gap-3" : "grid"}
          style={narrow ? undefined : { gridTemplateColumns: `repeat(${GRID_COLS}, minmax(0, 1fr))`, gridAutoRows: ROW_H, gap: GAP, minHeight: (rowsUsed(shown) + (drag ? 2 : 0)) * (ROW_H + GAP) }}
        >
          {ordered.map((tile) => (
            <Tile
              key={tile.id}
              tile={tile}
              workbook={workbook}
              editing={editing}
              narrow={narrow}
              dragging={drag?.id === tile.id}
              onKeyDown={(e) => onTileKey(e, tile)}
              onGrip={(e) => startDrag(e, tile, "move")}
              onCorner={(e) => startDrag(e, tile, "resize")}
              onPointerMove={moveDrag}
              onPointerUp={() => endDrag(true)}
              onPointerCancel={() => endDrag(false)}
              onRename={(title) => editDashboard((d) => void (d.tiles = d.tiles.map((x) => (x.id === tile.id ? { ...x, title } : x))))}
              onSettings={() => setForm({ tile })}
              onRemove={() => removeTile(tile)}
            />
          ))}
        </div>
      )}
    </section>
  );
}

function describe(t: DashboardTile, mode: "move" | "resize") {
  return mode === "move"
    ? `${t.title || "Tile"} moved to column ${t.x + 1}, row ${t.y + 1}.`
    : `${t.title || "Tile"} resized to ${t.w} columns by ${t.h} rows.`;
}

function EmptyState({ onSuggest, onAdd }: { onSuggest: () => void; onAdd: () => void }) {
  return (
    <div className="mx-auto mt-8 max-w-md rounded-xl border border-dashed border-slate-300 bg-white p-6 text-center">
      <LayoutDashboard className="mx-auto mb-3 h-8 w-8 text-slate-400" aria-hidden="true" />
      <h3 className="mb-1 text-base font-semibold text-slate-800">No dashboard yet</h3>
      <p className="mb-4 text-sm text-slate-500">
        A dashboard puts your key numbers and charts on one page. Its numbers stay live as the sheets change, and the Excel download gets a Dashboard
        sheet too.
      </p>
      <div className="flex flex-wrap justify-center gap-2">
        <button type="button" onClick={onSuggest} className="flex items-center gap-1.5 rounded-md bg-brand px-4 py-2 text-sm font-medium text-white hover:bg-brand-dark">
          <Sparkles className="h-4 w-4" aria-hidden="true" /> Suggest a dashboard
        </button>
        <button type="button" onClick={onAdd} className="flex items-center gap-1.5 rounded-md border border-slate-300 px-4 py-2 text-sm text-slate-700 hover:bg-slate-100">
          <Plus className="h-4 w-4" aria-hidden="true" /> Add a tile
        </button>
      </div>
      <p className="mt-3 text-xs text-slate-400">Suggest uses your Total rows as headline numbers and adds every chart.</p>
    </div>
  );
}

function Tile({
  tile,
  workbook,
  editing,
  narrow,
  dragging,
  onKeyDown,
  onGrip,
  onCorner,
  onPointerMove,
  onPointerUp,
  onPointerCancel,
  onRename,
  onSettings,
  onRemove,
}: {
  tile: DashboardTile;
  workbook: WorkbookSpec;
  editing: boolean;
  narrow: boolean;
  dragging: boolean;
  onKeyDown: (e: ReactKeyboardEvent<HTMLElement>) => void;
  onGrip: (e: ReactPointerEvent<HTMLElement>) => void;
  onCorner: (e: ReactPointerEvent<HTMLElement>) => void;
  onPointerMove: (e: ReactPointerEvent<HTMLElement>) => void;
  onPointerUp: () => void;
  onPointerCancel: () => void;
  onRename: (title: string) => void;
  onSettings: () => void;
  onRemove: () => void;
}) {
  const kpi = tile.kind === "kpi" ? kpiValue(workbook, tile) : null;
  const chart = tile.kind === "chart" ? tileChart(workbook, tile) : null;
  const place = narrow
    ? { minHeight: tile.kind === "chart" ? 240 : tile.h * ROW_H }
    : { gridColumn: `${tile.x + 1} / span ${tile.w}`, gridRow: `${tile.y + 1} / span ${tile.h}` };
  const clickable = !editing && !!kpi?.at;
  const dragHandlers = { onPointerMove, onPointerUp, onPointerCancel };

  const body = (
    <>
      {kpi && <KpiTile value={kpi} compact={tile.h < 2} />}
      {tile.kind === "chart" &&
        (chart ? (
          <div className="flex min-h-0 flex-1 items-center justify-center [&>svg]:h-full [&>svg]:max-h-full [&>svg]:w-full">
            <MiniChart data={chartData(workbook, chart.sheetIndex, workbook.sheets[chart.sheetIndex], chart.chart)} />
          </div>
        ) : (
          <p className="text-sm text-slate-500">This chart was removed from {tile.sheet ?? "its sheet"}.</p>
        ))}
      {tile.kind === "text" && <p className="min-h-0 flex-1 overflow-y-auto text-sm whitespace-pre-wrap text-slate-600">{tile.text}</p>}
    </>
  );

  const label = `${tile.kind === "kpi" ? "Number" : tile.kind === "chart" ? "Chart" : "Text"} tile: ${tile.title || "untitled"}`;

  return (
    <article
      aria-label={editing ? `${label}. Column ${tile.x + 1}, row ${tile.y + 1}, ${tile.w} by ${tile.h}.` : label}
      aria-roledescription={editing ? "movable tile" : undefined}
      tabIndex={editing ? 0 : undefined}
      onKeyDown={editing ? onKeyDown : undefined}
      style={place}
      className={`group relative flex min-w-0 flex-col overflow-hidden rounded-xl border bg-white p-3 shadow-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-brand ${
        dragging ? "z-10 border-brand opacity-90 shadow-lg" : "border-slate-200"
      } ${editing ? "ring-offset-1" : ""}`}
    >
      <div className="mb-1 flex items-center gap-1">
        {editing && !narrow && (
          <span
            role="presentation"
            title="Drag to move"
            className="-ml-1 cursor-grab touch-none rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700 active:cursor-grabbing"
            onPointerDown={onGrip}
            {...dragHandlers}
          >
            <GripVertical className="h-4 w-4" />
          </span>
        )}
        {editing ? (
          <input
            key={tile.title}
            aria-label="Tile title"
            className="min-w-0 flex-1 rounded border border-transparent bg-transparent px-1 text-sm font-medium text-slate-700 hover:border-slate-200 focus:border-brand focus:outline-none"
            defaultValue={tile.title}
            onBlur={(e) => e.target.value.trim() !== tile.title && onRename(e.target.value.trim())}
            onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
          />
        ) : (
          <h3 className="min-w-0 flex-1 truncate text-sm font-medium text-slate-600" title={tile.title}>
            {tile.title}
          </h3>
        )}
        {editing && (
          <>
            <button type="button" aria-label={`Edit ${tile.title || "tile"}`} title="Edit tile" className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700" onClick={onSettings}>
              <Settings2 className="h-3.5 w-3.5" />
            </button>
            <button type="button" aria-label={`Remove ${tile.title || "tile"}`} title="Remove tile" className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-red-600" onClick={onRemove}>
              <Trash2 className="h-3.5 w-3.5" />
            </button>
          </>
        )}
      </div>
      {clickable ? (
        <button
          type="button"
          className="-m-1 flex min-h-0 flex-1 flex-col rounded-lg p-1 text-left hover:bg-slate-50 focus:outline-none focus-visible:ring-2 focus-visible:ring-brand"
          title={`Go to ${tile.ref}`}
          aria-label={`${tile.title}: ${kpi!.text}. Go to ${tile.ref}`}
          onClick={() => jumpTo(kpi!.at!)}
        >
          {body}
        </button>
      ) : (
        body
      )}
      {editing && !narrow && (
        <span
          role="presentation"
          title="Drag to resize"
          className="absolute right-0 bottom-0 h-4 w-4 cursor-se-resize touch-none"
          onPointerDown={onCorner}
          {...dragHandlers}
        >
          <svg viewBox="0 0 16 16" className="h-4 w-4 text-slate-400" aria-hidden="true">
            <path d="M14 6v8H6M14 10v4h-4" fill="none" stroke="currentColor" strokeWidth="1.5" />
          </svg>
        </span>
      )}
    </article>
  );
}
