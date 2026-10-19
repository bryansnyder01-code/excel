"use client";

import { useEffect, useRef, useState } from "react";
import {
  ChevronDown,
  Cloud,
  Download,
  FilePlus2,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Keyboard,
  LayoutTemplate,
  Loader2,
  MessageSquare,
  Moon,
  MoreHorizontal,
  Share2,
  Sheet,
  Sun,
  Table2,
  Upload,
  UserRound, Command, } from "lucide-react";
import ChartsPanel from "./charts/ChartsPanel";
import ChatPanel, { ACCEPT } from "./chat/ChatPanel";
import AccountDialog from "./dialogs/AccountDialog";
import ChartDialog from "./dialogs/ChartDialog";
import CommandPalette from "./CommandPalette";
import DashboardView from "./dashboard/DashboardView";
import AiFillDialog from "./dialogs/AiFillDialog";
import BillingDialog from "./dialogs/BillingDialog";
import ConnectionsDialog from "./dialogs/ConnectionsDialog";
import ImportImageDialog from "./dialogs/ImportImageDialog";
import ReportsDialog from "./dialogs/ReportsDialog";
import CompareDialog from "./dialogs/CompareDialog";
import CellHistoryPanel from "./dialogs/CellHistoryPanel";
import CheckPanel from "./dialogs/CheckPanel";
import CommentsPanel from "./dialogs/CommentsPanel";
import ProtectDialog from "./dialogs/ProtectDialog";
import CleanupDialog from "./dialogs/CleanupDialog";
import SaveTemplateDialog from "./dialogs/SaveTemplateDialog";
import WhatIfDialog from "./dialogs/WhatIfDialog";
import ConditionalDialog from "./dialogs/ConditionalDialog";
import DocsDialog from "./dialogs/DocsDialog";
import DropdownDialog from "./dialogs/DropdownDialog";
import PivotDialog from "./dialogs/PivotDialog";
import ShareDialog from "./dialogs/ShareDialog";
import ShortcutsDialog from "./dialogs/ShortcutsDialog";
import TemplatesDialog from "./dialogs/TemplatesDialog";
import FindBar from "./grid/FindBar";
import FormatToolbar from "./grid/FormatToolbar";
import ReviewBar from "./grid/ReviewBar";
import FormulaBar from "./grid/FormulaBar";
import SheetTabs from "./grid/SheetTabs";
import SpreadsheetGrid from "./grid/SpreadsheetGrid";
import { restoreSession, schedulePush, setPendingReset, setViewOnly, useAccount, whenPulled } from "@/lib/account";
import { startConnectionTimer } from "@/lib/connections";
import { startLive, useLive } from "@/lib/live";
import { useShallow } from "zustand/react/shallow";
import { uploadFile } from "@/lib/actions";
import { cloud, downloadWorkbook, getCapabilities, saveBlob } from "@/lib/api";
import { sheetToCsv } from "@/lib/csv";
import { GOOGLE_CLIENT_ID, openInGoogleSheets } from "@/lib/google";
import { getDoc, lastDocId, putDoc, saveDoc } from "@/lib/persist";
import type { SavedDoc } from "@/lib/types";
import { isBlank, useStore } from "@/lib/store";
import { applyTheme } from "@/lib/theme";
import { useUi } from "@/lib/ui";

const CHAT_W_KEY = "sheetcraft:chatWidth";

export default function Workspace() {
  useAutosave();
  useSharedLink();
  const dialog = useUi((s) => s.dialog);
  const toast = useUi((s) => s.toast);
  const pane = useUi((s) => s.mobilePane);
  const view = useUi((s) => s.view);
  const [dragging, setDragging] = useState(false);
  const chatW = useChatWidth();

  // Shortcuts that work wherever focus is.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        const u = useUi.getState();
        if (u.dialog === "palette") u.close();
        else u.open("palette");
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === "/" || e.key === "?")) {
        e.preventDefault();
        useUi.getState().open("shortcuts");
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  return (
    <div
      className="flex h-dvh flex-col"
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes("Files")) {
          e.preventDefault();
          setDragging(true);
        }
      }}
      onDragLeave={(e) => e.currentTarget === e.target && setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        const f = e.dataTransfer.files?.[0];
        if (f) uploadFile(f);
      }}
    >
      <a
        href="#sheet"
        className="sr-only z-[80] rounded bg-brand px-3 py-2 text-sm text-white focus:not-sr-only focus:fixed focus:top-2 focus:left-2"
        onClick={(e) => {
          e.preventDefault();
          useUi.setState({ mobilePane: "sheet" });
          document.querySelector<HTMLElement>(".sc-grid")?.focus();
        }}
      >
        Skip to the spreadsheet
      </a>
      <Header />
      <div className="flex min-h-0 flex-1">
        <aside
          className={`shrink-0 flex-col border-r border-slate-200 bg-white max-md:w-full md:flex md:w-[var(--chat-w)] ${pane === "chat" ? "flex" : "hidden"}`}
          style={{ ["--chat-w" as string]: `${chatW.width}px` }}
        >
          <ChatPanel />
        </aside>
        <div
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize chat panel"
          title="Drag to resize"
          className="hidden w-1 shrink-0 cursor-col-resize bg-transparent hover:bg-brand/40 md:block"
          style={{ marginLeft: -2, marginRight: -2, zIndex: 30 }}
          onMouseDown={chatW.start}
          onDoubleClick={chatW.reset}
        />
        <main className={`min-w-0 flex-1 flex-col bg-white md:flex ${pane === "sheet" ? "flex" : "hidden"}`}>
          <ReviewBar />
          {view === "dashboard" ? (
            <DashboardView />
          ) : (
            <>
              <FormatToolbar />
              <FindBar />
              <FormulaBar />
              <SpreadsheetGrid />
              <ChartsPanel />
            </>
          )}
          <SheetTabs />
        </main>
      </div>
      <MobileTabs />

      {dialog === "templates" && <TemplatesDialog />}
      {dialog === "docs" && <DocsDialog />}
      {dialog === "pivot" && <PivotDialog />}
      {dialog === "dropdown" && <DropdownDialog />}
      {dialog === "chart" && <ChartDialog />}
      {dialog === "shortcuts" && <ShortcutsDialog />}
      {dialog === "conditional" && <ConditionalDialog />}
      {dialog === "account" && <AccountDialog />}
      {dialog === "share" && <ShareDialog />}
      {dialog === "check" && <CheckPanel />}
      {dialog === "cleanup" && <CleanupDialog />}
      {dialog === "whatif" && <WhatIfDialog />}
      {dialog === "saveTemplate" && <SaveTemplateDialog />}
      {dialog === "protect" && <ProtectDialog />}
      {dialog === "palette" && <CommandPalette />}
      {dialog === "aiFill" && <AiFillDialog />}
      {dialog === "importImage" && <ImportImageDialog />}
      {dialog === "connections" && <ConnectionsDialog />}
      {dialog === "reports" && <ReportsDialog />}
      {dialog === "compare" && <CompareDialog />}
      {dialog === "comments" && <CommentsPanel />}
      {dialog === "cellHistory" && <CellHistoryPanel />}
      {dialog === "billing" && <BillingDialog />}

      {dragging && (
        <div className="pointer-events-none fixed inset-0 z-50 flex items-center justify-center bg-brand/10 ring-4 ring-brand ring-inset">
          <div className="rounded-xl bg-white px-6 py-4 text-sm font-medium text-brand shadow-lg">Drop an .xlsx or .csv file to open it</div>
        </div>
      )}
      {toast && (
        <div
          role="status"
          className={`fixed bottom-14 left-1/2 z-[70] max-w-[90vw] -translate-x-1/2 rounded-lg px-4 py-2.5 text-sm text-white shadow-lg ${
            toast.error ? "bg-[#dc2626]" : "bg-[#1f2937] ring-1 ring-white/10"
          }`}
        >
          {toast.text}
        </div>
      )}
    </div>
  );
}

/** Phone layout: switch between the chat and the sheet. */
function MobileTabs() {
  const pane = useUi((s) => s.mobilePane);
  const busy = useStore((s) => s.busy);
  const tab = (value: "chat" | "sheet", label: string, icon: React.ReactNode) => (
    <button
      role="tab"
      aria-selected={pane === value}
      onClick={() => useUi.setState({ mobilePane: value })}
      className={`flex flex-1 flex-col items-center gap-0.5 py-2 text-xs [&_svg]:h-5 [&_svg]:w-5 ${pane === value ? "font-semibold text-brand" : "text-slate-500"}`}
    >
      {icon}
      {label}
    </button>
  );
  return (
    <nav role="tablist" aria-label="Show" className="flex shrink-0 border-t border-slate-200 bg-white pb-[env(safe-area-inset-bottom)] md:hidden">
      {tab("chat", busy ? "Chat (working)" : "Chat", <MessageSquare />)}
      {tab("sheet", "Sheet", <Table2 />)}
    </nav>
  );
}

/** The chat panel's width is draggable and remembered. */
function useChatWidth() {
  const [width, setWidth] = useState(380);
  useEffect(() => {
    try {
      const w = Number(localStorage.getItem(CHAT_W_KEY));
      if (w >= 280 && w <= 720) setWidth(w);
    } catch {}
  }, []);
  const save = (w: number) => {
    try {
      localStorage.setItem(CHAT_W_KEY, String(w));
    } catch {}
  };
  return {
    width,
    reset: () => {
      setWidth(380);
      save(380);
    },
    start: (e: React.MouseEvent) => {
      e.preventDefault();
      const x0 = e.clientX;
      const w0 = width;
      let w = w0;
      const move = (ev: MouseEvent) => {
        w = Math.max(280, Math.min(720, w0 + ev.clientX - x0, window.innerWidth - 360));
        setWidth(w);
      };
      const up = () => {
        save(w);
        window.removeEventListener("mousemove", move);
        window.removeEventListener("mouseup", up);
        document.body.style.cursor = "";
      };
      document.body.style.cursor = "col-resize";
      window.addEventListener("mousemove", move);
      window.addEventListener("mouseup", up);
    },
  };
}

const INVITE_KEY = "sheetcraft:invite";

/** An invite link (?invite=...) adds the workbook to the visitor's account to edit together; it needs a sign-in first. */
async function acceptPendingInvite() {
  let token: string | null = null;
  try {
    token = sessionStorage.getItem(INVITE_KEY);
  } catch {}
  if (!token) return;
  if (!useAccount.getState().user) {
    useUi.getState().notify("Sign in or create an account to open the workbook you were invited to.");
    useUi.getState().open("account");
    const unsub = useAccount.subscribe((s) => {
      if (s.user) {
        unsub();
        acceptPendingInvite();
      }
    });
    return;
  }
  try {
    sessionStorage.removeItem(INVITE_KEY);
  } catch {}
  try {
    const { id, title, role } = await cloud.accept(token);
    setViewOnly(id, role === "viewer");
    const remote = await cloud.get(id);
    putDoc({ ...(remote.data as SavedDoc), id, title: remote.title, updated: remote.updated });
    const doc = getDoc(id);
    if (doc) useStore.getState().loadDocument(doc);
    useUi.setState({ mobilePane: "sheet" });
    // Let the sign-in's own "Account created" message go first, so this one stays on screen.
    for (let i = 0; i < 50 && useAccount.getState().syncing; i++) await new Promise((r) => setTimeout(r, 200));
    setTimeout(() => useUi.getState().notify(
          role === "viewer"
            ? `You can now view "${title}" and comment. Changes appear as others make them.`
            : `You can now edit "${title}" together. Changes appear for everyone as you type.`,
        ),
      300,
    );
  } catch (e) {
    useUi.getState().notify(e instanceof Error ? e.message : "That invite link didn't work.", true);
  }
}

/** A link from an email (?doc=...&cell=Sheet!B5): open that workbook at that cell, with its comments. */
async function openDocLink(id: string, cell: string | null) {
  const open = () => {
    const doc = getDoc(id);
    if (!doc) return false;
    const st = useStore.getState();
    if (st.docId !== id) st.loadDocument(doc);
    const m = cell && /^(.*)!\$?([A-Z]+)\$?(\d+)$/.exec(cell);
    if (m) {
      const i = useStore.getState().workbook.sheets.findIndex((s) => s.name === m[1].replace(/^'|'$/g, ""));
      const c = [...m[2]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const r = Number(m[3]) - 1;
      if (i >= 0) {
        useStore.getState().setActive(i);
        useStore.getState().select({ r1: r, c1: c, r2: r, c2: c });
        useUi.getState().open("comments");
      }
    }
    useUi.setState({ mobilePane: "sheet", view: "sheet" });
    return true;
  };
  if (open()) return;
  // Not in this browser yet: it arrives with the first sync after sign-in.
  const unsub = useAccount.subscribe((s, prev) => {
    if (prev.syncing && !s.syncing && open()) unsub();
  });
}

/** Opening a share link (?share=...) makes the visitor their own copy of that workbook. */
function useSharedLink() {
  useEffect(() => {
    const url = new URL(window.location.href);
    const reset = url.searchParams.get("reset");
    const invite = url.searchParams.get("invite");
    const doc = url.searchParams.get("doc");
    const cell = url.searchParams.get("cell");
    const billing = url.searchParams.get("billing");
    for (const k of ["reset", "invite", "doc", "cell", "billing"]) url.searchParams.delete(k);
    if (reset || invite || doc || billing) window.history.replaceState(null, "", url.toString());
    if (billing === "success") useUi.getState().notify("Thanks! Your Pro plan is active.");
    if (doc) openDocLink(doc, cell);
    if (reset) {
      setPendingReset(reset);
      useUi.getState().open("account");
    }
    if (invite) {
      try {
        sessionStorage.setItem(INVITE_KEY, invite);
      } catch {}
    }
    const token = url.searchParams.get("share");
    if (!token) return;
    url.searchParams.delete("share");
    window.history.replaceState(null, "", url.toString());
    (async () => {
      try {
        const { title, workbook } = await cloud.shared(token);
        const st = useStore.getState();
        if (!isBlank(st.workbook) || st.messages.length) st.newWorkbook();
        useStore.getState().applyResult({ ...workbook, title }, "Shared copy", {
          text: `This is your own copy of "${title}", shared with you. Changes you make stay in your copy.`,
        });
        useUi.setState({ mobilePane: "sheet" });
      } catch (e) {
        useUi.getState().notify(e instanceof Error ? e.message : "That share link didn't open.", true);
      }
    })();
  }, []);
}

/** Save the open workbook to this browser now. Skips a brand-new, untouched workbook. */
export function saveCurrent(st: ReturnType<typeof useStore.getState> = useStore.getState()) {
  if (isBlank(st.workbook) && !st.messages.length) return;
  const ok = saveDoc({
    id: st.docId,
    title: st.workbook.title,
    updated: Date.now(),
    workbook: st.workbook,
    messages: st.messages,
    checkpoints: st.checkpoints,
  });
  if (!ok) useUi.getState().notify("This browser's storage is full, so the latest changes weren't saved.", true);
  else schedulePush(st.docId);
}

/** Reopen the last workbook on load, then save every change to this browser. */
function useAutosave() {
  useEffect(() => {
    const id = lastDocId();
    const doc = id ? getDoc(id) : undefined;
    if (doc) useStore.getState().loadDocument(doc);
    // A sync can bring a newer copy of the open workbook from another device.
    whenPulled((ids) => {
      const st = useStore.getState();
      const fresh = ids.includes(st.docId) ? getDoc(st.docId) : undefined;
      if (fresh && !st.busy) st.loadDocument(fresh);
    });
    restoreSession().then(() => {
      startLive();
      acceptPendingInvite();
    });

    let timer: ReturnType<typeof setTimeout> | undefined;
    const unsub = useStore.subscribe((s, prev) => {
      if (s.docId !== prev.docId) {
        // Switching documents: save the one we're leaving with its latest changes.
        clearTimeout(timer);
        saveCurrent(prev);
        return;
      }
      if (s.workbook === prev.workbook && s.messages === prev.messages && s.checkpoints === prev.checkpoints) return;
      clearTimeout(timer);
      timer = setTimeout(() => saveCurrent(), 500);
    });
    const flush = () => {
      clearTimeout(timer);
      saveCurrent();
    };
    window.addEventListener("beforeunload", flush);
    const stopConnections = startConnectionTimer();
    return () => {
      stopConnections();
      unsub();
      clearTimeout(timer);
      window.removeEventListener("beforeunload", flush);
    };
  }, []);
}

function Header() {
  const title = useStore((s) => s.workbook.title);
  const setTitle = useStore((s) => s.setTitle);
  const fileRef = useRef<HTMLInputElement>(null);
  const open = useUi((s) => s.open);
  const user = useAccount((s) => s.user);
  const syncing = useAccount((s) => s.syncing);
  const [more, setMore] = useState(false);
  const [dark, setDark] = useState(false);
  const dialogOpen = useUi((s) => s.dialog);
  // The command palette can switch the theme too, so re-read it whenever a dialog closes.
  useEffect(() => setDark(document.documentElement.classList.contains("dark")), [dialogOpen]);
  useEffect(() => {
    if (!more) return;
    const close = () => setMore(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [more]);

  const toggleTheme = () => {
    applyTheme(dark ? "light" : "dark");
    setDark(!dark);
  };

  const items: { icon: React.ReactNode; label: string; onClick: () => void }[] = [
    { icon: <FilePlus2 />, label: "New", onClick: () => useStore.getState().newWorkbook() },
    { icon: <LayoutTemplate />, label: "Templates", onClick: () => open("templates") },
    { icon: <FolderOpen />, label: "My workbooks", onClick: () => open("docs") },
    { icon: <Upload />, label: "Upload", onClick: () => fileRef.current?.click() },
    { icon: <Share2 />, label: "Share", onClick: () => open("share") },
  ];

  return (
    <header className="flex h-14 shrink-0 items-center gap-2 border-b border-slate-200 bg-white px-2 md:gap-3 md:px-4">
      <div className="flex items-center gap-2 font-semibold text-brand">
        <Sheet className="h-6 w-6" />
        <span className="text-lg whitespace-nowrap max-xl:hidden">SheetCraft AI</span>
      </div>
      <input
        aria-label="Workbook title"
        className="w-28 min-w-0 rounded-md border border-transparent bg-transparent px-2 py-1 text-sm text-slate-700 hover:border-slate-200 focus:border-brand focus:outline-none sm:w-48 xl:w-64"
        defaultValue={title}
        key={title}
        onBlur={(e) => e.target.value !== title && setTitle(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && (e.target as HTMLInputElement).blur()}
      />
      <LivePeople />
      <div className="ml-auto flex items-center gap-0.5 md:gap-1">
        {items.map((it) => (
          <HeaderButton key={it.label} icon={it.icon} onClick={it.onClick} className="max-xl:hidden">
            {it.label}
          </HeaderButton>
        ))}
        <div className="relative xl:hidden">
          <HeaderButton
            icon={<MoreHorizontal />}
            onClick={() => setTimeout(() => setMore(!more))}
            label="More"
          />
          {more && (
            <div role="menu" className="absolute top-full right-0 z-50 mt-1 w-48 rounded-lg border border-slate-200 bg-white py-1 shadow-lg">
              {items.map((it) => (
                <MenuItem key={it.label} icon={it.icon} label={it.label} onClick={() => (setMore(false), it.onClick())} />
              ))}
              <MenuItem icon={<Keyboard />} label="Keyboard shortcuts" onClick={() => (setMore(false), open("shortcuts"))} />
            </div>
          )}
        </div>
        <input
          ref={fileRef}
          type="file"
          accept={ACCEPT}
          aria-label="Upload a spreadsheet"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) uploadFile(f);
            e.target.value = "";
          }}
        />
        <HeaderButton icon={<Command />} onClick={() => open("palette")} label="Search actions (Ctrl+K)" />
        <HeaderButton icon={<Keyboard />} onClick={() => open("shortcuts")} label="Keyboard shortcuts (Ctrl+/)" className="max-xl:hidden" />
        <HeaderButton icon={dark ? <Sun /> : <Moon />} onClick={toggleTheme} label={dark ? "Light mode" : "Dark mode"} />
        <HeaderButton
          icon={user ? <Cloud className={syncing ? "animate-pulse" : ""} /> : <UserRound />}
          onClick={() => open("account")}
          label={user ? `Signed in as ${user.email}` : "Sign in"}
        >
          <span className="max-lg:hidden">{user ? user.email.split("@")[0] : "Sign in"}</span>
        </HeaderButton>
        <ExportButton />
      </div>
    </header>
  );
}

/** Who else is editing this workbook right now, as coloured initials. */
function LivePeople() {
  const { status, peers, role } = useLive(useShallow((s) => ({ status: s.status, peers: s.peers, role: s.role })));
  const people = Object.values(peers);
  if (status === "off") return null;
  return (
    <div className="flex items-center gap-1.5 max-sm:hidden" aria-label="People editing now">
      {status === "offline" && <span className="text-xs text-amber-600">Reconnecting...</span>}
      {status === "live" && role === "viewer" && (
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600" title="You can look and comment, but not change cells">
          View only
        </span>
      )}
      {people.map((p) => (
        <span
          key={p.id}
          title={`${p.email} is editing`}
          className="flex h-7 w-7 items-center justify-center rounded-full text-xs font-semibold text-white ring-2 ring-white"
          style={{ background: p.color }}
        >
          {p.email[0]?.toUpperCase()}
        </span>
      ))}
      {status === "live" && people.length > 0 && <span className="whitespace-nowrap text-xs text-slate-500 max-2xl:hidden">editing with you</span>}
    </div>
  );
}

function HeaderButton({
  icon,
  onClick,
  children,
  label,
  className = "",
}: {
  icon: React.ReactNode;
  onClick: () => void;
  children?: React.ReactNode;
  label?: string;
  className?: string;
}) {
  return (
    <button
      title={label}
      aria-label={label}
      className={`flex items-center gap-1.5 rounded-md px-2 py-2 text-sm whitespace-nowrap text-slate-600 hover:bg-slate-100 [&_svg]:h-4 [&_svg]:w-4 ${className}`}
      onClick={onClick}
    >
      {icon}
      {children}
    </button>
  );
}

function ExportButton() {
  const [busy, setBusy] = useState<string | null>(null);
  const [menu, setMenu] = useState(false);
  const [pdf, setPdf] = useState(true);
  const notify = useUi((s) => s.notify);

  useEffect(() => {
    getCapabilities()
      .then((c) => setPdf(c.pdf))
      .catch(() => {});
  }, []);
  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [menu]);

  const run = async (kind: string, fn: () => Promise<void>) => {
    setMenu(false);
    setBusy(kind);
    try {
      await fn();
    } catch (e) {
      notify(e instanceof Error ? e.message : "Export failed.", true);
    } finally {
      setBusy(null);
    }
  };

  const wb = () => useStore.getState().workbook;

  return (
    <div className="relative ml-1 flex md:ml-2">
      <button
        aria-label="Download Excel Workbook"
        className="flex items-center gap-2 rounded-l-md bg-brand px-2.5 py-2 md:px-4 text-sm font-semibold text-white shadow-sm hover:bg-brand-dark disabled:opacity-60"
        onClick={() => run("xlsx", () => downloadWorkbook(wb()))}
        disabled={!!busy}
      >
        {busy ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
        <span className="whitespace-nowrap max-md:hidden"><span className="max-2xl:hidden">Download Excel Workbook</span><span className="2xl:hidden">Download</span></span>
        <span className="md:hidden">.xlsx</span>
      </button>
      <button
        aria-label="More export options"
        aria-expanded={menu}
        className="rounded-r-md border-l border-white/30 bg-brand px-2 text-white hover:bg-brand-dark"
        onClick={(e) => {
          e.stopPropagation();
          setMenu(!menu);
        }}
      >
        <ChevronDown className="h-4 w-4" />
      </button>
      {menu && (
        <div className="absolute top-full right-0 z-50 mt-1 w-64 rounded-lg border border-slate-200 bg-white py-1 text-sm shadow-lg" onClick={(e) => e.stopPropagation()}>
          <MenuItem icon={<FileSpreadsheet />} label="Excel workbook (.xlsx)" onClick={() => run("xlsx", () => downloadWorkbook(wb()))} />
          <MenuItem
            icon={<FileText />}
            label="PDF"
            hint={pdf ? "All sheets, print-ready" : "Needs LibreOffice on the server"}
            disabled={!pdf}
            onClick={() => run("pdf", () => downloadWorkbook(wb(), "pdf"))}
          />
          <MenuItem
            icon={<FileText />}
            label="CSV (current sheet)"
            hint="Values only, formulas calculated"
            onClick={() =>
              run("csv", async () => {
                const { workbook, active } = useStore.getState();
                const name = `${workbook.title} - ${workbook.sheets[active].name}`.replace(/[^\w\- ]+/g, "").trim() || "sheet";
                saveBlob(new Blob([sheetToCsv(workbook, active)], { type: "text/csv;charset=utf-8" }), `${name}.csv`);
              })
            }
          />
          <MenuItem
            icon={<Sheet />}
            label="Open in Google Sheets"
            hint={GOOGLE_CLIENT_ID ? "Saves a copy to your Google Drive" : "Not set up on this site yet"}
            disabled={!GOOGLE_CLIENT_ID}
            onClick={() =>
              run("google", async () => {
                const tab = window.open("about:blank", "_blank");
                try {
                  const url = await openInGoogleSheets(wb());
                  if (tab) tab.location.href = url;
                  else window.open(url, "_blank");
                } catch (e) {
                  tab?.close();
                  throw e;
                }
              })
            }
          />
        </div>
      )}
    </div>
  );
}

function MenuItem({ icon, label, hint, onClick, disabled }: { icon: React.ReactNode; label: string; hint?: string; onClick: () => void; disabled?: boolean }) {
  return (
    <button
      disabled={disabled}
      onClick={onClick}
      className="flex w-full items-start gap-2.5 px-3 py-2 text-left hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50 [&_svg]:mt-0.5 [&_svg]:h-4 [&_svg]:w-4 [&_svg]:text-brand"
    >
      {icon}
      <span>
        <span className="block text-slate-800">{label}</span>
        {hint && <span className="block text-xs text-slate-500">{hint}</span>}
      </span>
    </button>
  );
}
