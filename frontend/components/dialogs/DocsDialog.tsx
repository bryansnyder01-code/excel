"use client";

import { useEffect, useState } from "react";
import { saveCurrent } from "@/components/Workspace";
import { FileSpreadsheet, Trash2 } from "lucide-react";
import Modal from "@/components/ui/Modal";
import { removeFromAccount, syncAll, useAccount } from "@/lib/account";
import { deleteDoc, getDoc, listDocs } from "@/lib/persist";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

function when(t: number) {
  const mins = Math.round((Date.now() - t) / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return `${mins} min ago`;
  if (mins < 60 * 24) return `${Math.round(mins / 60)} h ago`;
  return new Date(t).toLocaleDateString();
}

export default function DocsDialog() {
  const close = useUi((s) => s.close);
  const currentId = useStore((s) => s.docId);
  const user = useAccount((s) => s.user);
  const syncing = useAccount((s) => s.syncing);
  const [docs, setDocs] = useState(() => {
    saveCurrent();
    return listDocs();
  });
  // Pull in anything saved from another device.
  useEffect(() => {
    if (user) syncAll().then(() => setDocs(listDocs()));
  }, [user]);

  return (
    <Modal title="My workbooks" onClose={close} width="max-w-xl">
      <p className="mb-3 text-xs text-slate-500">
        {user
          ? syncing
            ? "Syncing with your account..."
            : `Workbooks save automatically and sync to ${user.email}.`
          : "Workbooks save automatically in this browser as you work. Sign in to keep them on every device."}
      </p>
      {docs.length === 0 && <p className="text-sm text-slate-500">Nothing saved yet.</p>}
      <ul className="divide-y divide-slate-100">
        {docs.map((d) => (
          <li key={d.id} className="flex items-center gap-3 py-2.5">
            <FileSpreadsheet className="h-5 w-5 shrink-0 text-brand" />
            <button
              className="min-w-0 flex-1 text-left"
              onClick={() => {
                const doc = getDoc(d.id);
                if (doc) useStore.getState().loadDocument(doc);
                close();
              }}
            >
              <span className="block truncate text-sm font-medium text-slate-800">
                {d.title}
                {d.id === currentId && <span className="ml-2 text-xs font-normal text-brand">open now</span>}
              </span>
              <span className="text-xs text-slate-500">
                {d.workbook.sheets.length} {d.workbook.sheets.length === 1 ? "sheet" : "sheets"} · edited {when(d.updated)}
              </span>
            </button>
            <button
              aria-label={`Delete ${d.title}`}
              title="Delete"
              className="rounded p-1.5 text-slate-400 hover:bg-red-50 hover:text-red-600"
              onClick={() => {
                if (!confirm(`Delete "${d.title}"?${user ? " It will be removed from your account too." : ""} This can't be undone.`)) return;
                // Switch away first: switching saves the document being left.
                if (d.id === currentId) useStore.getState().newWorkbook();
                deleteDoc(d.id);
                removeFromAccount(d.id);
                setDocs(listDocs());
              }}
            >
              <Trash2 className="h-4 w-4" />
            </button>
          </li>
        ))}
      </ul>
    </Modal>
  );
}
