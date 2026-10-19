"use client";

import { useEffect, useState } from "react";
import { LayoutTemplate, Trash2 } from "lucide-react";
import Modal from "@/components/ui/Modal";
import { openUserTemplate } from "@/lib/actions";
import type { UserTemplateInfo } from "@/lib/api";
import { useUi } from "@/lib/ui";
import { listUserTemplates, removeUserTemplate } from "@/lib/userTemplates";
import TemplateCards from "./TemplateCards";

export default function TemplatesDialog() {
  const close = useUi((s) => s.close);
  const [mine, setMine] = useState<UserTemplateInfo[]>([]);

  useEffect(() => {
    listUserTemplates()
      .then(setMine)
      .catch(() => {});
  }, []);

  return (
    <Modal title="Start from a template" onClose={close} width="max-w-2xl">
      {mine.length > 0 && (
        <section aria-label="Your templates" className="mb-5">
          <h3 className="mb-2 text-sm font-semibold text-slate-700">Your templates</h3>
          <div className="grid grid-cols-2 gap-2">
            {mine.map((t) => (
              <div key={t.id} className="group relative">
                <button
                  className="flex w-full items-start gap-3 rounded-lg border border-slate-200 p-3 text-left hover:border-brand hover:bg-brand/5"
                  onClick={() => {
                    close();
                    openUserTemplate(t.id, t.title);
                  }}
                >
                  <LayoutTemplate className="mt-0.5 h-5 w-5 shrink-0 text-brand" />
                  <span className="min-w-0">
                    <span className="block truncate text-sm font-medium text-slate-800">{t.title}</span>
                    <span className="block truncate text-xs text-slate-500">
                      {t.description || `${t.sheets} ${t.sheets === 1 ? "sheet" : "sheets"}`}
                    </span>
                  </span>
                </button>
                <button
                  aria-label={`Delete template ${t.title}`}
                  className="absolute top-2 right-2 rounded p-1 text-slate-400 opacity-0 group-hover:opacity-100 hover:bg-red-50 hover:text-red-600 focus:opacity-100"
                  onClick={async () => {
                    await removeUserTemplate(t.id).catch(() => {});
                    setMine(mine.filter((x) => x.id !== t.id));
                  }}
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              </div>
            ))}
          </div>
        </section>
      )}
      {mine.length > 0 && <h3 className="mb-2 text-sm font-semibold text-slate-700">Ready-made templates</h3>}
      <p className="mb-4 text-sm text-slate-500">
        Each template has live formulas. Open one, then edit cells or ask the AI to tailor it. Your current workbook stays saved.
      </p>
      <TemplateCards onPicked={close} />
    </Modal>
  );
}
