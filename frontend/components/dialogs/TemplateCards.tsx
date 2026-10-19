"use client";

import { useEffect, useState } from "react";
import { Briefcase, CalendarClock, Clock, FileText, Landmark, Package, TrendingUp, Wallet } from "lucide-react";
import { listTemplates } from "@/lib/api";
import { openTemplate } from "@/lib/actions";
import type { TemplateInfo } from "@/lib/types";

const ICONS: Record<string, typeof Wallet> = {
  "small-business-budget": Wallet,
  invoice: FileText,
  "loan-amortization": Landmark,
  inventory: Package,
  timesheet: Clock,
  "sales-tracker": TrendingUp,
  "project-timeline": CalendarClock,
};

let cached: TemplateInfo[] | null = null;

export default function TemplateCards({ onPicked, compact }: { onPicked?: () => void; compact?: boolean }) {
  const [templates, setTemplates] = useState<TemplateInfo[] | null>(cached);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (cached) return;
    listTemplates()
      .then((t) => setTemplates((cached = t)))
      .catch((e) => setError(e.message));
  }, []);

  if (error) return <p className="text-sm text-red-600">{error}</p>;
  if (!templates) return <p className="text-sm text-slate-400">Loading templates...</p>;

  return (
    <div className={`grid gap-2 ${compact ? "grid-cols-1" : "grid-cols-2"}`}>
      {templates.map((t) => {
        const Icon = ICONS[t.id] ?? Briefcase;
        return (
          <button
            key={t.id}
            onClick={() => {
              onPicked?.();
              openTemplate(t.id, t.title);
            }}
            className="flex gap-3 rounded-lg border border-slate-200 p-3 text-left hover:border-brand hover:bg-green-50"
          >
            <Icon className="mt-0.5 h-5 w-5 shrink-0 text-brand" />
            <span>
              <span className="block text-sm font-medium text-slate-800">{t.title}</span>
              {!compact && <span className="mt-0.5 block text-xs text-slate-500">{t.description}</span>}
            </span>
          </button>
        );
      })}
    </div>
  );
}
