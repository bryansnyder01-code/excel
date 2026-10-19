"use client";

import { useEffect, useState } from "react";
import { Check, CreditCard, Loader2, Sparkles } from "lucide-react";
import Modal, { Button } from "@/components/ui/Modal";
import { useAccount } from "@/lib/account";
import { send } from "@/lib/api";
import { useUi } from "@/lib/ui";

interface Plan {
  plan: "free" | "pro";
  limit: number;
  used: number;
  renews_at: number | null;
  billing_enabled: boolean;
}

/** Your plan, today's AI use, and upgrading or managing a subscription through Stripe. */
export default function BillingDialog() {
  const close = useUi((s) => s.close);
  const user = useAccount((s) => s.user);
  const [plan, setPlan] = useState<Plan | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!user) return;
    send<Plan>("GET", "/api/billing/plan")
      .then(setPlan)
      .catch((e) => setError(e.message));
  }, [user]);

  const go = async (path: "checkout" | "portal") => {
    setBusy(true);
    setError(null);
    try {
      const { url } = await send<{ url: string }>("POST", `/api/billing/${path}`, {});
      window.location.href = url;
    } catch (e) {
      setError(e instanceof Error ? e.message : "That didn't work.");
      setBusy(false);
    }
  };

  if (!user)
    return (
      <Modal title="Plans" onClose={close} width="max-w-md">
        <p className="mb-4 text-sm text-slate-600">Sign in to see your plan and upgrade for more AI requests each day.</p>
        <div className="flex justify-end">
          <Button primary onClick={() => useUi.getState().open("account")}>
            Sign in
          </Button>
        </div>
      </Modal>
    );

  const pct = plan ? Math.min(100, Math.round((plan.used / Math.max(1, plan.limit)) * 100)) : 0;
  return (
    <Modal title="Plan and usage" onClose={close} width="max-w-lg">
      {!plan && !error && (
        <p className="flex items-center gap-2 text-sm text-slate-500">
          <Loader2 className="h-4 w-4 animate-spin" /> Loading your plan...
        </p>
      )}
      {error && <p className="mb-3 text-sm text-red-600" role="alert">{error}</p>}
      {plan && (
        <>
          <div className="mb-4 rounded-lg border border-slate-200 p-3">
            <div className="mb-1 flex items-center justify-between text-sm">
              <span className="font-medium text-slate-800">
                {plan.plan === "pro" ? "Pro" : "Free"} plan
                {plan.renews_at && <span className="font-normal text-slate-500"> · renews {new Date(plan.renews_at).toLocaleDateString()}</span>}
              </span>
              <span className="text-slate-600">
                {plan.used} of {plan.limit} AI requests today
              </span>
            </div>
            <div className="h-2 overflow-hidden rounded-full bg-slate-100" role="progressbar" aria-valuenow={pct} aria-valuemin={0} aria-valuemax={100} aria-label="AI requests used today">
              <div className={`h-full ${pct >= 90 ? "bg-red-500" : "bg-brand"}`} style={{ width: `${pct}%` }} />
            </div>
          </div>
          {plan.plan === "free" ? (
            <div className="rounded-lg border-2 border-brand/40 p-4">
              <h3 className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-slate-800">
                <Sparkles className="h-4 w-4 text-brand" /> Pro
              </h3>
              <ul className="mb-3 space-y-1 text-sm text-slate-600">
                {["Ten times the daily AI requests", "AI column fill and =AI() on big sheets", "Scheduled email reports with AI summaries"].map((t) => (
                  <li key={t} className="flex items-center gap-2">
                    <Check className="h-3.5 w-3.5 text-brand" /> {t}
                  </li>
                ))}
              </ul>
              {plan.billing_enabled ? (
                <Button primary disabled={busy} onClick={() => go("checkout")}>
                  <span className="flex items-center gap-1.5">
                    {busy ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <CreditCard className="h-3.5 w-3.5" />} Upgrade
                  </span>
                </Button>
              ) : (
                <p className="text-xs text-slate-500">Upgrades aren&apos;t switched on for this site yet.</p>
              )}
            </div>
          ) : (
            plan.billing_enabled && (
              <Button disabled={busy} onClick={() => go("portal")}>
                <span className="flex items-center gap-1.5">
                  <CreditCard className="h-3.5 w-3.5" /> Manage subscription
                </span>
              </Button>
            )
          )}
        </>
      )}
    </Modal>
  );
}
