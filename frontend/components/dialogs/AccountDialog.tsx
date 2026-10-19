"use client";

import { useEffect, useRef, useState } from "react";
import { Cloud, Loader2, LogOut, RefreshCw } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { pendingReset, setPendingReset, signIn, signOut, startSession, syncAll, useAccount } from "@/lib/account";
import { auth } from "@/lib/api";
import { loadGis } from "@/lib/google";
import { useUi } from "@/lib/ui";

type Kind = "login" | "signup" | "forgot" | "reset";

/** Sign in or create an account so workbooks follow you to other devices. */
export default function AccountDialog() {
  const close = useUi((s) => s.close);
  const { user, syncing, lastSync, syncError } = useAccount();
  const [resetToken] = useState(() => pendingReset);
  const [kind, setKind] = useState<Kind>(resetToken ? "reset" : "login");
  const [sent, setSent] = useState(false);
  const [googleId, setGoogleId] = useState<string | null>(null);
  const [canReset, setCanReset] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    auth
      .config()
      .then((c) => {
        setGoogleId(c.google_client_id || null);
        setCanReset(!!c.password_reset);
      })
      .catch(() => {});
  }, []);

  const google = async (credential: string) => {
    setError(null);
    try {
      await startSession(await auth.google(credential));
      finish("Signed in with Google. Your workbooks now sync.");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Google sign-in didn't work.");
    }
  };

  if (user && kind !== "reset") {
    return (
      <Modal title="Your account" onClose={close}>
        <div className="flex items-center gap-3">
          <Cloud className="h-8 w-8 text-brand" />
          <div>
            <div className="text-sm font-medium text-slate-800">{user.email}</div>
            <div className="text-xs text-slate-500">
              {syncing ? "Syncing..." : syncError ? syncError : lastSync ? `Workbooks synced ${new Date(lastSync).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}` : "Workbooks sync automatically"}
            </div>
          </div>
        </div>
        <p className="mt-4 text-sm text-slate-600">
          Every workbook you edit is saved to your account, so it shows up under <strong>My workbooks</strong> on any device you sign in on.
        </p>
        <div className="mt-5 flex gap-2">
          <Button onClick={() => syncAll()} disabled={syncing}>
            <span className="flex items-center gap-1.5">
              <RefreshCw className={`h-3.5 w-3.5 ${syncing ? "animate-spin" : ""}`} /> Sync now
            </span>
          </Button>
          <Button
            onClick={async () => {
              await signOut();
              close();
            }}
          >
            <span className="flex items-center gap-1.5">
              <LogOut className="h-3.5 w-3.5" /> Sign out
            </span>
          </Button>
        </div>
        <p className="mt-3 text-xs text-slate-400">Signing out keeps the copies already in this browser.</p>
      </Modal>
    );
  }

  const finish = (text: string) => {
    useUi.getState().notify(text);
    close();
  };

  const submit = async () => {
    setBusy(true);
    setError(null);
    try {
      if (kind === "forgot") {
        await auth.forgot(email.trim());
        setSent(true);
      } else if (kind === "reset") {
        await startSession(await auth.reset(resetToken ?? "", password));
        setPendingReset(null);
        finish("Password changed. You're signed in.");
      } else {
        await signIn(kind, email.trim(), password);
        finish(kind === "signup" ? "Account created. Your workbooks now sync." : "Signed in. Your workbooks now sync.");
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Something went wrong.");
    } finally {
      setBusy(false);
    }
  };

  const switchTo = (k: Kind) => {
    setKind(k);
    setError(null);
    setSent(false);
  };

  const titles: Record<Kind, string> = {
    login: "Sign in",
    signup: "Create an account",
    forgot: "Reset your password",
    reset: "Choose a new password",
  };

  return (
    <Modal title={titles[kind]} onClose={close} width="max-w-md">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          submit();
        }}
      >
        {kind === "forgot" ? (
          <p className="mb-4 text-sm text-slate-600">
            {sent ? "If that email has an account, a link to reset the password is on its way. It works for one hour." : "Enter your email and we'll send you a link to choose a new password."}
          </p>
        ) : kind === "reset" ? (
          <p className="mb-4 text-sm text-slate-600">Pick a new password for your account.</p>
        ) : (
          <p className="mb-4 text-sm text-slate-600">Keep your workbooks in sync across devices, share them, and edit them together.</p>
        )}
        {(kind === "login" || kind === "signup") && googleId && (
          <>
            <GoogleButton clientId={googleId} onCredential={google} />
            <div className="my-4 flex items-center gap-3 text-xs text-slate-400">
              <span className="h-px flex-1 bg-slate-200" /> or with email <span className="h-px flex-1 bg-slate-200" />
            </div>
          </>
        )}
        {kind !== "reset" && !(kind === "forgot" && sent) && (
          <Field label="Email">
            <input autoFocus type="email" autoComplete="email" className={inputClass} value={email} onChange={(e) => setEmail(e.target.value)} />
          </Field>
        )}
        {kind !== "forgot" && (
          <Field label={kind === "reset" ? "New password" : "Password"} hint={kind !== "login" ? "At least 8 characters" : undefined}>
            <input
              autoFocus={kind === "reset"}
              type="password"
              autoComplete={kind === "login" ? "current-password" : "new-password"}
              className={inputClass}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </Field>
        )}
        {kind === "login" && canReset && (
          <button type="button" className="-mt-1 mb-3 block text-xs text-slate-500 hover:text-brand hover:underline" onClick={() => switchTo("forgot")}>
            Forgot your password?
          </button>
        )}
        {error && <p className="mb-3 text-sm text-red-600" role="alert">{error}</p>}
        <div className="flex items-center justify-between gap-2">
          <button type="button" className="text-sm text-brand hover:underline" onClick={() => switchTo(kind === "login" ? "signup" : "login")}>
            {kind === "login" ? "New here? Create an account" : "Back to sign in"}
          </button>
          {!(kind === "forgot" && sent) && (
            <Button type="submit" primary disabled={busy || (kind !== "reset" && !email.trim()) || (kind !== "forgot" && !password)}>
              <span className="flex items-center gap-1.5">
                {busy && <Loader2 className="h-3.5 w-3.5 animate-spin" />}
                {{ login: "Sign in", signup: "Create account", forgot: "Send link", reset: "Save password" }[kind]}
              </span>
            </Button>
          )}
        </div>
      </form>
    </Modal>
  );
}

/** Google's own "Sign in with Google" button (Google Identity Services). */
function GoogleButton({ clientId, onCredential }: { clientId: string; onCredential: (credential: string) => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const cb = useRef(onCredential);
  cb.current = onCredential;
  useEffect(() => {
    let live = true;
    loadGis()
      .then(() => {
        if (!live || !ref.current || !window.google) return;
        window.google.accounts.id.initialize({ client_id: clientId, callback: (r) => cb.current(r.credential) });
        window.google.accounts.id.renderButton(ref.current, { theme: "outline", size: "large", text: "continue_with", width: 380 });
      })
      .catch(() => {});
    return () => {
      live = false;
    };
  }, [clientId]);
  return <div ref={ref} className="flex min-h-10 justify-center" aria-label="Sign in with Google" />;
}
