"use client";

import { useEffect, useState } from "react";
import { Copy, Link2, Loader2, UserPlus, Users } from "lucide-react";
import Modal, { Button, inputClass } from "@/components/ui/Modal";
import { saveCurrent } from "@/components/Workspace";
import { useAccount } from "@/lib/account";
import { cloud, type Member } from "@/lib/api";
import { getDoc } from "@/lib/persist";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

export const inviteUrl = (token: string) => `${window.location.origin}${window.location.pathname}?invite=${encodeURIComponent(token)}`;
export const shareUrl = (token: string) => `${window.location.origin}${window.location.pathname}?share=${encodeURIComponent(token)}`;

/** Invite people to edit this workbook together, or share a copy of it with a link. */
export default function ShareDialog() {
  const close = useUi((s) => s.close);
  const user = useAccount((s) => s.user);
  const docId = useStore((s) => s.docId);
  const [link, setLink] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [members, setMembers] = useState<Member[]>([]);
  const [invites, setInvites] = useState<{ editor: string | null; viewer: string | null }>({ editor: null, viewer: null });
  const [ready, setReady] = useState(false);

  useEffect(() => {
    if (!user) return;
    let gone = false;
    (async () => {
      setBusy(true);
      try {
        // The account needs the latest version before it can be shared.
        saveCurrent();
        const doc = getDoc(docId);
        if (!doc) throw new Error("Add something to the workbook before sharing it.");
        let people = await cloud.members(doc.id).catch(() => null);
        const mine = people?.find((m) => m.id === user.id)?.role;
        if (mine !== "viewer") {
          await cloud.put(doc.id, doc.title, doc.updated, doc);
          people ??= await cloud.members(doc.id);
        }
        if (gone || !people) return;
        setMembers(people);
        setReady(true);
        if (people.find((m) => m.id === user.id)?.role === "owner") {
          const [{ token }, existing] = await Promise.all([cloud.share(doc.id), cloud.invites(doc.id).catch(() => ({ editor: null, viewer: null }))]);
          if (gone) return;
          setLink(shareUrl(token));
          setInvites({ editor: existing.editor && inviteUrl(existing.editor), viewer: existing.viewer && inviteUrl(existing.viewer) });
        }
      } catch (e) {
        if (!gone) setError(e instanceof Error ? e.message : "Couldn't create a link.");
      } finally {
        if (!gone) setBusy(false);
      }
    })();
    return () => {
      gone = true;
    };
  }, [user, docId]);

  if (!user) {
    return (
      <Modal title="Share this workbook" onClose={close} width="max-w-md">
        <p className="text-sm text-slate-600">Sign in to share workbooks with a link. It's free and keeps your workbooks in sync across devices.</p>
        <div className="mt-4 flex justify-end">
          <Button primary onClick={() => useUi.getState().open("account")}>
            Sign in or create an account
          </Button>
        </div>
      </Modal>
    );
  }

  const me = members.find((m) => m.id === user.id);
  const owner = me?.role === "owner" || (!me && !members.length);
  const copy = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      useUi.getState().notify("Link copied.");
    } catch {
      useUi.getState().notify("Select the link and copy it with Ctrl+C.", true);
    }
  };

  return (
    <Modal title="Share this workbook" onClose={close} width="max-w-lg">
      {busy && (
        <p className="flex items-center gap-2 text-sm text-slate-500">
          <Loader2 className="h-4 w-4 animate-spin" /> Getting your links ready...
        </p>
      )}
      {error && <p className="mb-2 text-sm text-red-600">{error}</p>}

      <section aria-label="Edit together">
        <h3 className="mb-1 flex items-center gap-1.5 text-sm font-semibold text-slate-800">
          <Users className="h-4 w-4 text-brand" /> Edit together
        </h3>
        <p className="mb-2 text-sm text-slate-600">
          People who open one of these links and sign in join this workbook. Editors' changes and everyone's selections show up as they happen. Viewers can look and comment, but not change anything.
        </p>
        {owner &&
          (["editor", "viewer"] as const).map((role) => {
            const url = invites[role];
            return (
              <div key={role} className="mb-2 flex items-center gap-2">
                <span className="w-20 shrink-0 text-xs font-medium text-slate-600">{role === "editor" ? "Can edit" : "Can view"}</span>
                {url ? (
                  <>
                    <input aria-label={role === "editor" ? "Invite link" : "View-only invite link"} readOnly className={inputClass} value={url} onFocus={(e) => e.target.select()} />
                    <Button primary={role === "editor"} onClick={() => copy(url)}>
                      <span className="flex items-center gap-1.5">
                        <Copy className="h-3.5 w-3.5" /> Copy
                      </span>
                    </Button>
                    <button
                      aria-label={`Turn off the ${role === "editor" ? "edit" : "view-only"} invite link`}
                      title="Turn off this link"
                      className="rounded p-1.5 text-red-600 hover:bg-red-50"
                      onClick={async () => {
                        try {
                          await cloud.uninvite(docId, role);
                          setInvites({ ...invites, [role]: null });
                          useUi.getState().notify("Invite link turned off. People who already joined keep access.");
                        } catch (e) {
                          setError(e instanceof Error ? e.message : "Couldn't turn the invite link off.");
                        }
                      }}
                    >
                      <Link2 className="h-4 w-4" />
                    </button>
                  </>
                ) : (
                  <Button
                    disabled={busy || !ready}
                    onClick={async () => {
                      try {
                        const { token } = await cloud.invite(docId, role);
                        setInvites({ ...invites, [role]: inviteUrl(token) });
                      } catch (e) {
                        setError(e instanceof Error ? e.message : "Couldn't create an invite link.");
                      }
                    }}
                  >
                    <span className="flex items-center gap-1.5">
                      <UserPlus className="h-3.5 w-3.5" /> {role === "editor" ? "Create invite link" : "Create view-only link"}
                    </span>
                  </Button>
                )}
              </div>
            );
          })}
        {members.length > 0 && (
          <ul className="mt-3 divide-y divide-slate-100 rounded-md border border-slate-200 text-sm" aria-label="People with access">
            {members.map((m) => (
              <li key={m.id} className="flex items-center gap-2 px-3 py-1.5">
                <span className="min-w-0 flex-1 truncate text-slate-700">
                  {m.email}
                  {m.id === user.id && <span className="text-slate-400"> (you)</span>}
                </span>
                {owner && m.role !== "owner" ? (
                  <select
                    aria-label={`Access for ${m.email}`}
                    className="rounded border border-slate-300 bg-white px-1 py-0.5 text-xs text-slate-700"
                    value={m.role}
                    onChange={async (e) => {
                      const role = e.target.value as "editor" | "viewer";
                      try {
                        await cloud.setRole(docId, m.id, role);
                        setMembers(members.map((x) => (x.id === m.id ? { ...x, role } : x)));
                      } catch (err) {
                        setError(err instanceof Error ? err.message : "Couldn't change their access.");
                      }
                    }}
                  >
                    <option value="editor">Can edit</option>
                    <option value="viewer">Can view</option>
                  </select>
                ) : (
                  <span className="text-xs text-slate-500">{m.role === "owner" ? "Owner" : m.role === "viewer" ? "Can view" : "Can edit"}</span>
                )}
                {m.role !== "owner" && (owner || m.id === user.id) && (
                  <button
                    className="text-xs text-red-600 hover:underline"
                    onClick={async () => {
                      try {
                        await cloud.removeMember(docId, m.id);
                        if (m.id === user.id) {
                          useUi.getState().notify("You left this workbook. Your copy in this browser stays.");
                          close();
                        } else setMembers(members.filter((x) => x.id !== m.id));
                      } catch (e) {
                        setError(e instanceof Error ? e.message : "Couldn't remove them.");
                      }
                    }}
                  >
                    {m.id === user.id ? "Leave" : "Remove"}
                  </button>
                )}
              </li>
            ))}
          </ul>
        )}
      </section>

      {owner && (
        <section aria-label="Send a copy" className="mt-5 border-t border-slate-200 pt-4">
          <h3 className="mb-1 flex items-center gap-1.5 text-sm font-semibold text-slate-800">
            <Copy className="h-4 w-4 text-brand" /> Send a copy
          </h3>
          <p className="mb-2 text-sm text-slate-600">Anyone with this link gets their own copy. Their changes never touch yours, and they don&apos;t need an account.</p>
          {link && (
            <>
              <div className="flex gap-2">
                <input aria-label="Share link" readOnly className={inputClass} value={link} onFocus={(e) => e.target.select()} />
                <Button onClick={() => copy(link)}>
                  <span className="flex items-center gap-1.5">
                    <Copy className="h-3.5 w-3.5" /> Copy
                  </span>
                </Button>
              </div>
              <button
                className="mt-2 flex items-center gap-1.5 text-xs text-red-600 hover:underline"
                onClick={async () => {
                  try {
                    await cloud.unshare(docId);
                    setLink(null);
                    useUi.getState().notify("Link turned off. It no longer opens this workbook.");
                  } catch (e) {
                    setError(e instanceof Error ? e.message : "Couldn't turn the link off.");
                  }
                }}
              >
                <Link2 className="h-3.5 w-3.5" /> Turn off this link
              </button>
            </>
          )}
        </section>
      )}
    </Modal>
  );
}
