"use client";

import { useState } from "react";
import { Loader2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { useAccount } from "@/lib/account";
import { useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";
import { saveUserTemplate, templateFrom } from "@/lib/userTemplates";

/** Save the open workbook's layout, formulas and formatting as a template to start new workbooks from. */
export default function SaveTemplateDialog() {
  const close = useUi((s) => s.close);
  const workbook = useStore((s) => s.workbook);
  const user = useAccount((s) => s.user);
  const [title, setTitle] = useState(workbook.title);
  const [description, setDescription] = useState("");
  const [keepData, setKeepData] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const save = async () => {
    setBusy(true);
    setError(null);
    try {
      await saveUserTemplate(title, description, templateFrom(workbook, keepData));
      useUi.getState().notify(`Saved "${title.trim() || workbook.title}" to your templates.`);
      close();
    } catch (e) {
      setError(e instanceof Error ? e.message : "The template wasn't saved.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <Modal title="Save as template" onClose={close} width="max-w-md">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <Field label="Name">
          <input autoFocus className={inputClass} value={title} onChange={(e) => setTitle(e.target.value)} />
        </Field>
        <Field label="Description" hint="Optional: what it's for">
          <input className={inputClass} value={description} onChange={(e) => setDescription(e.target.value)} />
        </Field>
        <label className="mb-4 flex items-start gap-2 text-sm text-slate-700">
          <input type="checkbox" className="mt-0.5" checked={keepData} onChange={(e) => setKeepData(e.target.checked)} />
          <span>
            Keep the numbers and text I typed
            <span className="block text-xs text-slate-500">Off: data rows are emptied, and headers, labels, formulas and formatting stay.</span>
          </span>
        </label>
        <p className="mb-3 text-xs text-slate-500">
          {user ? "Saved to your account, so it's there on every device." : "Saved in this browser. Sign in to keep templates in your account."}
        </p>
        {error && <p className="mb-3 text-sm text-red-600" role="alert">{error}</p>}
        <div className="flex justify-end gap-2">
          <Button onClick={close}>Cancel</Button>
          <Button type="submit" primary disabled={busy}>
            <span className="flex items-center gap-1.5">
              {busy && <Loader2 className="h-3.5 w-3.5 animate-spin" />} Save template
            </span>
          </Button>
        </div>
      </form>
    </Modal>
  );
}
