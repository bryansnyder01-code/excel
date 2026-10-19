"use client";

import { useEffect, useRef, useState } from "react";
import { Camera, FileText, ImageUp, Loader2 } from "lucide-react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { errorText, extractFile } from "@/lib/aiApi";
import { isBlank, useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

const MAX_BYTES = 20 * 1024 * 1024;
const ACCEPT = "image/*,application/pdf";

const accepted = (f: File) => f.type.startsWith("image/") || f.type === "application/pdf" || /\.(pdf|png|jpe?g|webp|heic|heif|gif)$/i.test(f.name);

/** Turn a photo, screenshot or PDF of a table (receipt, bank statement, invoice) into a new workbook. */
export default function ImportImageDialog() {
  const close = useUi((s) => s.close);
  const [file, setFile] = useState<File | null>(null);
  const [url, setUrl] = useState<string | null>(null);
  const [hint, setHint] = useState("");
  const [busy, setBusy] = useState(false);
  const [seconds, setSeconds] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [over, setOver] = useState(false);
  const pick = useRef<HTMLInputElement>(null);
  const camera = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!file || !file.type.startsWith("image/")) return setUrl(null);
    const u = URL.createObjectURL(file);
    setUrl(u);
    return () => URL.revokeObjectURL(u);
  }, [file]);

  useEffect(() => {
    if (!busy) return;
    setSeconds(0);
    const t = setInterval(() => setSeconds((s) => s + 1), 1000);
    return () => clearInterval(t);
  }, [busy]);

  const choose = (f: File | undefined | null) => {
    setError(null);
    if (!f) return;
    if (!accepted(f)) return setError("Choose a photo (JPG, PNG, HEIC...) or a PDF.");
    if (f.size > MAX_BYTES) return setError("That file is over 20 MB. Try a smaller photo or fewer pages.");
    setFile(f);
  };

  const extract = async () => {
    if (!file || busy) return;
    setBusy(true);
    setError(null);
    try {
      const res = await extractFile(file, hint);
      // Opens like an upload: a new workbook unless the open one is still empty.
      const st = useStore.getState();
      if (!isBlank(st.workbook)) st.newWorkbook();
      const label = `Imported ${file.name}`;
      useStore.getState().addMessage({ role: "user", text: `Imported ${file.name}${hint.trim() ? ` (${hint.trim()})` : ""}` });
      useStore.getState().applyResult(res.workbook, label, { text: res.summary, issues: res.issues });
      useUi.setState({ mobilePane: "sheet" });
      close();
    } catch (e) {
      setError(errorText(e, "Couldn't read a table from that file."));
    } finally {
      setBusy(false);
    }
  };

  const stage = seconds < 3 ? "Uploading" : seconds < 20 ? "Reading the table" : "Still reading (long documents take a minute)";

  return (
    <Modal
      title="Import from a photo or PDF"
      onClose={close}
      width="max-w-lg"
      footer={
        <>
          <Button onClick={close}>Cancel</Button>
          <Button primary disabled={!file || busy} onClick={extract}>
            {busy ? "Reading..." : "Make a spreadsheet"}
          </Button>
        </>
      }
    >
      <p className="mb-3 text-sm text-slate-600">
        Receipts, bank statements, invoices, price lists or a screenshot of a table. The AI reads it into a new workbook you can check and edit.
      </p>

      <div
        className={`mb-3 flex min-h-[10rem] flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed p-4 text-center ${
          over ? "border-brand bg-slate-50" : "border-slate-300"
        }`}
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          choose(e.dataTransfer.files?.[0]);
        }}
      >
        {file ? (
          url ? (
            // eslint-disable-next-line @next/next/no-img-element
            <img src={url} alt={`Preview of ${file.name}`} className="max-h-56 max-w-full rounded border border-slate-200 object-contain" />
          ) : (
            <div className="flex items-center gap-2 text-sm text-slate-700">
              <FileText className="h-8 w-8 text-red-600" /> {file.name}
            </div>
          )
        ) : (
          <>
            <ImageUp className="h-8 w-8 text-slate-400" />
            <p className="text-sm text-slate-500">Drop a photo or PDF here</p>
          </>
        )}
        <div className="mt-1 flex flex-wrap justify-center gap-2">
          <Button onClick={() => pick.current?.click()} disabled={busy}>
            {file ? "Choose another file" : "Choose a file"}
          </Button>
          <Button onClick={() => camera.current?.click()} disabled={busy}>
            <span className="flex items-center gap-1.5">
              <Camera className="h-4 w-4" /> Take a photo
            </span>
          </Button>
        </div>
        <input ref={pick} type="file" accept={ACCEPT} className="hidden" onChange={(e) => (choose(e.target.files?.[0]), (e.target.value = ""))} />
        <input
          ref={camera}
          type="file"
          accept={ACCEPT}
          capture="environment"
          className="hidden"
          onChange={(e) => (choose(e.target.files?.[0]), (e.target.value = ""))}
        />
      </div>

      <Field label="What is it? (optional)" hint="A hint helps, e.g. &quot;bank statement, keep date, description and amount&quot;.">
        <input className={inputClass} value={hint} placeholder="This is a bank statement" onChange={(e) => setHint(e.target.value)} disabled={busy} />
      </Field>

      {busy && (
        <p className="flex items-center gap-2 text-sm text-slate-600" role="status">
          <Loader2 className="h-4 w-4 animate-spin" /> {stage}... {seconds}s
        </p>
      )}
      {error && <p className="text-sm text-red-600">{error}</p>}
    </Modal>
  );
}
