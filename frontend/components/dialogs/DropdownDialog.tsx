"use client";

import { useState } from "react";
import Modal, { Button, Field, inputClass } from "@/components/ui/Modal";
import { selectionRange, useStore } from "@/lib/store";
import { useUi } from "@/lib/ui";

/** Turn the selected cells into an Excel dropdown list. */
export default function DropdownDialog() {
  const close = useUi((s) => s.close);
  const range = selectionRange(useStore.getState().selection);
  const existing = useStore((s) => (s.workbook.sheets[s.active].validations ?? []).find((v) => v.range === range));
  const [text, setText] = useState((existing?.options ?? []).join("\n"));
  const options = [...new Set(text.split("\n").map((o) => o.trim()).filter(Boolean))];

  return (
    <Modal
      title="Dropdown list"
      onClose={close}
      footer={
        <>
          {existing && (
            <Button
              onClick={() => {
                useStore.getState().setDropdown(null);
                close();
              }}
            >
              Remove dropdown
            </Button>
          )}
          <Button onClick={close}>Cancel</Button>
          <Button
            primary
            disabled={!options.length}
            onClick={() => {
              useStore.getState().setDropdown(options);
              close();
            }}
          >
            Apply to {range}
          </Button>
        </>
      }
    >
      <Field label="Choices, one per line" hint="Cells in this range will only accept these values in Excel.">
        <textarea
          autoFocus
          rows={6}
          className={inputClass}
          placeholder={"Not started\nIn progress\nDone"}
          value={text}
          onChange={(e) => setText(e.target.value)}
        />
      </Field>
    </Modal>
  );
}
