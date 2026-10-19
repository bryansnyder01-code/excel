// Keep formulas pointing at the right cells when rows/columns are inserted or deleted,
// the way Excel does.

import { columnLetter } from "./format";

const REF_RE =
  /(?<![A-Za-z0-9_.])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])/g;

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n; // 1-based
}

function sheetOf(prefix: string | undefined, own: string): string {
  if (!prefix) return own;
  const name = prefix.slice(0, -1);
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

export type Axis = "row" | "col";

/**
 * Shift references in `formula` (living on sheet `own`) that point at sheet `target`.
 * Indices are Excel 1-based. delta=+1 inserts before `at`; delta=-1 deletes `at`.
 */
export function shiftFormula(formula: string, own: string, target: string, axis: Axis, at: number, delta: 1 | -1): string {
  // Split on string literals so text like "A1" inside quotes is never touched.
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 === 1 ? part : shiftPart(part, own, target, axis, at, delta)))
    .join("");
}

function shiftPart(text: string, own: string, target: string, axis: Axis, at: number, delta: 1 | -1): string {
  return text.replace(REF_RE, (m, prefix, ca1, c1, ra1, r1, ca2, c2, ra2, r2) => {
    if (sheetOf(prefix, own).toLowerCase() !== target.toLowerCase()) return m;
    const get = (c: string, r: string) => (axis === "row" ? parseInt(r, 10) : colIndex(c));
    let start = get(c1, r1);
    let end = c2 ? get(c2, r2) : start;

    if (delta === 1) {
      if (start >= at) start += 1;
      if (end >= at) end += 1;
    } else {
      if (start === at && end === at) return "#REF!";
      if (start > at) start -= 1;
      if (end >= at) end -= 1;
    }
    const fmt = (c: string, cAbs: string, r: string, rAbs: string, v: number) =>
      axis === "row" ? `${cAbs}${c}${rAbs}${v}` : `${cAbs}${columnLetter(v - 1)}${rAbs}${r}`;
    const first = fmt(c1, ca1, r1, ra1, start);
    return (prefix ?? "") + (c2 ? `${first}:${fmt(c2, ca2, r2, ra2, end)}` : first);
  });
}

/** Rewrite references to a renamed sheet. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const quoted = `'${newName.replace(/'/g, "''")}'!`;
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => {
      if (i % 2 === 1) return part;
      return part.replace(/(?<![A-Za-z0-9_.])('(?:[^']|'')+'|[A-Za-z_][\w.]*)!/g, (m, name: string) =>
        sheetOf(`${name}!`, "").toLowerCase() === oldName.toLowerCase() ? quoted : m,
      );
    })
    .join("");
}

/**
 * Move a formula by (dr, dc) the way copy/paste and fill do in Excel: relative references
 * move with it, $-anchored parts stay. References pushed off the sheet become #REF!.
 */
export function offsetFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) =>
      i % 2 === 1
        ? part
        : part.replace(REF_RE, (m, prefix, ca1, c1, ra1, r1, ca2, c2, ra2, r2) => {
            const one = (cAbs: string, c: string, rAbs: string, r: string) => {
              const col = cAbs ? colIndex(c) : colIndex(c) + dc;
              const row = rAbs ? parseInt(r, 10) : parseInt(r, 10) + dr;
              if (col < 1 || row < 1) return null;
              return `${cAbs}${columnLetter(col - 1)}${rAbs}${row}`;
            };
            const first = one(ca1, c1, ra1, r1);
            const second = c2 ? one(ca2, c2, ra2, r2) : "";
            if (first === null || second === null) return "#REF!";
            return (prefix ?? "") + (c2 ? `${first}:${second}` : first);
          }),
    )
    .join("");
}

/** Cell references in a formula, in order of appearance (for colouring them in the grid). */
export function formulaRefs(formula: string): { sheet?: string; r1: number; c1: number; r2: number; c2: number }[] {
  const out: { sheet?: string; r1: number; c1: number; r2: number; c2: number }[] = [];
  formula.split(/("(?:[^"]|"")*")/).forEach((part, i) => {
    if (i % 2 === 1) return;
    for (const m of part.matchAll(REF_RE)) {
      const [, prefix, , c1, , r1, , c2, , r2] = m;
      const a = { r: parseInt(r1, 10), c: colIndex(c1) };
      const b = c2 ? { r: parseInt(r2, 10), c: colIndex(c2) } : a;
      out.push({
        sheet: prefix ? sheetOf(prefix, "") : undefined,
        r1: Math.min(a.r, b.r),
        r2: Math.max(a.r, b.r),
        c1: Math.min(a.c, b.c),
        c2: Math.max(a.c, b.c),
      });
    }
  });
  return out;
}
