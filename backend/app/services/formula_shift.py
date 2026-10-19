"""Structural operations on a workbook dict (insert/delete rows and columns), moving every
reference the way Excel does. A line-for-line port of the frontend's `applyOp` / `shiftAll`
(frontend/lib/store.ts) and `shiftFormula` (frontend/lib/refs.ts), so the live server's copy
of a workbook stays identical to every client's after the same ops.

Indices: `shift_formula` works in Excel 1-based rows/columns. Ops use grid coordinates:
rows `at` is the grid row (grid row r = rows[r-1], Excel row r+1); columns `at` is 0-based.
"""

from __future__ import annotations

import re
from typing import Literal

Axis = Literal["row", "col"]

# Same pattern as refs.ts REF_RE (\w and \d are ASCII-only in JavaScript, hence re.ASCII).
REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)"
    r"(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])",
    re.ASCII,
)
_STRINGS = re.compile(r'("(?:[^"]|"")*")')

OPS = ("insert_rows", "delete_rows", "insert_cols", "delete_cols")

# Inserting/deleting columns inside the frozen columns grows/shrinks freeze_columns, like Excel.
# Off until the frontend's shiftAll does the same (store.ts: `if (axis === "col" && at - 1 <
# (sheet.freeze_columns ?? 0)) sheet.freeze_columns = Math.max(0, sheet.freeze_columns + delta)`):
# the server's copy must stay identical to the clients'.
SHIFT_FREEZE = True


def col_index(letters: str) -> int:
    """'A' -> 1, 'AA' -> 27 (1-based)."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def column_letter(index: int) -> str:
    """0-based index -> letters (format.ts columnLetter); negative gives ''."""
    n, s = index + 1, ""
    while n > 0:
        m = (n - 1) % 26
        s = chr(65 + m) + s
        n = (n - 1) // 26
    return s


def _sheet_of(prefix: str | None, own: str) -> str:
    if not prefix:
        return own
    name = prefix[:-1]
    return name[1:-1].replace("''", "'") if name.startswith("'") else name


def shift_formula(formula: str, own: str, target: str, axis: Axis, at: int, delta: int) -> str:
    """Shift references in `formula` (on sheet `own`) that point at sheet `target`.
    1-based indices; delta=+1 inserts before `at`, delta=-1 deletes `at`. Text in quotes is untouched."""
    parts = _STRINGS.split(formula)
    return "".join(p if i % 2 else _shift_part(p, own, target, axis, at, delta) for i, p in enumerate(parts))


def _shift_part(text: str, own: str, target: str, axis: Axis, at: int, delta: int) -> str:
    def repl(m: re.Match) -> str:
        prefix, ca1, c1, ra1, r1, ca2, c2, ra2, r2 = m.groups()
        if _sheet_of(prefix, own).lower() != target.lower():
            return m.group(0)

        def get(c: str, r: str) -> int:
            return int(r) if axis == "row" else col_index(c)

        start = get(c1, r1)
        end = get(c2, r2) if c2 else start
        if delta == 1:
            if start >= at:
                start += 1
            if end >= at:
                end += 1
        else:
            if start == at and end == at:
                return "#REF!"
            if start > at:
                start -= 1
            if end >= at:
                end -= 1

        def fmt(c: str, c_abs: str, r: str, r_abs: str, v: int) -> str:
            return f"{c_abs}{c}{r_abs}{v}" if axis == "row" else f"{c_abs}{column_letter(v - 1)}{r_abs}{r}"

        first = fmt(c1, ca1, r1, ra1, start)
        body = f"{first}:{fmt(c2, ca2 or '', r2, ra2 or '', end)}" if c2 else first
        return (prefix or "") + body

    return REF_RE.sub(repl, text)


def offset_formula(formula: str, dr: int, dc: int) -> str:
    """Move a formula by (dr, dc) like copy/paste: relative parts move, $ parts stay; off-sheet -> #REF!."""
    if not dr and not dc:
        return formula

    def repl(m: re.Match) -> str:
        prefix, ca1, c1, ra1, r1, ca2, c2, ra2, r2 = m.groups()

        def one(c_abs, c, r_abs, r):
            col = col_index(c) if c_abs else col_index(c) + dc
            row = int(r) if r_abs else int(r) + dr
            if col < 1 or row < 1:
                return None
            return f"{c_abs}{column_letter(col - 1)}{r_abs}{row}"

        first = one(ca1, c1, ra1, r1)
        second = one(ca2 or "", c2, ra2 or "", r2) if c2 else ""
        if first is None or second is None:
            return "#REF!"
        return (prefix or "") + (f"{first}:{second}" if c2 else first)

    parts = _STRINGS.split(formula)
    return "".join(p if i % 2 else REF_RE.sub(repl, p) for i, p in enumerate(parts))


# ---------------------------------------------------------------------------
# A1 helpers (used by protection checks and cell history)
# ---------------------------------------------------------------------------

_A1 = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_cell(a1: str) -> tuple[int, int] | None:
    """'B5' -> (grid row 4, col 1). Excel row n is grid row n-1 (row 1 is the header row)."""
    m = _A1.match(a1.strip()) if isinstance(a1, str) else None
    if not m or int(m.group(2)) < 1:
        return None
    return int(m.group(2)) - 1, col_index(m.group(1).upper()) - 1


def parse_range(a1: str) -> tuple[int, int, int, int] | None:
    """'A1:C3' (or a single cell, with or without a sheet prefix) -> grid (r1, c1, r2, c2), normalised."""
    if not isinstance(a1, str):
        return None
    text = a1.rsplit("!", 1)[-1]
    first, _, second = text.partition(":")
    a, b = parse_cell(first), parse_cell(second or first)
    if not a or not b:
        return None
    return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])


def cell_name(r: int, c: int) -> str:
    """Grid (r, c) -> Excel A1."""
    return f"{column_letter(c)}{r + 1}"


# ---------------------------------------------------------------------------
# Applying ops (store.ts applyOp / shiftAll)
# ---------------------------------------------------------------------------


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _for_each_formula(wb: dict, fn) -> None:
    for sheet in wb.get("sheets") or []:
        for row in sheet.get("rows") or []:
            for cell in (row or {}).get("cells") or []:
                if isinstance(cell, dict) and _is_formula(cell.get("value")):
                    cell["value"] = fn(cell["value"], sheet)


def shift_all(wb: dict, target: str, axis: Axis, at: int, delta: int) -> None:
    _for_each_formula(wb, lambda f, sheet: shift_formula(f, sheet.get("name", ""), target, axis, at, delta))
    sheet = next((s for s in wb.get("sheets") or [] if s.get("name") == target), None)
    if sheet is None:
        return

    def move(rng: str) -> str:
        return shift_formula(f"={rng}", target, target, axis, at, delta)[1:]

    sheet["validations"] = [v for v in ({**v, "range": move(v.get("range", ""))} for v in sheet.get("validations") or [])
                            if "#REF!" not in v["range"]]
    sheet["conditional_formats"] = [
        v for v in ({**v, "range": move(v.get("range", ""))} for v in sheet.get("conditional_formats") or [])
        if "#REF!" not in v["range"]]
    sheet["merges"] = [m for m in (move(m) for m in sheet.get("merges") or []) if "#REF!" not in m]
    if sheet.get("locked_ranges"):
        sheet["locked_ranges"] = [m for m in (move(m) for m in sheet["locked_ranges"]) if "#REF!" not in m]

    def move_ref(ref: str) -> str | None:
        bang = ref.rfind("!")
        name = "" if bang < 0 else re.sub(r"^'|'$", "", ref[:bang]).replace("''", "'")
        if name != target:
            return ref
        out = move(ref[bang + 1:])
        return None if "#REF!" in out else f"{ref[:bang]}!{out}"

    if wb.get("scenarios"):
        wb["scenarios"] = [
            {**sc, "values": {m: v for k, v in (sc.get("values") or {}).items() if (m := move_ref(k))}}
            for sc in wb["scenarios"]
        ]
    dashboard = wb.get("dashboard")
    if isinstance(dashboard, dict) and dashboard.get("tiles"):
        tiles = []
        for t in dashboard["tiles"]:
            t = dict(t)
            if t.get("ref"):
                t["ref"] = move_ref(t["ref"]) or "#REF!"
            if t.get("compare_ref"):
                t["compare_ref"] = move_ref(t["compare_ref"])
            tiles.append(t)
        wb["dashboard"] = {**dashboard, "tiles": tiles}
    if axis == "col" and sheet.get("filters"):
        sheet["filters"] = [
            {**f, "column": f["column"] + delta} if f["column"] >= at - 1 else f
            for f in sheet["filters"] if not (delta == -1 and f["column"] == at - 1)
        ]
    if axis == "col" and SHIFT_FREEZE and isinstance(sheet.get("freeze_columns"), int):
        # Inserting or deleting inside the frozen columns grows or shrinks them, like Excel.
        if at - 1 < sheet["freeze_columns"]:
            sheet["freeze_columns"] = max(0, sheet["freeze_columns"] + delta)
    charts = []
    for c in sheet.get("charts") or []:
        c = {**c, "categories": move(c.get("categories", "")),
             "series": [x for x in ({**x, "range": move(x.get("range", ""))} for x in c.get("series") or [])
                        if "#REF!" not in x["range"]]}
        if c["series"] and "#REF!" not in c["categories"]:
            charts.append(c)
    sheet["charts"] = charts


def apply_op(wb: dict, o: dict) -> None:
    """Apply one structural op in place (store.ts applyOp). Unknown sheets and count < 1 do nothing."""
    sheet = next((s for s in wb.get("sheets") or [] if s.get("name") == o["sheet"]), None)
    if sheet is None or o["count"] < 1:
        return
    rows = sheet.setdefault("rows", [])
    columns = sheet.setdefault("columns", [])
    at, count, name = o["at"], o["count"], sheet["name"]
    if o["op"] == "insert_rows":
        while len(rows) < at - 1:
            rows.append({"kind": "data", "cells": []})
        rows[at - 1:at - 1] = [{"kind": "data", "cells": []} for _ in range(count)]
        for _ in range(count):
            shift_all(wb, name, "row", at + 1, 1)
    elif o["op"] == "delete_rows":
        for r in range(at + count - 1, max(1, at) - 1, -1):
            if r - 1 >= len(rows):
                continue
            del rows[r - 1]
            shift_all(wb, name, "row", r + 1, -1)
    elif o["op"] == "insert_cols":
        while len(columns) < at:
            columns.append({"header": f"Column {len(columns) + 1}", "format": "text"})
        for _ in range(count):
            columns.insert(at, {"header": "New column", "format": "text"})
            for row in rows:
                cells = row.get("cells") or []
                if len(cells) > at:
                    cells.insert(at, {"value": None})
            shift_all(wb, name, "col", at + 1, 1)
    else:
        for c in range(at + count - 1, at - 1, -1):
            if c >= len(columns) or len(columns) <= 1:
                continue
            del columns[c]
            for row in rows:
                cells = row.get("cells") or []
                if len(cells) > c:
                    del cells[c]
            shift_all(wb, name, "col", c + 1, -1)


def apply_ops(wb: dict, ops: list[dict]) -> None:
    for o in ops:
        apply_op(wb, o)
