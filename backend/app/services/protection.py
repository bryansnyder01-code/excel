"""Locked sheets and locked ranges: checks run for everyone but the owner.

A sheet with `locked: true` is read-only as a whole; `locked_ranges` (A1 ranges such as
"A1:H1") are read-only cells. Grid row r is Excel row r+1 (row 1 is the header row), grid
column c is Excel column c+1. Workbooks are the raw dicts the live room and saved docs hold,
so the fields are read directly (they may be missing on older workbooks).

Every check returns None when the change is allowed, or a user-facing sentence.
"""

from __future__ import annotations

import json

from app.services.formula_shift import cell_name, parse_range

CHANGE_LOCKS = "Only the owner can change which cells are locked."


def _sheets(wb) -> list[dict]:
    return [s for s in (wb or {}).get("sheets") or [] if isinstance(s, dict)] if isinstance(wb, dict) else []


def _sheet(wb, name: str) -> dict | None:
    return next((s for s in _sheets(wb) if s.get("name") == name), None)


def _ranges(sheet: dict) -> list[tuple[int, int, int, int]]:
    out = []
    for text in sheet.get("locked_ranges") or []:
        rng = parse_range(text)
        if rng:
            out.append(rng)
    return out


def has_locks(wb) -> bool:
    return any(s.get("locked") or s.get("locked_ranges") for s in _sheets(wb))


def locked_sheet_message(name: str) -> str:
    return f'The sheet "{name}" is locked. Only the owner can change it.'


def locked_cell_message(name: str, r: int, c: int) -> str:
    return f"{cell_name(r, c)} on {name} is locked. Only the owner can change it."


def is_locked(sheet: dict, r: int, c: int) -> bool:
    if sheet.get("locked"):
        return True
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in _ranges(sheet))


def check_patch(wb: dict, cells: list[dict], headers: list[dict]) -> str | None:
    """A live patch: cells {sheet, r, c} and header changes {sheet, c} (the header is grid row 0)."""
    for ch in [*headers, *cells]:
        sheet = _sheet(wb, ch["sheet"])
        r = ch.get("r", 0)
        if sheet is not None and is_locked(sheet, r, ch["c"]):
            return locked_sheet_message(sheet["name"]) if sheet.get("locked") else \
                locked_cell_message(sheet["name"], r, ch["c"])
    return None


def check_ops(wb: dict, ops: list[dict]) -> str | None:
    """Structural ops: none on a locked sheet; no deleting rows/columns that hold locked cells; no
    inserting inside a locked range (moving a whole locked range is fine, its cells move intact).
    `wb` is the state before the ops; each op is checked against the state the previous ones leave."""
    from app.services.formula_shift import apply_op  # local: keeps the import graph one-way

    current = json.loads(json.dumps(wb))
    for o in ops:
        sheet = _sheet(current, o["sheet"])
        if sheet is None:
            continue
        if sheet.get("locked"):
            return locked_sheet_message(sheet["name"])
        rows = o["op"].endswith("_rows")
        lo, hi = o["at"], o["at"] + o["count"] - 1
        for r1, c1, r2, c2 in _ranges(sheet):
            first, last = (r1, r2) if rows else (c1, c2)
            if o["op"].startswith("delete"):
                hit = lo <= last and hi >= first
            else:
                hit = first < lo <= last
            if hit:
                what = "rows" if rows else "columns"
                verb = "deleted" if o["op"].startswith("delete") else "inserted inside"
                return (f"Those {what} can't be {verb}: {cell_name(r1, c1)}:{cell_name(r2, c2)} on "
                        f"{sheet['name']} is locked. Only the owner can change it.")
        apply_op(current, o)
    return None


def _cell(sheet: dict, r: int, c: int):
    """Comparable content of grid cell (r, c): header text for r == 0, else the cell minus default fields."""
    if r == 0:
        cols = sheet.get("columns") or []
        col = cols[c] if c < len(cols) and isinstance(cols[c], dict) else {}
        return col.get("header") or ""
    rows = sheet.get("rows") or []
    row = rows[r - 1] if r - 1 < len(rows) and isinstance(rows[r - 1], dict) else {}
    cells = row.get("cells") or []
    cell = cells[c] if c < len(cells) and isinstance(cells[c], dict) else {}
    return {k: v for k, v in cell.items() if v is not None and (v is not False or k == "value")}


def check_replace(old: dict, new) -> str | None:
    """A whole new workbook (live snapshot or synced PUT) in place of `old`: locked sheets must be identical,
    locked cells unchanged, and the locks themselves the same."""
    new_sheets = {s.get("name"): s for s in _sheets(new)}
    old_names = {s.get("name") for s in _sheets(old)}
    for s in _sheets(old):
        if not (s.get("locked") or s.get("locked_ranges")):
            continue
        n = new_sheets.get(s.get("name"))
        if n is None:
            return locked_sheet_message(s.get("name"))
        if bool(n.get("locked")) != bool(s.get("locked")) or (n.get("locked_ranges") or []) != (s.get("locked_ranges") or []):
            return CHANGE_LOCKS
        if s.get("locked"):
            if json.dumps(s, sort_keys=True) != json.dumps(n, sort_keys=True):
                return locked_sheet_message(s["name"])
            continue
        height = max(len(s.get("rows") or []), len(n.get("rows") or []))
        width = max(len(s.get("columns") or []), len(n.get("columns") or []),
                    *(len((row or {}).get("cells") or []) for row in (s.get("rows") or []) + (n.get("rows") or [])))
        for r1, c1, r2, c2 in _ranges(s):
            for r in range(r1, min(r2, height) + 1):
                for c in range(c1, min(c2, width - 1) + 1):
                    if _cell(s, r, c) != _cell(n, r, c):
                        return locked_cell_message(s["name"], r, c)
    for name, n in new_sheets.items():
        if name not in old_names and (n.get("locked") or n.get("locked_ranges")):
            return CHANGE_LOCKS
    return None
