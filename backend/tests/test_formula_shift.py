"""formula_shift: the Python port of the frontend's refs.ts shiftFormula and store.ts applyOp/shiftAll.
Expected values follow the TypeScript behaviour (checked against it while porting)."""

import copy

import pytest

from app.services.formula_shift import (apply_op, apply_ops, cell_name, col_index, column_letter, offset_formula,
                                        parse_cell, parse_range, shift_formula)


@pytest.mark.parametrize("formula, axis, at, delta, expected", [
    # rows: insert before `at` (Excel 1-based)
    ("=SUM(B2:B10)", "row", 5, 1, "=SUM(B2:B11)"),
    ("=SUM(B2:B10)", "row", 2, 1, "=SUM(B3:B11)"),
    ("=SUM(B2:B10)", "row", 11, 1, "=SUM(B2:B10)"),
    ("=B5*2", "row", 5, 1, "=B6*2"),
    ("=B5*2", "row", 6, 1, "=B5*2"),
    ("=$B$5+B$6+$C7", "row", 5, 1, "=$B$6+B$7+$C8"),
    # rows: delete `at`
    ("=B5*2", "row", 5, -1, "=#REF!*2"),
    ("=B6*2", "row", 5, -1, "=B5*2"),
    ("=SUM(B2:B10)", "row", 2, -1, "=SUM(B2:B9)"),
    ("=SUM(B2:B10)", "row", 10, -1, "=SUM(B2:B9)"),
    ("=SUM(B2:B10)", "row", 5, -1, "=SUM(B2:B9)"),
    ("=SUM(B5:B5)", "row", 5, -1, "=SUM(#REF!)"),
    ("=SUM($B$2:$B$10)", "row", 1, -1, "=SUM($B$1:$B$9)"),
    # columns
    ("=SUM(A1:C1)", "col", 2, 1, "=SUM(A1:D1)"),
    ("=A1+B1", "col", 1, 1, "=B1+C1"),
    ("=Z1", "col", 1, 1, "=AA1"),
    ("=AA1", "col", 1, -1, "=Z1"),
    ("=A1", "col", 1, -1, "=#REF!"),
    ("=SUM(A1:C1)", "col", 1, -1, "=SUM(A1:B1)"),
    ("=$C$3", "col", 2, -1, "=$B$3"),
    # quoted text, function names and non-references are left alone
    ('=IF(B5="B5",B5,"x B5")', "row", 1, 1, '=IF(B6="B5",B6,"x B5")'),
    ("=LOG10(B5)", "row", 1, 1, "=LOG10(B6)"),
    ("=B5B", "row", 1, 1, "=B5B"),
    ("=Sheet1.B5", "row", 1, 1, "=Sheet1.B5"),
])
def test_shift_formula_own_sheet(formula, axis, at, delta, expected):
    assert shift_formula(formula, "Data", "Data", axis, at, delta) == expected


def test_shift_formula_cross_sheet():
    # A formula on "Summary" pointing at "My Data" moves; its own-sheet refs don't.
    f = "='My Data'!B5+B5+Other!B5+'it''s'!B5"
    assert shift_formula(f, "Summary", "My Data", "row", 3, 1) == "='My Data'!B6+B5+Other!B5+'it''s'!B5"
    assert shift_formula(f, "Summary", "summary", "row", 3, 1) == "='My Data'!B5+B6+Other!B5+'it''s'!B5"
    assert shift_formula(f, "Summary", "it's", "row", 3, -1) == "='My Data'!B5+B5+Other!B5+'it''s'!B4"
    assert shift_formula("=SUM(Data!A1:A9)", "Summary", "data", "col", 1, -1) == "=SUM(#REF!)"  # like the TS: the prefix goes too
    assert shift_formula("=SUM(Data!A1:B9)", "Summary", "Data", "col", 1, -1) == "=SUM(Data!A1:A9)"


def test_offset_formula_and_helpers():
    assert offset_formula("=A1+$B$2+C$3", 1, 1) == "=B2+$B$2+D$3"
    assert offset_formula("=A1", -1, 0) == "=#REF!"
    assert col_index("A") == 1 and col_index("AA") == 27 and column_letter(0) == "A" and column_letter(27) == "AB"
    assert column_letter(-1) == ""
    assert parse_cell("B5") == (4, 1) and parse_cell("$b$1") == (0, 1) and parse_cell("B0") is None
    assert parse_cell("5B") is None and cell_name(4, 1) == "B5"
    assert parse_range("C3:A1") == (0, 0, 2, 2) and parse_range("'S'!B2") == (1, 1, 1, 1)


def workbook():
    return {
        "title": "T",
        "sheets": [
            {
                "name": "Data",
                "columns": [{"header": h, "format": "text"} for h in ("Item", "Q1", "Q2", "Total")],
                "rows": [
                    {"kind": "data", "cells": [{"value": "a"}, {"value": 1}, {"value": 2}, {"value": "=B2+C2"}]},
                    {"kind": "data", "cells": [{"value": "b"}, {"value": 3}, {"value": 4}, {"value": "=B3+C3"}]},
                    {"kind": "data", "cells": [{"value": "c"}, {"value": 5}, {"value": 6}, {"value": "=B4+C4"}]},
                    {"kind": "total", "cells": [{"value": "Total"}, {"value": "=SUM(B2:B4)"},
                                                {"value": "=SUM(C2:C4)"}, {"value": "=SUM($D$2:$D$4)"}]},
                ],
                "freeze_header": True,
                "freeze_columns": 1,
                "merges": ["A1:B1", "A3:D3"],
                "conditional_formats": [{"range": "D2:D4", "operator": "negative"}, {"range": "B3:B3", "operator": "negative"}],
                "validations": [{"range": "C2:C4", "type": "whole"}],
                "filters": [{"column": 1, "values": ["1"]}, {"column": 3, "values": ["x"]}],
                "charts": [{"type": "column", "title": "Q", "categories": "A2:A4",
                            "series": [{"name": "Q1", "range": "B2:B4"}, {"name": "Q2", "range": "C2:C4"}]}],
                "locked_ranges": ["A1:D1", "B3"],
            },
            {
                "name": "Summary Sheet",
                "columns": [{"header": "x", "format": "text"}],
                "rows": [{"kind": "data", "cells": [{"value": "=Data!D5*2"}, {"value": "=SUM(Data!B2:B4)"}]}],
            },
        ],
        "scenarios": [{"name": "Best", "values": {"Data!B3": 9, "'Summary Sheet'!A2": 1, "B3": 4}}],
        "dashboard": {"tiles": [{"ref": "Data!D5", "compare_ref": "Data!B3"}]},
    }


def test_insert_rows_moves_everything():
    wb = workbook()
    apply_op(wb, {"op": "insert_rows", "sheet": "Data", "at": 2, "count": 2})  # before grid row 2 (Excel 3)
    s = wb["sheets"][0]
    assert len(s["rows"]) == 6 and s["rows"][1] == {"kind": "data", "cells": []} and s["rows"][2]["cells"] == []
    assert s["rows"][0]["cells"][3]["value"] == "=B2+C2"
    assert s["rows"][3]["cells"][3]["value"] == "=B5+C5"
    assert [c["value"] for c in s["rows"][5]["cells"][1:]] == ["=SUM(B2:B6)", "=SUM(C2:C6)", "=SUM($D$2:$D$6)"]
    assert wb["sheets"][1]["rows"][0]["cells"] == [{"value": "=Data!D7*2"}, {"value": "=SUM(Data!B2:B6)"}]
    assert s["merges"] == ["A1:B1", "A5:D5"]
    assert [c["range"] for c in s["conditional_formats"]] == ["D2:D6", "B5:B5"]
    assert s["validations"][0]["range"] == "C2:C6"
    assert s["charts"][0]["categories"] == "A2:A6" and s["charts"][0]["series"][1]["range"] == "C2:C6"
    assert s["locked_ranges"] == ["A1:D1", "B5"]
    assert wb["scenarios"][0]["values"] == {"Data!B5": 9, "'Summary Sheet'!A2": 1, "B3": 4}
    assert wb["dashboard"]["tiles"] == [{"ref": "Data!D7", "compare_ref": "Data!B5"}]
    assert s["filters"] == [{"column": 1, "values": ["1"]}, {"column": 3, "values": ["x"]}]
    assert s["freeze_columns"] == 1


def test_insert_rows_past_the_end_pads():
    wb = workbook()
    apply_op(wb, {"op": "insert_rows", "sheet": "Data", "at": 8, "count": 1})
    rows = wb["sheets"][0]["rows"]
    assert len(rows) == 8 and rows[4:] == [{"kind": "data", "cells": []}] * 4


def test_delete_rows_shrinks_ranges_and_drops_dead_things():
    wb = workbook()
    apply_op(wb, {"op": "delete_rows", "sheet": "Data", "at": 2, "count": 1})  # grid row 2 = Excel row 3
    s = wb["sheets"][0]
    assert [r["cells"][0]["value"] for r in s["rows"]] == ["a", "c", "Total"]
    assert s["rows"][1]["cells"][3]["value"] == "=B3+C3"
    assert [c["value"] for c in s["rows"][2]["cells"][1:]] == ["=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM($D$2:$D$3)"]
    assert wb["sheets"][1]["rows"][0]["cells"] == [{"value": "=Data!D4*2"}, {"value": "=SUM(Data!B2:B3)"}]
    assert s["merges"] == ["A1:B1"]  # A3:D3 was the deleted row
    assert [c["range"] for c in s["conditional_formats"]] == ["D2:D3"]  # B3:B3 went
    assert s["locked_ranges"] == ["A1:D1"]
    assert wb["scenarios"][0]["values"] == {"'Summary Sheet'!A2": 1, "B3": 4}
    assert wb["dashboard"]["tiles"] == [{"ref": "Data!D4", "compare_ref": None}]

    # A formula pointing only at a deleted cell becomes #REF!; rows past the end are skipped.
    wb = workbook()
    apply_op(wb, {"op": "delete_rows", "sheet": "Data", "at": 4, "count": 3})
    s = wb["sheets"][0]
    assert len(s["rows"]) == 3
    assert wb["sheets"][1]["rows"][0]["cells"][0]["value"] == "=#REF!*2"
    assert wb["dashboard"]["tiles"][0]["ref"] == "#REF!"


def test_insert_and_delete_columns(monkeypatch):
    monkeypatch.setattr("app.services.formula_shift.SHIFT_FREEZE", True)
    wb = workbook()
    apply_op(wb, {"op": "insert_cols", "sheet": "Data", "at": 1, "count": 1})  # before B
    s = wb["sheets"][0]
    assert [c["header"] for c in s["columns"]] == ["Item", "New column", "Q1", "Q2", "Total"]
    assert s["rows"][0]["cells"][1] == {"value": None} and s["rows"][0]["cells"][4]["value"] == "=C2+D2"
    assert s["rows"][3]["cells"][4]["value"] == "=SUM($E$2:$E$4)"
    assert s["merges"] == ["A1:C1", "A3:E3"]
    assert s["filters"] == [{"column": 2, "values": ["1"]}, {"column": 4, "values": ["x"]}]
    assert s["charts"][0]["series"][0]["range"] == "C2:C4" and s["charts"][0]["categories"] == "A2:A4"
    assert s["locked_ranges"] == ["A1:E1", "C3"]
    assert s["freeze_columns"] == 1  # inserted right of the frozen column
    assert wb["sheets"][1]["rows"][0]["cells"][1]["value"] == "=SUM(Data!C2:C4)"

    wb = workbook()
    apply_op(wb, {"op": "insert_cols", "sheet": "Data", "at": 0, "count": 2})
    s = wb["sheets"][0]
    assert s["freeze_columns"] == 3 and s["columns"][0]["header"] == "New column"

    wb = workbook()
    apply_op(wb, {"op": "delete_cols", "sheet": "Data", "at": 1, "count": 2})  # B and C
    s = wb["sheets"][0]
    assert [c["header"] for c in s["columns"]] == ["Item", "Total"]
    assert s["rows"][0]["cells"] == [{"value": "a"}, {"value": "=#REF!+#REF!"}]
    assert s["rows"][3]["cells"][1]["value"] == "=SUM($B$2:$B$4)"
    assert s["filters"] == [{"column": 1, "values": ["x"]}]
    assert s["charts"] == []  # both series were in deleted columns
    assert s["validations"] == []
    assert s["merges"] == ["A1:A1", "A3:B3"]  # A1:B1 shrinks to a one-cell range, as in the TS
    assert s["locked_ranges"] == ["A1:B1"]


def test_delete_columns_keeps_one_and_ignores_missing():
    wb = workbook()
    apply_op(wb, {"op": "delete_cols", "sheet": "Data", "at": 0, "count": 10})
    assert [c["header"] for c in wb["sheets"][0]["columns"]] == ["Item"]
    before = copy.deepcopy(wb)
    apply_ops(wb, [{"op": "delete_rows", "sheet": "Nope", "at": 1, "count": 1},
                   {"op": "insert_rows", "sheet": "Data", "at": 1, "count": 0}])
    assert wb == before


def test_insert_cols_past_the_end_pads_columns():
    wb = workbook()
    apply_op(wb, {"op": "insert_cols", "sheet": "Data", "at": 6, "count": 1})
    assert [c["header"] for c in wb["sheets"][0]["columns"]][4:] == ["Column 5", "Column 6", "New column"]


def test_freeze_columns_grow_when_inserting_inside_them():
    wb = workbook()
    apply_op(wb, {"op": "insert_cols", "sheet": "Data", "at": 0, "count": 2})
    assert wb["sheets"][0]["freeze_columns"] == 3
