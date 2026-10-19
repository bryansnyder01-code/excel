from app.schemas import WorkbookSpec
from app.services.formula_validator import has_errors, iter_refs, validate_workbook


def sheet(name, rows, ncols=3):
    return {"name": name, "columns": [{"header": f"H{i}"} for i in range(ncols)],
            "rows": [{"cells": [{"value": v} for v in r]} for r in rows]}


def wb(*sheets):
    return WorkbookSpec.model_validate({"title": "t", "sheets": list(sheets)})


def test_refs_ignore_function_names_and_strings():
    refs = iter_refs('=LOG10(A2)+SUM(\'My Sheet\'!B2:C9)&"see D4"')
    assert [(r.sheet, r.min_col, r.min_row, r.max_col, r.max_row) for r in refs] == [
        (None, 1, 2, 1, 2), ("My Sheet", 2, 2, 3, 9)]


def test_sample_budget_is_clean(budget_spec):
    issues = validate_workbook(budget_spec)
    assert not has_errors(issues)


def test_unknown_sheet_is_error():
    issues = validate_workbook(wb(sheet("A", [[1, "=Missing!A2"]])))
    assert has_errors(issues)
    assert "does not exist" in issues[0].message


def test_circular_reference_is_error():
    issues = validate_workbook(wb(sheet("A", [[1, 2, "=SUM(A2:C2)"]])))
    assert any("circular" in i.message for i in issues)


def test_literal_divide_by_zero_is_error():
    assert has_errors(validate_workbook(wb(sheet("A", [[1, "=A2/0"]]))))


def test_cell_division_is_wrapped_in_iferror():
    spec = wb(sheet("A", [[10, 0, "=A2/B2"]]))
    issues = validate_workbook(spec)
    assert spec.sheets[0].rows[0].cells[2].value == "=IFERROR(A2/B2,0)"
    assert issues[0].fixed and not has_errors(issues)


def test_constant_division_left_alone():
    spec = wb(sheet("A", [[10, "=A2/100"]]))
    assert validate_workbook(spec) == []
    assert spec.sheets[0].rows[0].cells[1].value == "=A2/100"


def test_bad_sheet_name_renamed_and_formulas_follow():
    spec = wb(sheet("Q1/Q2 Sales", [[5]]), sheet("Sum", [["='Q1/Q2 Sales'!A2"]]))
    issues = validate_workbook(spec)
    assert spec.sheets[0].name == "Q1-Q2 Sales"
    assert spec.sheets[1].rows[0].cells[0].value == "='Q1-Q2 Sales'!A2"
    assert not has_errors(issues)


def test_duplicate_sheet_names_made_unique():
    spec = wb(sheet("Data", [[1]]), sheet("data", [[2]]))
    validate_workbook(spec)
    assert [s.name for s in spec.sheets] == ["Data", "data (2)"]


def test_unknown_function_and_unbalanced_parens():
    issues = validate_workbook(wb(sheet("A", [[1, "=FOOBAR(A2)", "=SUM(A2"]])))
    msgs = " ".join(i.message for i in issues)
    assert "FOOBAR is not a recognised" in msgs and "unclosed parenthesis" in msgs
