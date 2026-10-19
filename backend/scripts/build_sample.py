"""Build an .xlsx from a workbook JSON file, without calling the LLM.

    python -m scripts.build_sample samples/small_business_budget.json out.xlsx
"""

import json
import sys
from pathlib import Path

from app.schemas import WorkbookSpec
from app.services.formula_validator import validate_workbook
from app.services.workbook_builder import build_workbook


def main(src: str, dest: str) -> None:
    spec = WorkbookSpec.model_validate(json.loads(Path(src).read_text()))
    for issue in validate_workbook(spec):
        flag = "fixed" if issue.fixed else issue.severity
        print(f"[{flag}] {issue.sheet or ''} {issue.cell or ''}: {issue.message}")
    build_workbook(spec).save(dest)
    print(f"Wrote {dest} ({len(spec.sheets)} sheets)")


if __name__ == "__main__":
    main(*sys.argv[1:3])
