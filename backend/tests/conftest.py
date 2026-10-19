import json
from pathlib import Path

import pytest

from app.schemas import WorkbookSpec

SAMPLES = Path(__file__).parent.parent / "samples"


@pytest.fixture
def budget_spec() -> WorkbookSpec:
    return WorkbookSpec.model_validate(json.loads((SAMPLES / "small_business_budget.json").read_text()))


@pytest.fixture(autouse=True)
def tmp_data_dir(tmp_path, monkeypatch):
    """Keep the accounts database out of the repo: every test gets its own data_dir."""
    from app.config import get_settings

    monkeypatch.setattr(get_settings(), "data_dir", str(tmp_path / "data"))
    monkeypatch.setattr(get_settings(), "background_jobs", False)  # no once-a-minute loop in tests
