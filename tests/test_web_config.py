from __future__ import annotations

import tomllib
from pathlib import Path

from openpyxl import Workbook

from sheetproof.services.review_service import ReviewService
from tests.conftest import WorkbookFactory
from web.app import _formula_overwrite_rows


def test_streamlit_upload_limit_matches_application_limit() -> None:
    config_path = Path(__file__).parents[1] / ".streamlit" / "config.toml"
    config = tomllib.loads(config_path.read_text(encoding="utf-8"))

    assert config["server"]["maxUploadSize"] == 25
    assert config["server"]["maxMessageSize"] == 25


def test_streamlit_overwrite_rows_show_canonical_replacement_evidence(
    workbook_factory: WorkbookFactory,
) -> None:
    def before(workbook: Workbook) -> None:
        workbook.active["A1"] = "=B1*2"
        workbook.active["B1"] = 21

    def after(workbook: Workbook) -> None:
        workbook.active["A1"] = "manual"
        workbook.active["B1"] = 21

    result = ReviewService().review(
        workbook_factory("web-before.xlsx", before),
        workbook_factory("web-after.xlsx", after),
    )

    rows = _formula_overwrite_rows(result)

    assert rows[0]["replacement_value"] == "manual"
    assert rows[0]["replacement_kind"] == "text"
    assert str(rows[0]["manual_review_recommendation"]).startswith("Manually verify")
