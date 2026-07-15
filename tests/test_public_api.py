"""Tests for the stable, documented SheetProof Python API."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from sheetproof import ReviewResult, __version__, compare_workbooks
from sheetproof._version import __version__ as canonical_version
from tests.conftest import WorkbookFactory


def _before(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = "before"


def _after(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = "after"


def test_public_version_uses_the_canonical_version_module() -> None:
    assert __version__ == canonical_version == "0.1.0"


def test_compare_workbooks_returns_versioned_review_result(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    before = workbook_factory("public-api-before.xlsx", _before)
    after = workbook_factory("public-api-after.xlsx", _after)
    config = tmp_path / "sheetproof.yml"
    config.write_text("rules: []\n", encoding="utf-8")

    result = compare_workbooks(before, after, config)

    assert isinstance(result, ReviewResult)
    assert result.tool_version == __version__
    assert result.before_file.path == str(before.resolve())
    assert result.after_file.path == str(after.resolve())
    assert result.summary.changed_cells == 1
