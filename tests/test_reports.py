"""Canonical report serialization and HTML escaping tests."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from sheetproof.exceptions import ReportGenerationError
from sheetproof.reports.html_report import render_html
from sheetproof.reports.json_report import render_json
from sheetproof.services.review_service import ReviewService
from tests.conftest import WorkbookFactory


def _plain(workbook: Workbook) -> None:
    workbook.active.title = "审计"
    workbook.active["A1"] = "plain"


def _hostile(workbook: Workbook) -> None:
    workbook.active.title = "审计"
    workbook.active["A1"] = '<script>alert("sheet")</script>'


def _formula_before(workbook: Workbook) -> None:
    workbook.active["A1"] = "=B1*2"
    workbook.active["B1"] = 21


def _formula_after(workbook: Workbook) -> None:
    workbook.active["A1"] = 42
    workbook.active["B1"] = 21


def _dependency_before(workbook: Workbook) -> None:
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["B1"] = 21
    calc = workbook.create_sheet("Calc")
    calc["A1"] = "=Inputs!B1*2"
    output = workbook.create_sheet("Output")
    output["A1"] = "=Calc!A1"


def _dependency_after(workbook: Workbook) -> None:
    _dependency_before(workbook)
    workbook["Calc"]["A1"] = 42


def test_json_report_preserves_unicode_and_canonical_shape(
    workbook_factory: WorkbookFactory,
) -> None:
    result = ReviewService().review(
        workbook_factory("unicode-before.xlsx", _plain),
        workbook_factory("unicode-after.xlsx", _hostile),
    )

    text = render_json(result)
    payload = json.loads(text)

    assert "审计" in text
    assert payload["summary"]["changed_cells"] == 1
    assert payload["cell_changes"][0]["after"]["value"] == '<script>alert("sheet")</script>'
    assert set(payload) >= {
        "summary",
        "structure_changes",
        "cell_changes",
        "formula_changes",
        "dependency_impacts",
        "rule_results",
        "risk_factors",
        "limitations",
        "errors",
    }


def test_html_report_escapes_workbook_controlled_text(
    workbook_factory: WorkbookFactory,
) -> None:
    result = ReviewService().review(
        workbook_factory("escape-before.xlsx", _plain),
        workbook_factory("escape-after.xlsx", _hostile),
    )

    html = render_html(result)

    assert '<script>alert("sheet")</script>' not in html
    assert "&lt;script&gt;" in html
    assert "alert" in html


def test_json_report_rejects_non_standard_non_finite_numbers(
    workbook_factory: WorkbookFactory,
) -> None:
    result = ReviewService().review(
        workbook_factory("finite-before.xlsx", _plain),
        workbook_factory("finite-after.xlsx", _hostile),
    )
    assert result.cell_changes[0].after is not None
    result.cell_changes[0].after.value = float("inf")

    with pytest.raises(ReportGenerationError, match="Out of range float values"):
        render_json(result)


def test_reports_expose_formula_overwrite_replacement_and_review_guidance(
    workbook_factory: WorkbookFactory,
) -> None:
    result = ReviewService().review(
        workbook_factory("formula-before.xlsx", _formula_before),
        workbook_factory("formula-after.xlsx", _formula_after),
    )

    payload = json.loads(render_json(result))
    change = payload["formula_changes"][0]
    assert change["replacement_value"] == 42
    assert change["replacement_kind"] == "number"
    assert change["manual_review_recommendation"].startswith("Manually verify")

    html = render_html(result)
    assert "Replacement value" in html
    assert "Replacement kind" in html
    assert "Manual review recommendation" in html
    assert change["manual_review_recommendation"] in html


def test_html_dependency_summary_exposes_critical_and_before_graph_evidence(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    config = tmp_path / "sheetproof.yml"
    config.write_text('critical_cells:\n  - "Output!A1"\n', encoding="utf-8")
    result = ReviewService().review(
        workbook_factory("dependency-before.xlsx", _dependency_before),
        workbook_factory("dependency-after.xlsx", _dependency_after),
        config,
    )

    html = render_html(result)

    assert "Critical cells" in html
    assert "Output!A1" in html
    assert "direct_upstream_before" in html
    assert "Inputs!B1" in html
