"""Workbook-level bounds for formula reference expansion."""

from __future__ import annotations

from openpyxl import Workbook

from sheetproof.diff.formula_diff import FormulaDiffer
from sheetproof.diff.workbook_diff import WorkbookDiffer
from sheetproof.parser.workbook import WorkbookParser
from tests.conftest import WorkbookConfigurator, WorkbookFactory


def _formula_ranges(size: int) -> WorkbookConfigurator:
    def configure(workbook: Workbook) -> None:
        sheet = workbook.active
        sheet.title = "Calc"
        for row in range(1, 11):
            sheet[f"A{row}"] = row
        for row in range(1, 4):
            sheet[f"B{row}"] = f"=SUM(A1:A{size})"

    return configure


def test_workbook_compare_shares_a_total_formula_expansion_budget(
    workbook_factory: WorkbookFactory,
) -> None:
    parser = WorkbookParser()
    before = parser.parse(workbook_factory("budget-before.xlsx", _formula_ranges(10)))
    after = parser.parse(workbook_factory("budget-after.xlsx", _formula_ranges(9)))
    differ = WorkbookDiffer()
    differ.formula_differ = FormulaDiffer(
        expansion_limit=100,
        total_expansion_limit=20,
    )

    _, _, changes = differ.compare(before, after)

    assert len(changes) == 3
    assert changes[0].change_type == "formula_range_reduced"
    assert changes[0].supported_analysis is True
    for change in changes[1:]:
        assert change.change_type == "formula_changed"
        assert change.supported_analysis is False
        assert change.evidence["reference_expansion_truncated"] is True
        assert "reference_expansion_truncated" in change.evidence["unsupported_reasons"]
