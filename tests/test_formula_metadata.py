"""Array and data-table formula metadata regression tests."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from sheetproof.diff.workbook_diff import WorkbookDiffer
from sheetproof.models import CellKind, RiskLevel
from sheetproof.parser.formulas import FormulaParser
from sheetproof.parser.workbook import WorkbookParser
from tests.conftest import WorkbookConfigurator, WorkbookFactory


def _array_formula(ref: str) -> WorkbookConfigurator:
    def configure(workbook: Workbook) -> None:
        sheet = workbook.active
        sheet.title = "Array"
        sheet["B1"] = 1
        sheet["B2"] = 2
        sheet["B3"] = 3
        sheet["A1"] = ArrayFormula(ref=ref, text="=SUM(B1:B3)")

    return configure


def test_array_formula_parse_is_stable_and_identical_diff_is_empty(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory("array-stable.xlsx", _array_formula("A1:A3"))
    parser = WorkbookParser()

    first = parser.parse(path)
    second = parser.parse(path)
    cell = first.sheet_map()["Array"].cells["A1"]

    assert first == second
    assert cell.kind is CellKind.FORMULA
    assert cell.value == "=SUM(B1:B3)"
    assert cell.formula == "=SUM(B1:B3)"
    assert cell.formula_attributes == {"kind": "array", "ref": "A1:A3"}
    assert WorkbookDiffer().compare(first, second) == ([], [], [])


def test_array_formula_reference_change_is_a_high_risk_formula_change(
    workbook_factory: WorkbookFactory,
) -> None:
    parser = WorkbookParser()
    before = parser.parse(
        workbook_factory("array-before.xlsx", _array_formula("A1:A3"))
    )
    after = parser.parse(
        workbook_factory("array-after.xlsx", _array_formula("A1:A4"))
    )

    _, _, formulas = WorkbookDiffer().compare(before, after)

    assert len(formulas) == 1
    change = formulas[0]
    assert change.location == "Array!A1"
    assert change.change_type == "formula_changed"
    assert change.risk_level is RiskLevel.HIGH
    assert change.high_impact is True
    assert change.before_formula == change.after_formula == "=SUM(B1:B3)"
    assert change.evidence["before_formula_attributes"] == {
        "kind": "array",
        "ref": "A1:A3",
    }
    assert change.evidence["after_formula_attributes"] == {
        "kind": "array",
        "ref": "A1:A4",
    }


def test_data_table_formula_has_stable_unsupported_metadata(
    workbook_factory: WorkbookFactory,
) -> None:
    def configure(workbook: Workbook) -> None:
        sheet = workbook.active
        sheet.title = "Table"
        sheet["A1"] = DataTableFormula(
            ref="A1:B2",
            ca=True,
            dt2D=True,
            r1="C1",
            r2="D1",
            del1=True,
        )

    path = workbook_factory("data-table.xlsx", configure)
    parser = WorkbookParser()

    first = parser.parse(path)
    second = parser.parse(path)
    cell = first.sheet_map()["Table"].cells["A1"]
    analysis = FormulaParser().parse(cell.formula or "")

    assert first == second
    assert cell.kind is CellKind.FORMULA
    assert cell.formula == "=UNSUPPORTED_DATA_TABLE()"
    assert cell.formula_attributes == {
        "kind": "data_table",
        "ref": "A1:B2",
        "ca": True,
        "dt2D": True,
        "dtr": False,
        "r1": "C1",
        "r2": "D1",
        "del1": True,
        "del2": False,
    }
    assert analysis.supported is False
    assert analysis.unsupported_reasons == (
        "unsupported_function:UNSUPPORTED_DATA_TABLE",
    )
    assert WorkbookDiffer().compare(first, second) == ([], [], [])
