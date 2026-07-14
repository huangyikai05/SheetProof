"""Regression tests for semantic cell-style comparison."""

from __future__ import annotations

from collections.abc import Iterator

import pytest
from openpyxl import Workbook
from openpyxl.styles import Color, Font

from sheetproof.diff.cell_diff import CellDiffer
from sheetproof.diff.workbook_diff import WorkbookDiffer
from sheetproof.models import CellKind, CellSnapshot, RiskLevel, SheetSnapshot, StyleSummary
from sheetproof.parser.workbook import WorkbookParser
from tests.conftest import WorkbookConfigurator, WorkbookFactory


def _style(style_id: int, **updates: str) -> StyleSummary:
    values = {
        "style_id": style_id,
        "number_format": "General",
        "font": "font:base",
        "fill": "fill:none",
        "border": "border:none",
        "alignment": "alignment:default",
        "protection": "protection:locked",
    }
    values.update(updates)
    return StyleSummary(**values)


def _cell(
    value: str | int,
    kind: CellKind,
    data_type: str,
    style: StyleSummary,
    *,
    coordinate: str = "A1",
    is_date: bool = False,
) -> CellSnapshot:
    return CellSnapshot(
        coordinate=coordinate,
        value=value,
        kind=kind,
        data_type=data_type,
        is_date=is_date,
        style=style,
    )


def _sheet(*cells: CellSnapshot) -> SheetSnapshot:
    return SheetSnapshot(
        name="Sheet",
        index=0,
        state="visible",
        cells={cell.coordinate: cell for cell in cells},
        hidden_rows=[],
        hidden_columns=[],
        merged_cells=[],
        data_validations=[],
        freeze_panes=None,
        tables=[],
    )


@pytest.fixture(params=[
    ("text", CellKind.TEXT, "s", False),
    (42, CellKind.NUMBER, "n", False),
    ("2026-07-14", CellKind.DATE, "d", True),
    ("#DIV/0!", CellKind.ERROR, "e", False),
])
def nonempty_cell_case(
    request: pytest.FixtureRequest,
) -> Iterator[tuple[str | int, CellKind, str, bool]]:
    yield request.param


def test_style_id_is_not_part_of_cell_style_semantics(
    nonempty_cell_case: tuple[str | int, CellKind, str, bool],
) -> None:
    value, kind, data_type, is_date = nonempty_cell_case
    before = _cell(value, kind, data_type, _style(3), is_date=is_date)
    after = _cell(value, kind, data_type, _style(91), is_date=is_date)

    assert CellDiffer().compare(_sheet(before), _sheet(after)) == []


def test_nonempty_style_only_changes_are_classified_as_style_changes(
    nonempty_cell_case: tuple[str | int, CellKind, str, bool],
) -> None:
    value, kind, data_type, is_date = nonempty_cell_case
    before = _cell(value, kind, data_type, _style(3), is_date=is_date)
    after = _cell(
        value,
        kind,
        data_type,
        _style(91, font="font:bold"),
        is_date=is_date,
    )

    changes = CellDiffer().compare(_sheet(before), _sheet(after))

    assert len(changes) == 1
    assert changes[0].change_type == "style_changed"
    assert changes[0].risk_level is RiskLevel.LOW
    assert changes[0].evidence["style_changed"] is True


@pytest.mark.parametrize(
    ("field", "changed_value"),
    [
        ("number_format", "0.00%"),
        ("font", "font:bold"),
        ("fill", "fill:yellow"),
        ("border", "border:thin"),
        ("alignment", "alignment:center"),
        ("protection", "protection:unlocked"),
    ],
)
def test_every_semantic_style_field_is_compared(field: str, changed_value: str) -> None:
    before = _cell("same", CellKind.TEXT, "s", _style(1))
    after = _cell("same", CellKind.TEXT, "s", _style(2, **{field: changed_value}))

    changes = CellDiffer().compare(_sheet(before), _sheet(after))

    assert [change.change_type for change in changes] == ["style_changed"]


def test_content_change_takes_priority_over_simultaneous_style_change() -> None:
    before = _cell(10, CellKind.NUMBER, "n", _style(1))
    after = _cell(20, CellKind.NUMBER, "n", _style(2, fill="fill:red"))

    changes = CellDiffer().compare(_sheet(before), _sheet(after))

    assert len(changes) == 1
    assert changes[0].change_type == "numeric_value_changed"
    assert changes[0].evidence["style_changed"] is True


def test_reordered_workbook_style_table_does_not_create_cell_changes() -> None:
    before = _sheet(
        _cell("bold", CellKind.TEXT, "s", _style(1, font="font:bold")),
        _cell(
            "percent",
            CellKind.TEXT,
            "s",
            _style(2, number_format="0.00%"),
            coordinate="B1",
        ),
    )
    after = _sheet(
        _cell("bold", CellKind.TEXT, "s", _style(2, font="font:bold")),
        _cell(
            "percent",
            CellKind.TEXT,
            "s",
            _style(1, number_format="0.00%"),
            coordinate="B1",
        ),
    )

    assert CellDiffer().compare(before, after) == []


def test_theme_color_changes_are_preserved_in_parsed_style_semantics(
    workbook_factory: WorkbookFactory,
) -> None:
    def configure(theme: int) -> WorkbookConfigurator:
        def apply(workbook: Workbook) -> None:
            workbook.active["A1"] = "same"
            workbook.active["A1"].font = Font(color=Color(theme=theme))

        return apply

    parser = WorkbookParser()
    before = parser.parse(workbook_factory("theme-before.xlsx", configure(1)))
    after = parser.parse(workbook_factory("theme-after.xlsx", configure(2)))

    _, changes, _ = WorkbookDiffer().compare(before, after)

    assert [change.change_type for change in changes] == ["style_changed"]
    change = changes[0]
    assert change.before is not None
    assert change.after is not None
    assert '"type": "theme"' in change.before.style.font
    assert '"value": 1' in change.before.style.font
    assert '"value": 2' in change.after.style.font
