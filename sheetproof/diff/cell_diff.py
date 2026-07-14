"""Semantic cell-level workbook differences."""

from __future__ import annotations

from pydantic import JsonValue

from sheetproof.models import (
    CellChange,
    CellKind,
    CellSnapshot,
    RiskLevel,
    SheetSnapshot,
    StyleSummary,
)
from sheetproof.parser.formulas import qualify_cell


class CellDiffer:
    """Classify cell changes by business meaning instead of raw serialization."""

    def compare(self, before: SheetSnapshot, after: SheetSnapshot) -> list[CellChange]:
        if before.name != after.name:
            raise ValueError("CellDiffer requires snapshots of the same worksheet")
        changes: list[CellChange] = []
        for coordinate in sorted(set(before.cells) | set(after.cells), key=_coordinate_sort_key):
            before_cell = before.cells.get(coordinate)
            after_cell = after.cells.get(coordinate)
            if _cells_equal(before_cell, after_cell):
                continue
            change_type, risk_level, description = _classify_cell_change(
                before_cell,
                after_cell,
            )
            location = qualify_cell(before.name, coordinate)
            evidence: dict[str, JsonValue] = {
                "before_kind": _kind_value(before_cell),
                "after_kind": _kind_value(after_cell),
                "before_value": before_cell.value if before_cell is not None else None,
                "after_value": after_cell.value if after_cell is not None else None,
                "before_data_type": before_cell.data_type if before_cell is not None else None,
                "after_data_type": after_cell.data_type if after_cell is not None else None,
            }
            if before_cell is not None and after_cell is not None:
                if not _semantic_styles_equal(before_cell.style, after_cell.style):
                    evidence["style_changed"] = True
                if not _same_scalar(before_cell.cached_value, after_cell.cached_value):
                    evidence["cached_value_changed"] = True
            changes.append(
                CellChange(
                    change_type=change_type,
                    risk_level=risk_level,
                    sheet=before.name,
                    coordinate=coordinate,
                    location=location,
                    before=before_cell,
                    after=after_cell,
                    description=f"{location}: {description}",
                    evidence=evidence,
                )
            )
        return changes


def _classify_cell_change(
    before: CellSnapshot | None,
    after: CellSnapshot | None,
) -> tuple[str, RiskLevel, str]:
    before_blank = _is_blank(before)
    after_blank = _is_blank(after)
    before_formula = before is not None and before.formula is not None
    after_formula = after is not None and after.formula is not None

    if before_blank and after_blank:
        if before is not None and after is not None and _cell_types_differ(before, after):
            return "data_type_changed", RiskLevel.MEDIUM, "cell data type changed"
        return "style_changed", RiskLevel.LOW, "blank cell style changed"

    if before_formula and after_formula and before is not None and after is not None:
        if before.formula != after.formula:
            return "formula_changed", RiskLevel.MEDIUM, "formula changed"
        if before.formula_attributes != after.formula_attributes:
            return "formula_changed", RiskLevel.HIGH, "array or data-table metadata changed"
        if not _same_scalar(before.cached_value, after.cached_value):
            return "cached_value_changed", RiskLevel.LOW, "cached formula value changed"
    if before_formula and not after_formula:
        if after_blank:
            return "formula_deleted", RiskLevel.HIGH, "formula was deleted"
        return "formula_overwritten", RiskLevel.HIGH, "formula was replaced by a fixed value"
    if not before_formula and after_formula:
        if before_blank:
            return "blank_to_formula", RiskLevel.MEDIUM, "blank cell now contains a formula"
        return "fixed_value_to_formula", RiskLevel.MEDIUM, "fixed value was replaced by a formula"

    if before_blank and not after_blank:
        if after is not None and after.kind is CellKind.NUMBER:
            return "blank_to_number", RiskLevel.LOW, "blank cell now contains a number"
        if after is not None and after.kind is CellKind.TEXT:
            return "blank_to_text", RiskLevel.LOW, "blank cell now contains text"
        return "cell_populated", RiskLevel.LOW, "blank cell now contains a value"
    if not before_blank and after_blank:
        return "cell_cleared", RiskLevel.MEDIUM, "cell was cleared"

    if before is not None and after is not None:
        if _cell_types_differ(before, after):
            return "data_type_changed", RiskLevel.MEDIUM, "cell data type changed"
        if not _same_scalar(before.value, after.value):
            if before.kind is CellKind.DATE:
                return "date_changed", RiskLevel.MEDIUM, "date value changed"
            if before.kind is CellKind.ERROR:
                return "error_value_changed", RiskLevel.HIGH, "error value changed"
            if before.kind is CellKind.NUMBER:
                return "numeric_value_changed", RiskLevel.MEDIUM, "numeric value changed"
            if before.kind is CellKind.TEXT:
                return "text_changed", RiskLevel.LOW, "text changed"
            return "value_changed", RiskLevel.MEDIUM, "cell value changed"
        if not _semantic_styles_equal(before.style, after.style):
            return "style_changed", RiskLevel.LOW, "cell style changed"
    return "value_changed", RiskLevel.MEDIUM, "cell content changed"


def _cells_equal(before: CellSnapshot | None, after: CellSnapshot | None) -> bool:
    if before is None and after is None:
        return True
    if before is None or after is None:
        return False
    return (
        _same_scalar(before.value, after.value)
        and before.formula == after.formula
        and before.formula_attributes == after.formula_attributes
        and _same_scalar(before.cached_value, after.cached_value)
        and not _cell_types_differ(before, after)
        and _semantic_styles_equal(before.style, after.style)
    )


def _is_blank(cell: CellSnapshot | None) -> bool:
    return cell is None or (
        cell.formula is None and cell.value is None and cell.kind is CellKind.BLANK
    )


def _same_scalar(left: object, right: object) -> bool:
    return type(left) is type(right) and left == right


def _cell_types_differ(before: CellSnapshot, after: CellSnapshot) -> bool:
    return (
        before.kind is not after.kind
        or before.data_type != after.data_type
        or before.is_date != after.is_date
    )


def _semantic_styles_equal(before: StyleSummary, after: StyleSummary) -> bool:
    """Compare rendered style semantics, not workbook-local style table indexes."""

    return _semantic_style_signature(before) == _semantic_style_signature(after)


def _semantic_style_signature(style: StyleSummary) -> tuple[str, str, str, str, str, str]:
    return (
        style.number_format,
        style.font,
        style.fill,
        style.border,
        style.alignment,
        style.protection,
    )


def _kind_value(cell: CellSnapshot | None) -> str:
    return CellKind.BLANK.value if cell is None else cell.kind.value


def _coordinate_sort_key(coordinate: str) -> tuple[int, int, str]:
    column = 0
    row_text = ""
    for character in coordinate.upper():
        if character.isalpha():
            column = column * 26 + (ord(character) - ord("A") + 1)
        elif character.isdigit():
            row_text += character
    row = int(row_text) if row_text else 0
    return row, column, coordinate


__all__ = ["CellDiffer"]
