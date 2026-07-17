"""Formula-specific semantic differences and overwrite detection."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any

from tabulint.diff.pattern_diff import FormulaPatternDetector
from tabulint.models import (
    CellKind,
    FormulaChange,
    RiskLevel,
    ScalarValue,
    SheetSnapshot,
)
from tabulint.parser.formulas import FormulaAnalysis, FormulaParser, qualify_cell


class FormulaDiffer:
    """Compare formulas and identify high-impact formula overwrites."""

    _MANUAL_REVIEW_RECOMMENDATION = (
        "Manually verify that the formula replacement was intentional and that "
        "dependent results remain correct."
    )

    def __init__(
        self,
        *,
        expansion_limit: int = 20_000,
        total_expansion_limit: int = 200_000,
        pattern_radius: int = 2,
    ) -> None:
        if (
            isinstance(expansion_limit, bool)
            or not isinstance(expansion_limit, int)
            or expansion_limit < 1
        ):
            raise ValueError("expansion_limit must be a positive integer")
        if (
            isinstance(total_expansion_limit, bool)
            or not isinstance(total_expansion_limit, int)
            or total_expansion_limit < 2
        ):
            raise ValueError("total_expansion_limit must be an integer of at least 2")
        self.expansion_limit = expansion_limit
        self.total_expansion_limit = total_expansion_limit
        self.parser = FormulaParser()
        self.pattern_detector = FormulaPatternDetector(radius=pattern_radius)

    def new_budget(self) -> _ExpansionBudget:
        """Create one shared expansion budget for a complete workbook review."""

        return _ExpansionBudget(self.total_expansion_limit)

    def compare(
        self,
        before: SheetSnapshot,
        after: SheetSnapshot,
        *,
        budget: _ExpansionBudget | None = None,
    ) -> list[FormulaChange]:
        if before.name != after.name:
            raise ValueError("FormulaDiffer requires snapshots of the same worksheet")
        active_budget = budget or self.new_budget()
        changes: list[FormulaChange] = []
        coordinates = sorted(set(before.cells) | set(after.cells), key=_coordinate_sort_key)
        for coordinate in coordinates:
            before_cell = before.cells.get(coordinate)
            after_cell = after.cells.get(coordinate)
            before_formula = before_cell.formula if before_cell is not None else None
            after_formula = after_cell.formula if after_cell is not None else None
            before_attributes = (
                before_cell.formula_attributes if before_cell is not None else {}
            )
            after_attributes = after_cell.formula_attributes if after_cell is not None else {}
            attributes_changed = before_attributes != after_attributes
            if before_formula == after_formula and not attributes_changed:
                continue
            if before_formula is None and after_formula is None:
                continue
            changes.append(
                self._compare_formula(
                    before,
                    after,
                    coordinate,
                    before_formula,
                    after_formula,
                    before_attributes,
                    after_attributes,
                    active_budget,
                )
            )
        return changes

    def _compare_formula(
        self,
        before_sheet: SheetSnapshot,
        after_sheet: SheetSnapshot,
        coordinate: str,
        before_formula: str | None,
        after_formula: str | None,
        before_attributes: Mapping[str, object],
        after_attributes: Mapping[str, object],
        budget: _ExpansionBudget,
    ) -> FormulaChange:
        before_analysis = self.parser.parse(before_formula) if before_formula is not None else None
        after_analysis = self.parser.parse(after_formula) if after_formula is not None else None
        location = qualify_cell(before_sheet.name, coordinate)
        references_before = _reference_strings(before_analysis, before_sheet.name)
        references_after = _reference_strings(after_analysis, after_sheet.name)
        excluded: list[str] = []
        expansion_truncated = False
        neighboring: list[str] = []
        pattern_broken = False
        pattern_row_break = False
        pattern_column_break = False

        if before_formula is not None:
            pattern = self.pattern_detector.analyze(after_sheet, coordinate, before_formula)
            neighboring = list(pattern.matches)
            pattern_broken = pattern.broken
            pattern_row_break = pattern.row_break
            pattern_column_break = pattern.column_break

        change_type: str
        risk_level: RiskLevel
        description: str
        high_impact: bool
        replacement_value: ScalarValue = None
        replacement_kind: CellKind | None = None
        manual_review_recommendation: str | None = None
        if before_formula is not None and after_formula is None:
            after_cell = after_sheet.cells.get(coordinate)
            replacement_value = after_cell.value if after_cell is not None else None
            replacement_kind = after_cell.kind if after_cell is not None else CellKind.BLANK
            manual_review_recommendation = self._MANUAL_REVIEW_RECOMMENDATION
            if after_cell is None or after_cell.value is None:
                change_type = "formula_deleted"
                description = "formula was deleted"
            else:
                change_type = "formula_overwritten"
                description = "formula was replaced by a fixed value"
            risk_level = RiskLevel.HIGH
            high_impact = True
        elif before_formula is None and after_formula is not None:
            change_type = "formula_added"
            risk_level = RiskLevel.MEDIUM
            description = "fixed or blank cell was replaced by a formula"
            high_impact = False
        else:
            assert before_analysis is not None
            assert after_analysis is not None
            same_formula_shape = (
                before_analysis.structure == after_analysis.structure
                and before_analysis.functions == after_analysis.functions
                and len(before_analysis.references) == len(after_analysis.references)
            )
            excluded_qualified: list[str]
            per_formula_limit = min(self.expansion_limit, budget.remaining // 2)
            if per_formula_limit < 1:
                excluded_qualified = []
                references_reduced = False
                expansion_truncated = True
            else:
                (
                    excluded_qualified,
                    references_reduced,
                    expansion_truncated,
                    expansion_cost,
                ) = _reference_reduction(
                    before_analysis,
                    after_analysis,
                    before_sheet.name,
                    limit=per_formula_limit,
                )
                budget.consume(expansion_cost)
            is_reduction = (
                before_analysis.supported
                and after_analysis.supported
                and not expansion_truncated
                and same_formula_shape
                and references_reduced
            )
            excluded = [
                _compact_location(value, before_sheet.name) for value in excluded_qualified
            ]
            if before_formula == after_formula and before_attributes != after_attributes:
                change_type = "formula_changed"
                risk_level = RiskLevel.HIGH
                description = "array or data-table formula metadata changed"
                high_impact = True
            elif is_reduction:
                change_type = "formula_range_reduced"
                risk_level = RiskLevel.HIGH
                description = (
                    "formula reference range was reduced; excluded "
                    + ", ".join(excluded[:8])
                    + (" and more" if len(excluded) > 8 else "")
                )
                high_impact = True
            else:
                change_type = "formula_changed"
                risk_level = RiskLevel.HIGH if pattern_broken else RiskLevel.MEDIUM
                description = "formula expression changed"
                high_impact = pattern_broken

        analyses = [item for item in (before_analysis, after_analysis) if item is not None]
        supported_analysis = all(item.supported for item in analyses) and not expansion_truncated
        unsupported_reasons = sorted(
            {reason for item in analyses for reason in item.unsupported_reasons}
        )
        if expansion_truncated:
            unsupported_reasons.append("reference_expansion_truncated")
        evidence: dict[str, Any] = {
            "analysis_status": (
                "supported" if supported_analysis else "unsupported_formula_analysis"
            ),
            "unsupported_reasons": unsupported_reasons,
            "pattern_broken": pattern_broken,
            "pattern_axis": (
                "row_and_column"
                if pattern_row_break and pattern_column_break
                else "row"
                if pattern_row_break
                else "column"
                if pattern_column_break
                else None
            ),
            "reference_expansion_truncated": expansion_truncated,
            "before_formula_attributes": dict(before_attributes),
            "after_formula_attributes": dict(after_attributes),
        }
        if not supported_analysis:
            description = f"{description}; unsupported_formula_analysis"
        if pattern_broken:
            description = f"{description}; neighboring copied-formula pattern was broken"

        return FormulaChange(
            change_type=change_type,
            risk_level=risk_level,
            sheet=before_sheet.name,
            coordinate=coordinate,
            location=location,
            before_formula=before_formula,
            after_formula=after_formula,
            replacement_value=replacement_value,
            replacement_kind=replacement_kind,
            manual_review_recommendation=manual_review_recommendation,
            description=f"{location}: {description}",
            supported_analysis=supported_analysis,
            high_impact=high_impact,
            references_before=references_before,
            references_after=references_after,
            excluded_references=excluded,
            neighboring_formula_pattern=neighboring,
            evidence=evidence,
        )


def _reference_strings(analysis: FormulaAnalysis | None, sheet: str) -> list[str]:
    if analysis is None:
        return []
    return analysis.reference_strings(sheet)


def _reference_reduction(
    before: FormulaAnalysis,
    after: FormulaAnalysis,
    sheet: str,
    *,
    limit: int,
) -> tuple[list[str], bool, bool, int]:
    """Compare corresponding reference rectangles without hiding overlaps."""

    if len(before.references) != len(after.references):
        return [], False, False, 0
    remaining = limit
    expansion_cost = 0
    excluded: set[str] = set()
    reduced = False
    for before_reference, after_reference in zip(
        before.references,
        after.references,
        strict=True,
    ):
        if remaining < 1:
            return sorted(excluded, key=_location_sort_key), False, True, expansion_cost
        before_cells, before_truncated = before_reference.expand(
            sheet,
            limit=remaining,
            qualified=True,
        )
        after_cells, after_truncated = after_reference.expand(
            sheet,
            limit=remaining,
            qualified=True,
        )
        expansion_cost += len(before_cells) + len(after_cells)
        if before_truncated or after_truncated:
            return sorted(excluded, key=_location_sort_key), False, True, expansion_cost
        before_set = set(before_cells)
        after_set = set(after_cells)
        if not after_set <= before_set:
            return sorted(excluded, key=_location_sort_key), False, False, expansion_cost
        difference = before_set - after_set
        if difference:
            reduced = True
            excluded.update(difference)
        remaining -= max(len(before_set), len(after_set))
    return sorted(excluded, key=_location_sort_key), reduced, False, expansion_cost


@dataclass(slots=True)
class _ExpansionBudget:
    remaining: int

    def consume(self, amount: int) -> None:
        if amount < 0 or amount > self.remaining:
            raise ValueError("formula expansion budget accounting error")
        self.remaining -= amount


def _compact_location(location: str, current_sheet: str) -> str:
    sheet, coordinate = location.rsplit("!", 1)
    return coordinate if sheet == current_sheet else location


def _coordinate_sort_key(coordinate: str) -> tuple[int, int, str]:
    column = 0
    row_text = ""
    for character in coordinate.upper():
        if character.isalpha():
            column = column * 26 + (ord(character) - ord("A") + 1)
        elif character.isdigit():
            row_text += character
    return int(row_text or "0"), column, coordinate


def _location_sort_key(location: str) -> tuple[str, int, int, str]:
    sheet, coordinate = location.rsplit("!", 1)
    row, column, _ = _coordinate_sort_key(coordinate)
    return sheet.casefold(), row, column, coordinate


__all__ = ["FormulaDiffer"]
