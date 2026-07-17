"""Detect copied-formula patterns around an edited cell."""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from tabulint.models import SheetSnapshot
from tabulint.parser.formulas import formula_pattern, normalise_coordinate


@dataclass(frozen=True, slots=True)
class NeighborPattern:
    """Nearby formulas matching the formula expected at a target cell."""

    matches: tuple[str, ...]
    row_break: bool
    column_break: bool

    @property
    def broken(self) -> bool:
        return self.row_break or self.column_break


class FormulaPatternDetector:
    """Compare formulas by relative-reference signature, not literal text."""

    def __init__(self, *, radius: int = 2) -> None:
        if radius < 1:
            raise ValueError("radius must be at least 1")
        self.radius = radius

    def analyze(
        self,
        sheet: SheetSnapshot,
        coordinate: str,
        expected_formula: str,
    ) -> NeighborPattern:
        coordinate = normalise_coordinate(coordinate)
        expected = formula_pattern(
            expected_formula,
            origin=coordinate,
            default_sheet=sheet.name,
        )
        if expected is None:
            return NeighborPattern(matches=(), row_break=False, column_break=False)

        column_text = "".join(character for character in coordinate if character.isalpha())
        row_text = "".join(character for character in coordinate if character.isdigit())
        column = column_index_from_string(column_text)
        row = int(row_text)
        matches: dict[tuple[int, int], str] = {}
        for column_delta, row_delta in self._offsets():
            candidate_column = column + column_delta
            candidate_row = row + row_delta
            if not (1 <= candidate_column <= 16_384 and 1 <= candidate_row <= 1_048_576):
                continue
            candidate_coordinate = f"{get_column_letter(candidate_column)}{candidate_row}"
            candidate = sheet.cells.get(candidate_coordinate)
            if candidate is None or candidate.formula is None:
                continue
            candidate_pattern = formula_pattern(
                candidate.formula,
                origin=candidate_coordinate,
                default_sheet=sheet.name,
            )
            if candidate_pattern == expected:
                matches[(column_delta, row_delta)] = (
                    f"{candidate_coordinate}:{candidate.formula}"
                )

        left = any(delta_row == 0 and delta_column < 0 for delta_column, delta_row in matches)
        right = any(delta_row == 0 and delta_column > 0 for delta_column, delta_row in matches)
        above = any(delta_column == 0 and delta_row < 0 for delta_column, delta_row in matches)
        below = any(delta_column == 0 and delta_row > 0 for delta_column, delta_row in matches)
        ordered = tuple(
            value
            for _, value in sorted(
                matches.items(),
                key=lambda item: (abs(item[0][0]) + abs(item[0][1]), item[0][1], item[0][0]),
            )
        )
        return NeighborPattern(
            matches=ordered,
            row_break=left and right,
            column_break=above and below,
        )

    def _offsets(self) -> list[tuple[int, int]]:
        offsets: list[tuple[int, int]] = []
        for distance in range(1, self.radius + 1):
            offsets.extend(
                [
                    (0, -distance),
                    (0, distance),
                    (-distance, 0),
                    (distance, 0),
                ]
            )
        return offsets


def neighboring_formula_pattern(
    sheet: SheetSnapshot,
    coordinate: str,
    expected_formula: str,
    *,
    radius: int = 2,
) -> NeighborPattern:
    """Functional wrapper for callers that do not need a detector instance."""

    return FormulaPatternDetector(radius=radius).analyze(sheet, coordinate, expected_formula)


__all__ = ["FormulaPatternDetector", "NeighborPattern", "neighboring_formula_pattern"]
