"""Conservative Excel formula tokenisation and A1-reference extraction.

Tabulint does not try to be an Excel calculation engine.  This module only
recognises a deliberately small, deterministic subset that is sufficient for
formula differencing and dependency analysis.  A formula that falls outside
that subset is retained and explicitly marked unsupported.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Final

from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter

_CELL_RE: Final[re.Pattern[str]] = re.compile(
    r"^(?P<column_abs>\$?)(?P<column>[A-Z]{1,3})"
    r"(?P<row_abs>\$?)(?P<row>[1-9][0-9]*)$",
    re.IGNORECASE,
)
_REFERENCE_RE: Final[re.Pattern[str]] = re.compile(
    r"^(?P<start>\$?[A-Z]{1,3}\$?[1-9][0-9]*)"
    r"(?::(?P<end>\$?[A-Z]{1,3}\$?[1-9][0-9]*))?$",
    re.IGNORECASE,
)
_COLUMN_RANGE_RE: Final[re.Pattern[str]] = re.compile(
    r"^(?P<start>\$?[A-Z]{1,3}):(?P<end>\$?[A-Z]{1,3})$",
    re.IGNORECASE,
)
_ROW_RANGE_RE: Final[re.Pattern[str]] = re.compile(
    r"^(?P<start>\$?[1-9][0-9]*):(?P<end>\$?[1-9][0-9]*)$"
)
_EXTERNAL_SHEET_RE: Final[re.Pattern[str]] = re.compile(
    r"^\[(?P<book>[^\]]+)\](?P<sheet>.*)$"
)

# These functions receive reference-level analysis.  Their values are never
# evaluated.  Dynamic indirection functions (INDIRECT/OFFSET) are intentionally
# absent because their real dependencies cannot be inferred from text alone.
SUPPORTED_FUNCTIONS: Final[frozenset[str]] = frozenset(
    {
        "ABS",
        "AND",
        "AVERAGE",
        "CHOOSE",
        "CONCAT",
        "CONCATENATE",
        "COUNT",
        "COUNTA",
        "COUNTIF",
        "COUNTIFS",
        "DATE",
        "DAY",
        "HLOOKUP",
        "IF",
        "IFERROR",
        "IFS",
        "INDEX",
        "LEFT",
        "LEN",
        "LOWER",
        "MATCH",
        "MAX",
        "MID",
        "MIN",
        "MONTH",
        "NOT",
        "OR",
        "RIGHT",
        "ROUND",
        "ROUNDDOWN",
        "ROUNDUP",
        "SUBTOTAL",
        "SUM",
        "SUMIF",
        "SUMIFS",
        "TEXT",
        "TODAY",
        "UPPER",
        "VLOOKUP",
        "XLOOKUP",
        "YEAR",
    }
)


def normalise_coordinate(coordinate: str) -> str:
    """Return an uppercase A1 coordinate without absolute-reference markers."""

    match = _CELL_RE.fullmatch(coordinate.strip())
    if match is None:
        raise ValueError(f"Invalid A1 coordinate: {coordinate!r}")
    column = match.group("column").upper()
    row = int(match.group("row"))
    if column_index_from_string(column) > 16_384 or row > 1_048_576:
        raise ValueError(f"Coordinate is outside Excel worksheet bounds: {coordinate!r}")
    return f"{column}{row}"


def qualify_cell(sheet: str, coordinate: str) -> str:
    """Create Tabulint's stable ``Sheet!A1`` location representation."""

    return f"{sheet}!{normalise_coordinate(coordinate)}"


def split_location(location: str, *, default_sheet: str | None = None) -> tuple[str, str]:
    """Split a Tabulint or Excel-like location into sheet and coordinate."""

    value = location.strip()
    if "!" in value:
        sheet, coordinate = value.rsplit("!", 1)
        sheet = _unquote_sheet(sheet)
    elif default_sheet is not None:
        sheet, coordinate = default_sheet, value
    else:
        raise ValueError(f"Location must include a worksheet: {location!r}")
    return sheet, normalise_coordinate(coordinate)


@dataclass(frozen=True, slots=True)
class FormulaReference:
    """A parsed single-cell or rectangular A1 reference."""

    sheet: str | None
    start: str
    end: str | None = None
    raw: str = ""
    external_workbook: str | None = None
    range_kind: str = "cells"

    @property
    def is_range(self) -> bool:
        return self.end is not None

    def display(self, default_sheet: str | None = None) -> str:
        """Return a compact reference, omitting the current sheet when possible."""

        if self.range_kind == "columns":
            address = f"{_column_text(self.start)}:{_column_text(self.end or self.start)}"
        elif self.range_kind == "rows":
            address = f"{_row_text(self.start)}:{_row_text(self.end or self.start)}"
        else:
            address = self.start.upper()
            if self.end is not None:
                address = f"{address}:{self.end.upper()}"
        if self.external_workbook is None and (
            self.sheet is None or self.sheet == default_sheet
        ):
            return address
        prefix = self.sheet
        if self.external_workbook is not None:
            prefix = f"[{self.external_workbook}]{prefix}"
        return f"{prefix}!{address}"

    def expand(
        self,
        default_sheet: str,
        *,
        limit: int = 10_000,
        qualified: bool = True,
    ) -> tuple[tuple[str, ...], bool]:
        """Expand a rectangular range with an explicit cell-count limit.

        The boolean result is true when expansion was truncated.  External
        workbook references deliberately return no local graph nodes.
        """

        if limit < 1:
            return (), True
        if self.external_workbook is not None:
            return (), False
        sheet = self.sheet or default_sheet
        start_column, start_row = _coordinate_parts(self.start)
        if self.end is None:
            end_column, end_row = start_column, start_row
        else:
            end_column, end_row = _coordinate_parts(self.end)
        min_column, max_column = sorted((start_column, end_column))
        min_row, max_row = sorted((start_row, end_row))
        total = (max_column - min_column + 1) * (max_row - min_row + 1)
        result: list[str] = []
        for row in range(min_row, max_row + 1):
            for column in range(min_column, max_column + 1):
                coordinate = f"{get_column_letter(column)}{row}"
                result.append(qualify_cell(sheet, coordinate) if qualified else coordinate)
                if len(result) >= limit:
                    return tuple(result), total > limit
        return tuple(result), False


@dataclass(frozen=True, slots=True)
class FormulaAnalysis:
    """Reference-level analysis of one formula."""

    formula: str
    supported: bool
    functions: tuple[str, ...]
    references: tuple[FormulaReference, ...]
    unsupported_reasons: tuple[str, ...] = ()
    structure: str = ""

    @property
    def analysis_status(self) -> str:
        return "supported" if self.supported else "unsupported_formula_analysis"

    def reference_strings(self, default_sheet: str | None = None) -> list[str]:
        return [reference.display(default_sheet) for reference in self.references]

    def expanded_references(
        self,
        default_sheet: str,
        *,
        limit: int = 20_000,
        qualified: bool = False,
    ) -> tuple[list[str], bool]:
        """Return unique referenced cells in deterministic order."""

        remaining = limit
        values: set[str] = set()
        truncated = False
        for reference in self.references:
            if remaining < 1:
                truncated = True
                break
            expanded, was_truncated = reference.expand(
                default_sheet,
                limit=remaining,
                qualified=qualified,
            )
            values.update(expanded)
            remaining = max(0, limit - len(values))
            truncated = truncated or was_truncated
        return sorted(values, key=_location_sort_key), truncated


class FormulaParser:
    """Parse formulas without evaluating them or resolving external content."""

    def __init__(self, *, supported_functions: frozenset[str] | None = None) -> None:
        self.supported_functions = supported_functions or SUPPORTED_FUNCTIONS

    def parse(self, formula: str) -> FormulaAnalysis:
        reasons: set[str] = set()
        if not isinstance(formula, str) or not formula.startswith("="):
            return FormulaAnalysis(
                formula=str(formula),
                supported=False,
                functions=(),
                references=(),
                unsupported_reasons=("not_a_formula",),
            )
        try:
            tokens = tuple(Tokenizer(formula).items)
        except Exception as exc:  # openpyxl raises several tokenizer exception types
            return FormulaAnalysis(
                formula=formula,
                supported=False,
                functions=(),
                references=(),
                unsupported_reasons=(f"tokenization_failed:{type(exc).__name__}",),
            )

        functions: list[str] = []
        references: list[FormulaReference] = []
        structure_parts: list[str] = []
        function_depth = 0
        for token in tokens:
            token_type = str(token.type)
            subtype = str(token.subtype)
            value = str(token.value)
            if token_type == "FUNC" and subtype == "OPEN":
                function_depth += 1
                name = _normalise_function_name(value[:-1])
                functions.append(name)
                structure_parts.append(f"{name}(")
                if name not in self.supported_functions:
                    reasons.add(f"unsupported_function:{name}")
                continue
            if token_type == "FUNC" and subtype == "CLOSE":
                function_depth -= 1
                structure_parts.append(")")
                continue
            if token_type == "OPERAND" and subtype == "RANGE":
                reference = parse_reference(value)
                if reference is None:
                    reasons.add(f"unsupported_reference:{value}")
                    structure_parts.append(value.upper())
                else:
                    references.append(reference)
                    structure_parts.append("<REF>")
                    if reference.external_workbook is not None:
                        reasons.add("external_workbook_reference")
                continue
            if token_type == "OPERAND" and subtype == "ERROR" and value.upper() == "#REF!":
                reasons.add("unresolved_reference_error")
            if token_type == "WSPACE":
                continue
            structure_parts.append(value.upper())

        if function_depth != 0:
            reasons.add("unbalanced_function_parentheses")
        return FormulaAnalysis(
            formula=formula,
            supported=not reasons,
            functions=tuple(functions),
            references=tuple(references),
            unsupported_reasons=tuple(sorted(reasons)),
            structure="".join(structure_parts),
        )

    def extract_references(
        self, formula: str, *, default_sheet: str | None = None
    ) -> list[str]:
        """Convenience API returning compact references from ``formula``."""

        return self.parse(formula).reference_strings(default_sheet)


def analyze_formula(formula: str) -> FormulaAnalysis:
    """Analyze one formula with the default conservative function set."""

    return FormulaParser().parse(formula)


def extract_references(
    formula: str,
    *,
    default_sheet: str | None = None,
) -> list[str]:
    """Return compact A1 references found in a formula."""

    return FormulaParser().extract_references(formula, default_sheet=default_sheet)


def parse_reference(value: str) -> FormulaReference | None:
    """Parse one tokenizer RANGE operand as an A1 reference."""

    raw = value.strip()
    sheet: str | None = None
    external_workbook: str | None = None
    address = raw
    if "!" in raw:
        sheet_part, address = raw.rsplit("!", 1)
        sheet_part = _unquote_sheet(sheet_part)
        external_match = _EXTERNAL_SHEET_RE.fullmatch(sheet_part)
        if external_match is not None:
            external_workbook = external_match.group("book")
            sheet = external_match.group("sheet")
        else:
            if "[" in sheet_part or "]" in sheet_part:
                return None
            sheet = sheet_part
        if not sheet or ":" in sheet:
            return None
    match = _REFERENCE_RE.fullmatch(address)
    range_kind = "cells"
    if match is not None:
        start = match.group("start").upper()
        end = match.group("end")
        end = end.upper() if end is not None else None
    else:
        column_match = _COLUMN_RANGE_RE.fullmatch(address)
        row_match = _ROW_RANGE_RE.fullmatch(address)
        if column_match is not None:
            range_kind = "columns"
            start_column = column_match.group("start").upper()
            end_column = column_match.group("end").upper()
            start = f"{start_column}$1"
            end = f"{end_column}$1048576"
        elif row_match is not None:
            range_kind = "rows"
            start_row = row_match.group("start")
            end_row = row_match.group("end")
            start = f"$A{start_row}"
            end = f"$XFD{end_row}"
        else:
            return None
    try:
        normalise_coordinate(start)
        if end is not None:
            normalise_coordinate(end)
    except ValueError:
        return None
    return FormulaReference(
        sheet=sheet,
        start=start,
        end=end,
        raw=raw,
        external_workbook=external_workbook,
        range_kind=range_kind,
    )


def formula_pattern(formula: str, *, origin: str, default_sheet: str) -> str | None:
    """Return a copy-pattern signature using relative row/column offsets.

    Formulas copied down a column or across a row have the same signature even
    though their textual A1 references differ.  ``None`` means a range operand
    could not be represented safely.
    """

    if not formula.startswith("="):
        return None
    try:
        tokens = tuple(Tokenizer(formula).items)
        origin_column, origin_row = _coordinate_parts(origin)
    except Exception:
        return None

    parts: list[str] = []
    for token in tokens:
        token_type = str(token.type)
        subtype = str(token.subtype)
        value = str(token.value)
        if token_type == "WSPACE":
            continue
        if token_type == "OPERAND" and subtype == "RANGE":
            reference = parse_reference(value)
            if reference is None or reference.external_workbook is not None:
                return None
            sheet = reference.sheet or default_sheet
            start = _relative_marker(reference.start, origin_column, origin_row)
            end = (
                _relative_marker(reference.end, origin_column, origin_row)
                if reference.end is not None
                else None
            )
            marker = start if end is None else f"{start}:{end}"
            parts.append(f"[{sheet}!{marker}]")
        elif token_type == "FUNC" and subtype == "OPEN":
            parts.append(f"{_normalise_function_name(value[:-1])}(")
        else:
            parts.append(value.upper())
    # Unsupported functions still yield a useful syntactic pattern, but an
    # un-tokenisable/unsupported range does not.
    return "".join(parts)


def _coordinate_parts(coordinate: str) -> tuple[int, int]:
    match = _CELL_RE.fullmatch(coordinate.strip())
    if match is None:
        raise ValueError(f"Invalid A1 coordinate: {coordinate!r}")
    column = column_index_from_string(match.group("column").upper())
    row = int(match.group("row"))
    if column > 16_384 or row > 1_048_576:
        raise ValueError(f"Coordinate is outside Excel worksheet bounds: {coordinate!r}")
    return column, row


def _relative_marker(
    coordinate: str | None, origin_column: int, origin_row: int
) -> str:
    if coordinate is None:
        raise ValueError("A coordinate is required")
    match = _CELL_RE.fullmatch(coordinate)
    if match is None:
        raise ValueError(f"Invalid A1 coordinate: {coordinate!r}")
    column = column_index_from_string(match.group("column").upper())
    row = int(match.group("row"))
    column_marker = (
        f"C${column}" if match.group("column_abs") else f"C{column - origin_column:+d}"
    )
    row_marker = f"R${row}" if match.group("row_abs") else f"R{row - origin_row:+d}"
    return f"{column_marker}{row_marker}"


def _normalise_function_name(name: str) -> str:
    value = name.strip().upper()
    for prefix in ("_XLFN.", "_XLWS."):
        if value.startswith(prefix):
            value = value[len(prefix) :]
    return value


def _unquote_sheet(sheet: str) -> str:
    value = sheet.strip()
    if len(value) >= 2 and value.startswith("'") and value.endswith("'"):
        value = value[1:-1].replace("''", "'")
    return value


def _column_text(coordinate: str) -> str:
    match = _CELL_RE.fullmatch(coordinate)
    if match is None:
        raise ValueError(f"Invalid A1 coordinate: {coordinate!r}")
    absolute = "$" if match.group("column_abs") else ""
    return f"{absolute}{match.group('column').upper()}"


def _row_text(coordinate: str) -> str:
    match = _CELL_RE.fullmatch(coordinate)
    if match is None:
        raise ValueError(f"Invalid A1 coordinate: {coordinate!r}")
    absolute = "$" if match.group("row_abs") else ""
    return f"{absolute}{match.group('row')}"


def _location_sort_key(location: str) -> tuple[str, int, int, str]:
    if "!" in location:
        sheet, coordinate = location.rsplit("!", 1)
    else:
        sheet, coordinate = "", location
    try:
        column, row = _coordinate_parts(coordinate)
    except ValueError:
        return sheet.casefold(), 2**31 - 1, 2**31 - 1, coordinate
    return sheet.casefold(), row, column, coordinate


__all__ = [
    "SUPPORTED_FUNCTIONS",
    "FormulaAnalysis",
    "FormulaParser",
    "FormulaReference",
    "analyze_formula",
    "extract_references",
    "formula_pattern",
    "normalise_coordinate",
    "parse_reference",
    "qualify_cell",
    "split_location",
]
