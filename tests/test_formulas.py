"""Formula parsing and copied-pattern tests."""

from __future__ import annotations

import pytest

from tabulint.parser.formulas import (
    FormulaParser,
    formula_pattern,
    normalise_coordinate,
    parse_reference,
    split_location,
)


def test_formula_parser_extracts_local_cross_sheet_and_mixed_references() -> None:
    analysis = FormulaParser().parse(
        "=SUM($A1:B$3)+'Other Sheet'!$C$4+IF(D5>0,Data!E6,0)"
    )

    assert analysis.supported is True
    assert analysis.functions == ("SUM", "IF")
    assert analysis.reference_strings("Calc") == [
        "$A1:B$3",
        "Other Sheet!$C$4",
        "D5",
        "Data!E6",
    ]
    expanded, truncated = analysis.expanded_references("Calc", qualified=True)
    assert truncated is False
    assert "Calc!A1" in expanded
    assert "Calc!B3" in expanded
    assert "Other Sheet!C4" in expanded
    assert "Data!E6" in expanded


def test_formula_parser_marks_external_and_unknown_functions_unsupported() -> None:
    analysis = FormulaParser().parse("=MYSTERY('[Budget.xlsx]Plan'!A1)")

    assert analysis.supported is False
    assert analysis.reference_strings() == ["[Budget.xlsx]Plan!A1"]
    assert "external_workbook_reference" in analysis.unsupported_reasons
    assert "unsupported_function:MYSTERY" in analysis.unsupported_reasons


@pytest.mark.parametrize(
    "formula",
    [
        "=AVERAGE(A1:A3)",
        "=MIN(A1:A3)",
        "=MAX(A1:A3)",
        "=COUNT(A1:A3)",
        "=COUNTA(A1:A3)",
        '=SUMIF(A1:A3,">0",B1:B3)',
        '=SUMIFS(B1:B3,A1:A3,">0")',
        '=COUNTIF(A1:A3,">0")',
        '=COUNTIFS(A1:A3,">0",B1:B3,"<10")',
        "=VLOOKUP(A1,Data!A:B,2,FALSE)",
        "=XLOOKUP(A1,Data!A:A,Data!B:B)",
        "=INDEX(Data!B:B,MATCH(A1,Data!A:A,0))",
    ],
)
def test_documented_common_functions_have_supported_reference_analysis(
    formula: str,
) -> None:
    analysis = FormulaParser().parse(formula)

    assert analysis.supported is True
    assert analysis.references
    assert analysis.unsupported_reasons == ()


def test_absolute_relative_and_mixed_references_keep_copy_pattern() -> None:
    original = formula_pattern(
        "=$A2+B$1+$C$3",
        origin="D2",
        default_sheet="Calc",
    )
    copied = formula_pattern(
        "=$A3+C$1+$C$3",
        origin="E3",
        default_sheet="Calc",
    )

    assert copied == original
    assert original is not None
    assert "C$1" in original
    assert "R$3" in original
    assert "C-2" in original


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("$A$1", (None, "$A$1", None)),
        ("'Sales Plan'!B2:C4", ("Sales Plan", "B2", "C4")),
        ("A:A", (None, "A$1", "A$1048576")),
        ("2:3", (None, "$A2", "$XFD3")),
    ],
)
def test_parse_reference_modes(
    value: str,
    expected: tuple[str | None, str, str | None],
) -> None:
    reference = parse_reference(value)

    assert reference is not None
    assert (reference.sheet, reference.start, reference.end) == expected


def test_coordinate_and_location_validation() -> None:
    assert normalise_coordinate("$b$12") == "B12"
    assert split_location("'Sales Plan'!$b$12") == ("Sales Plan", "B12")
    assert split_location("C3", default_sheet="Inputs") == ("Inputs", "C3")

    with pytest.raises(ValueError, match="outside Excel worksheet bounds"):
        normalise_coordinate("XFE1")
    with pytest.raises(ValueError, match="must include a worksheet"):
        split_location("A1")
