"""Explainable risk-score calculation, bounds, and de-duplication tests."""

from __future__ import annotations

import pytest

from tabulint.models import CellChange, FormulaChange, RiskLevel, StructureChange
from tabulint.risk.scorer import MANY_CELLS_THRESHOLD, RiskScorer, risk_level_for_score


def _structure(change_type: str, location: str) -> StructureChange:
    return StructureChange(
        change_type=change_type,
        risk_level=RiskLevel.HIGH,
        location=location,
        before=None,
        after=True,
        description=f"{change_type} at {location}",
        evidence={"source": "test"},
    )


def _cell(change_type: str, coordinate: str, *, sheet: str = "Sheet1") -> CellChange:
    return CellChange(
        change_type=change_type,
        risk_level=RiskLevel.MEDIUM,
        sheet=sheet,
        coordinate=coordinate,
        location=f"{sheet}!{coordinate}",
        before=None,
        after=None,
        description=f"{change_type} at {sheet}!{coordinate}",
        evidence={"source": "test"},
    )


def _formula(change_type: str, coordinate: str = "A1") -> FormulaChange:
    return FormulaChange(
        change_type=change_type,
        risk_level=RiskLevel.HIGH,
        sheet="Sheet1",
        coordinate=coordinate,
        location=f"Sheet1!{coordinate}",
        before_formula="=B1",
        after_formula=None,
        description=f"{change_type} at Sheet1!{coordinate}",
        supported_analysis=True,
        high_impact=True,
        evidence={"source": "test"},
    )


def test_risk_score_sums_explainable_factors() -> None:
    score, level, factors = RiskScorer().score(
        [
            _structure("macro_added", "workbook:vba"),
            _structure("external_link_added", "workbook:external_links"),
        ],
        [_cell("text_changed", "A1")],
        [],
    )

    assert score == 76
    assert level is RiskLevel.HIGH
    assert [(factor.risk_type, factor.points) for factor in factors] == [
        ("macro_added", 40),
        ("external_link_added", 35),
        ("text_changed", 1),
    ]
    assert all(factor.description and factor.evidence for factor in factors)


def test_risk_score_is_capped_at_one_hundred() -> None:
    score, level, factors = RiskScorer().score(
        [
            _structure("macro_added", "workbook:vba"),
            _structure("external_link_added", "workbook:external_links"),
        ],
        [],
        [_formula("formula_overwritten")],
    )

    assert sum(factor.points for factor in factors) == 105
    assert score == 100
    assert level is RiskLevel.CRITICAL


@pytest.mark.parametrize(
    ("score", "expected"),
    [
        (0, RiskLevel.LOW),
        (19, RiskLevel.LOW),
        (20, RiskLevel.MEDIUM),
        (49, RiskLevel.MEDIUM),
        (50, RiskLevel.HIGH),
        (79, RiskLevel.HIGH),
        (80, RiskLevel.CRITICAL),
        (100, RiskLevel.CRITICAL),
    ],
)
def test_risk_level_boundaries(score: int, expected: RiskLevel) -> None:
    assert risk_level_for_score(score) is expected


@pytest.mark.parametrize("invalid", [-1, 101, True, 1.5])
def test_risk_level_rejects_out_of_range_or_non_integer_values(invalid: object) -> None:
    with pytest.raises(ValueError, match="integer from 0 through 100"):
        risk_level_for_score(invalid)  # type: ignore[arg-type]


def test_duplicate_facts_and_mirrored_formula_cell_change_count_once() -> None:
    duplicated = _cell("text_changed", "A1")
    mirrored = _cell("formula_overwritten", "B2")

    score, _, factors = RiskScorer().score(
        [],
        [duplicated, duplicated.model_copy(deep=True), mirrored],
        [_formula("formula_overwritten", "B2")],
    )

    assert score == 31
    assert [(factor.risk_type, factor.location) for factor in factors] == [
        ("formula_overwritten", "Sheet1!B2"),
        ("text_changed", "Sheet1!A1"),
    ]


def test_custom_weights_accept_aliases_and_override_defaults() -> None:
    scorer = RiskScorer({"formula_overwrite": 7, "text_changed": 4})

    score, level, factors = scorer.score(
        [],
        [_cell("text_changed", "A1")],
        [_formula("formula_to_value", "B2")],
    )

    assert score == 11
    assert level is RiskLevel.LOW
    assert {factor.risk_type: factor.points for factor in factors} == {
        "formula_overwritten": 7,
        "text_changed": 4,
    }


def test_unknown_high_risk_change_gets_nonzero_fallback_score() -> None:
    score, level, factors = RiskScorer().score(
        [_structure("future_high_risk_fact", "workbook:future")],
        [],
        [],
    )

    assert score == 15
    assert level is RiskLevel.LOW
    assert len(factors) == 1
    assert factors[0].risk_type == "future_high_risk_fact"
    assert factors[0].points == 15
    assert factors[0].evidence["source_risk_level"] == "HIGH"


def test_many_unique_changed_cells_adds_one_bulk_factor() -> None:
    changes = [_cell("value_changed", f"A{row}") for row in range(1, MANY_CELLS_THRESHOLD + 1)]

    score, level, factors = RiskScorer().score([], changes, [])

    bulk = [factor for factor in factors if factor.risk_type == "many_cells_changed"]
    assert score == 60
    assert level is RiskLevel.HIGH
    assert len(bulk) == 1
    assert bulk[0].evidence["changed_cell_count"] == MANY_CELLS_THRESHOLD


@pytest.mark.parametrize(
    "weights",
    [
        {"": 1},
        {"macro_added": -1},
        {"macro_added": 101},
        {"macro_added": True},
    ],
)
def test_invalid_custom_weights_are_rejected(weights: dict[str, int]) -> None:
    with pytest.raises(ValueError, match="risk weight"):
        RiskScorer(weights)
