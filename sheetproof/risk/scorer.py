"""Deterministic, explainable risk scoring for workbook changes."""

from __future__ import annotations

from collections.abc import Iterable

from sheetproof.models import (
    CANONICAL_RISK_WEIGHT_KEYS,
    DEFAULT_RISK_WEIGHT_VALUES,
    RISK_WEIGHT_ALIASES,
    CellChange,
    FormulaChange,
    RiskFactor,
    RiskLevel,
    StructureChange,
    canonical_risk_weight_name,
)

DEFAULT_RISK_WEIGHTS: dict[str, int] = DEFAULT_RISK_WEIGHT_VALUES.copy()

MANY_CELLS_THRESHOLD = 25

_FALLBACK_RISK_WEIGHTS: dict[RiskLevel, int] = {
    RiskLevel.LOW: 1,
    RiskLevel.MEDIUM: 5,
    RiskLevel.HIGH: 15,
    RiskLevel.CRITICAL: 30,
}

RiskChange = StructureChange | CellChange | FormulaChange


def risk_level_for_score(score: int) -> RiskLevel:
    """Classify an already bounded risk score using public MVP thresholds."""

    if isinstance(score, bool) or not isinstance(score, int) or not 0 <= score <= 100:
        raise ValueError("risk score must be an integer from 0 through 100")
    if score < 20:
        return RiskLevel.LOW
    if score < 50:
        return RiskLevel.MEDIUM
    if score < 80:
        return RiskLevel.HIGH
    return RiskLevel.CRITICAL


def risk_level_from_score(score: int) -> RiskLevel:
    """Readable alias for :func:`risk_level_for_score`."""

    return risk_level_for_score(score)


class RiskScorer:
    """Score recognized risks, keeping a traceable factor for each contribution."""

    def __init__(self, weights: dict[str, int] | None = None) -> None:
        self.weights = DEFAULT_RISK_WEIGHTS.copy()
        if weights is not None:
            for raw_name, value in weights.items():
                if not isinstance(raw_name, str) or not raw_name.strip():
                    raise ValueError("risk weight names must be non-empty strings")
                if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= 100:
                    raise ValueError(
                        f"risk weight '{raw_name}' must be an integer from 0 through 100"
                    )
                name = canonical_risk_weight_name(raw_name)
                if not name:
                    raise ValueError("risk weight names must contain a letter or digit")
                self.weights[name] = value

    @staticmethod
    def level_for_score(score: int) -> RiskLevel:
        """Return the public risk band for ``score``."""

        return risk_level_for_score(score)

    def score(
        self,
        structure_changes: list[StructureChange],
        cell_changes: list[CellChange],
        formula_changes: list[FormulaChange],
    ) -> tuple[int, RiskLevel, list[RiskFactor]]:
        """Return capped score, level, and de-duplicated factor evidence."""

        factors_by_key: dict[tuple[str, str], RiskFactor] = {}
        formula_locations = {change.location.casefold() for change in formula_changes}
        # Formula facts precede their mirrored cell facts so duplicate findings
        # retain the more specific formula explanation and evidence.
        sources: Iterable[RiskChange] = [*structure_changes, *formula_changes, *cell_changes]
        for change in sources:
            risk_type = self._risk_type(change)
            if risk_type is None:
                continue
            if (
                isinstance(change, CellChange)
                and change.location.casefold() in formula_locations
                and risk_type
                in {
                    "formula_overwritten",
                    "formula_deleted",
                    "formula_changed",
                    "formula_added",
                }
            ):
                continue
            factor = self._factor(change, risk_type)
            key = (risk_type, change.location.casefold())
            factors_by_key.setdefault(key, factor)

        unique_cells = {
            (change.sheet.casefold(), change.coordinate.replace("$", "").upper())
            for change in cell_changes
        }
        if len(unique_cells) >= MANY_CELLS_THRESHOLD:
            risk_type = "many_cells_changed"
            factor = RiskFactor(
                risk_type=risk_type,
                location="Workbook",
                points=self.weights[risk_type],
                description=(
                    f"{len(unique_cells)} cells changed, meeting the bulk-change threshold "
                    f"of {MANY_CELLS_THRESHOLD}."
                ),
                evidence={
                    "changed_cell_count": len(unique_cells),
                    "threshold": MANY_CELLS_THRESHOLD,
                },
            )
            factors_by_key.setdefault((risk_type, "workbook"), factor)

        factors = list(factors_by_key.values())
        total = min(100, sum(factor.points for factor in factors))
        return total, risk_level_for_score(total), factors

    def _risk_type(self, change: RiskChange) -> str | None:
        risk_type = canonical_risk_weight_name(change.change_type)
        return risk_type or None

    def _factor(self, change: RiskChange, risk_type: str) -> RiskFactor:
        evidence = dict(change.evidence)
        evidence["source_change_type"] = change.change_type
        evidence["source_risk_level"] = change.risk_level.value
        if isinstance(change, FormulaChange):
            evidence["supported_analysis"] = change.supported_analysis
            evidence["high_impact"] = change.high_impact
        return RiskFactor(
            risk_type=risk_type,
            location=change.location,
            points=self.weights.get(risk_type, _FALLBACK_RISK_WEIGHTS[change.risk_level]),
            description=change.description,
            evidence=evidence,
        )


__all__ = [
    "CANONICAL_RISK_WEIGHT_KEYS",
    "DEFAULT_RISK_WEIGHTS",
    "MANY_CELLS_THRESHOLD",
    "RISK_WEIGHT_ALIASES",
    "RiskScorer",
    "canonical_risk_weight_name",
    "risk_level_for_score",
    "risk_level_from_score",
]
