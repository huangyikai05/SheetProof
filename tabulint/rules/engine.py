"""Rule-engine orchestration with per-rule failure isolation."""

from __future__ import annotations

from tabulint.models import (
    CellChange,
    RuleFailureStatus,
    RuleResult,
    RuleStatus,
    StructureChange,
    TabulintConfig,
    WorkbookSnapshot,
)
from tabulint.rules.builtin_rules import (
    RuleContext,
    RuleEvaluationError,
    error_result,
    evaluate_builtin_rule,
)


class RuleEngine:
    """Evaluate validated business rules against immutable workbook facts."""

    def __init__(self, *, max_range_cells: int = 100_000) -> None:
        if (
            isinstance(max_range_cells, bool)
            or not isinstance(max_range_cells, int)
            or max_range_cells < 1
        ):
            raise ValueError("max_range_cells must be a positive integer")
        self.max_range_cells = max_range_cells

    def evaluate(
        self,
        config: TabulintConfig,
        before: WorkbookSnapshot,
        after: WorkbookSnapshot,
        structure_changes: list[StructureChange],
        cell_changes: list[CellChange],
    ) -> list[RuleResult]:
        """Evaluate all configured rules in declaration order.

        A malformed target or an unexpected evaluator failure affects only the
        corresponding rule and is emitted as an explicit ``ERROR`` result.
        """

        context = RuleContext(
            before=before,
            after=after,
            structure_changes=structure_changes,
            cell_changes=cell_changes,
            max_range_cells=self.max_range_cells,
        )
        results: list[RuleResult] = []
        for spec in config.rules:
            try:
                result = evaluate_builtin_rule(spec, context)
                if (
                    result.status is RuleStatus.FAILED
                    and spec.failure_status is RuleFailureStatus.WARNING
                ):
                    evidence = {
                        **result.evidence,
                        "configured_failure_status": RuleStatus.WARNING.value,
                    }
                    result = result.model_copy(
                        update={
                            "status": RuleStatus.WARNING,
                            "reason": f"{result.reason} This rule is configured as advisory.",
                            "evidence": evidence,
                        }
                    )
                results.append(result)
            except RuleEvaluationError as exc:
                results.append(error_result(spec, str(exc), error_type="rule_evaluation_error"))
            except Exception as exc:  # pragma: no cover - defensive isolation boundary
                results.append(
                    error_result(
                        spec,
                        f"Rule evaluation failed unexpectedly: {exc}",
                        error_type=type(exc).__name__,
                    )
                )
        return results


__all__ = ["RuleEngine"]
