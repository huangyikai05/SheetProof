"""Single orchestration path used by the CLI, reports, web UI, and CI."""

from __future__ import annotations

from collections import Counter
from collections.abc import Sequence
from datetime import UTC, datetime
from pathlib import Path

from tabulint import __version__
from tabulint.diff.workbook_diff import WorkbookDiffer
from tabulint.exceptions import ConfigurationError
from tabulint.graph.impact_analysis import DependencyAnalyzer
from tabulint.models import (
    FormulaChange,
    ReviewResult,
    ReviewSummary,
    RuleStatus,
    TabulintConfig,
    WorkbookSnapshot,
)
from tabulint.parser.workbook import WorkbookParser
from tabulint.risk.scorer import RiskScorer
from tabulint.rules.builtin_rules import parse_range
from tabulint.rules.engine import RuleEngine
from tabulint.rules.loader import load_config

_FORMULA_OVERWRITE_TYPES = {"formula_overwritten", "formula_deleted"}


class ReviewService:
    """Coordinate deterministic review components without interface-specific logic."""

    def __init__(self, *, parser: WorkbookParser | None = None) -> None:
        self.parser = parser or WorkbookParser()

    def review(
        self,
        before_path: str | Path,
        after_path: str | Path,
        config_path: str | Path | None = None,
    ) -> ReviewResult:
        config = load_config(config_path)
        before = self.parser.parse(before_path)
        after = self.parser.parse(after_path)
        self._validate_workbook_config(config, after)

        structure_changes, cell_changes, formula_changes = WorkbookDiffer().compare(before, after)
        changed_formula_locations = sorted({item.location for item in formula_changes})
        dependency_impacts = DependencyAnalyzer(
            max_depth=config.max_dependency_depth,
            max_nodes=config.max_dependency_nodes,
        ).analyze(
            after,
            changed_formula_locations,
            config.critical_cells,
            before_snapshot=before,
        )

        rule_results = RuleEngine().evaluate(
            config,
            before,
            after,
            structure_changes,
            cell_changes,
        )
        risk_score, risk_level, risk_factors = RiskScorer(config.risk_weights).score(
            structure_changes,
            cell_changes,
            formula_changes,
        )
        statuses = Counter(item.status for item in rule_results)

        summary = ReviewSummary(
            risk_score=risk_score,
            risk_level=risk_level,
            changed_cells=len(cell_changes),
            changed_formulas=len(formula_changes),
            formula_overwrites=sum(
                item.change_type in _FORMULA_OVERWRITE_TYPES for item in formula_changes
            ),
            added_hidden_sheets=sum(
                item.change_type == "hidden_sheet_added" for item in structure_changes
            ),
            added_external_links=sum(
                item.change_type == "external_link_added" for item in structure_changes
            ),
            macro_status_changed=any(
                item.change_type in {"macro_added", "macro_removed"}
                for item in structure_changes
            ),
            rules_passed=statuses[RuleStatus.PASSED],
            rules_failed=statuses[RuleStatus.FAILED],
            rules_warnings=statuses[RuleStatus.WARNING],
            rules_skipped=statuses[RuleStatus.SKIPPED],
            rules_errors=statuses[RuleStatus.ERROR],
        )

        return ReviewResult(
            tool_version=__version__,
            reviewed_at=datetime.now(UTC),
            before_file=before.file,
            after_file=after.file,
            summary=summary,
            structure_changes=structure_changes,
            cell_changes=cell_changes,
            formula_changes=formula_changes,
            dependency_impacts=dependency_impacts,
            rule_results=rule_results,
            risk_factors=risk_factors,
            limitations=self._limitations(formula_changes),
            errors=[],
        )

    @staticmethod
    def load_config(config_path: str | Path | None) -> TabulintConfig:
        """Expose validated configuration for interfaces that need gating settings."""

        return load_config(config_path)

    @staticmethod
    def _limitations(formula_changes: Sequence[FormulaChange]) -> list[str]:
        limitations = [
            "Tabulint does not recalculate formulas or emulate the Excel calculation engine.",
            "Complex formulas may receive reference-level analysis only.",
            "VBA is detected but never executed.",
            "External links are reported but never opened or fetched.",
            "Risk scores support human review and are not a substitute for a professional audit.",
        ]
        if any(not getattr(item, "supported_analysis", True) for item in formula_changes):
            limitations.append(
                "One or more changed formulas were marked unsupported_formula_analysis."
            )
        return limitations

    @staticmethod
    def _validate_workbook_config(
        config: TabulintConfig,
        after: WorkbookSnapshot,
    ) -> None:
        sheet_names = {sheet.name.casefold() for sheet in after.sheets}
        missing = sorted(
            {
                parsed.sheet
                for value in config.critical_cells
                if (parsed := parse_range(value, require_single_cell=True)).sheet.casefold()
                not in sheet_names
            },
            key=str.casefold,
        )
        if missing:
            raise ConfigurationError(
                "critical_cells refers to missing worksheet(s): "
                + ", ".join(repr(name) for name in missing)
            )
