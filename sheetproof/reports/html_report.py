"""Render the canonical review result as a self-contained HTML report."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from jinja2 import Environment, PackageLoader, StrictUndefined, select_autoescape

from sheetproof.exceptions import ReportGenerationError
from sheetproof.models import (
    CellChange,
    FormulaChange,
    ReviewResult,
    RiskLevel,
    RuleStatus,
    Severity,
    StructureChange,
)

_TEMPLATE_NAME = "report.html.j2"


def _format_value(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, dict | list | tuple):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value)


def _pretty_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True)


def _environment() -> Environment:
    environment = Environment(
        loader=PackageLoader("sheetproof.reports", "templates"),
        autoescape=select_autoescape(
            enabled_extensions=("html", "htm", "xml", "j2"),
            default_for_string=True,
            default=True,
        ),
        undefined=StrictUndefined,
        trim_blocks=True,
        lstrip_blocks=True,
    )
    environment.filters["display"] = _format_value
    environment.filters["pretty_json"] = _pretty_json
    return environment


def _high_risk_findings(result: ReviewResult) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    high_levels = {RiskLevel.HIGH, RiskLevel.CRITICAL}

    def add_change(
        category: str,
        change: StructureChange | CellChange | FormulaChange,
    ) -> None:
        if change.risk_level in high_levels:
            findings.append(
                {
                    "category": category,
                    "level": change.risk_level.value,
                    "location": change.location,
                    "description": change.description,
                    "evidence": change.evidence,
                }
            )

    for structure_change in result.structure_changes:
        add_change("Structure", structure_change)
    for cell_change in result.cell_changes:
        add_change("Cell", cell_change)
    for formula_change in result.formula_changes:
        add_change("Formula", formula_change)

    for rule in result.rule_results:
        if rule.status in {
            RuleStatus.FAILED,
            RuleStatus.WARNING,
            RuleStatus.ERROR,
        } and rule.severity in {
            Severity.HIGH,
            Severity.CRITICAL,
        }:
            findings.append(
                {
                    "category": "Rule",
                    "level": rule.severity.value.upper(),
                    "location": rule.location or "—",
                    "description": f"{rule.name}: {rule.reason}",
                    "evidence": rule.evidence,
                }
            )
    return findings


def render_html(result: ReviewResult) -> str:
    """Return a fully offline HTML report generated with autoescaping enabled."""

    try:
        template = _environment().get_template(_TEMPLATE_NAME)
        payload = result.model_dump(mode="json")
        formula_overwrites = [
            change
            for change in payload["formula_changes"]
            if change["before_formula"] is not None and change["after_formula"] is None
        ]
        skipped_or_error_rules = [
            rule
            for rule in payload["rule_results"]
            if rule["status"] in {RuleStatus.SKIPPED.value, RuleStatus.ERROR.value}
        ]
        return template.render(
            report=payload,
            high_risk_findings=_high_risk_findings(result),
            formula_overwrites=formula_overwrites,
            skipped_or_error_rules=skipped_or_error_rules,
        )
    except Exception as exc:
        if isinstance(exc, ReportGenerationError):
            raise
        raise ReportGenerationError(f"Unable to render HTML report: {exc}") from exc


def write_html_report(result: ReviewResult, path: str | Path) -> Path:
    """Write a self-contained HTML report to ``path`` and return its resolved path."""

    target = Path(path).expanduser()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(render_html(result), encoding="utf-8", newline="\n")
    except OSError as exc:
        raise ReportGenerationError(f"Unable to write HTML report to {target}: {exc}") from exc
    return target.resolve()
