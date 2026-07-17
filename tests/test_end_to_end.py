"""End-to-end review of the source-generated public demo workbooks."""

from __future__ import annotations

from pathlib import Path

from examples.generate_demo_workbooks import generate
from tabulint.models import RiskLevel, RuleStatus
from tabulint.reports.html_report import render_html
from tabulint.reports.json_report import render_json
from tabulint.services.review_service import ReviewService


def test_demo_generator_and_review_pipeline(tmp_path: Path) -> None:
    before, safe, risky = generate(tmp_path / "generated")
    config = Path(__file__).parents[1] / "examples" / "tabulint.example.yml"
    service = ReviewService()

    safe_result = service.review(before, safe, config)
    risky_result = service.review(before, risky, config)

    assert [path.name for path in (before, safe, risky)] == [
        "before.xlsx",
        "after_safe.xlsx",
        "after_risky.xlsx",
    ]
    assert all(path.exists() and path.stat().st_size > 0 for path in (before, safe, risky))
    assert safe_result.summary.risk_level is RiskLevel.LOW
    assert safe_result.summary.changed_cells == 1
    assert safe_result.summary.rules_failed == 0

    assert risky_result.summary.risk_level in {RiskLevel.HIGH, RiskLevel.CRITICAL}
    assert risky_result.summary.formula_overwrites >= 1
    assert risky_result.summary.added_hidden_sheets == 1
    assert risky_result.summary.added_external_links == 1
    assert any(
        change.change_type == "formula_range_reduced"
        for change in risky_result.formula_changes
    )
    assert any(result.status is RuleStatus.FAILED for result in risky_result.rule_results)
    assert risky_result.summary.rules_warnings == 1
    assert any(result.status is RuleStatus.WARNING for result in risky_result.rule_results)
    assert "formula_overwritten" in {factor.risk_type for factor in risky_result.risk_factors}

    json_report = render_json(risky_result)
    html_report = render_html(risky_result)
    assert '"risk_score"' in json_report
    assert "Tabulint" in html_report
    assert "formula_overwritten" in html_report
