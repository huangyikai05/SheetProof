"""Typer CLI integration tests with real temporary workbooks."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook
from typer.testing import CliRunner

from tabulint.cli import app
from tests.conftest import WorkbookFactory, add_vba_project

runner = CliRunner()


def _safe_before(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = "before"


def _safe_after(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = "after"


def _high_risk_before(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = "=B1*2"
    workbook.active["B1"] = 10


def _high_risk_after(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = 20
    workbook.active["B1"] = 10
    workbook.active["C1"] = "='[External.xlsx]Data'!A1"


def _formula_overwrite_only_after(workbook: Workbook) -> None:
    workbook.active.title = "Main"
    workbook.active["A1"] = 20
    workbook.active["B1"] = 10


def test_cli_help_and_version_are_available() -> None:
    help_result = runner.invoke(app, ["--help"])
    version_result = runner.invoke(app, ["version"])

    assert help_result.exit_code == 0, help_result.output
    assert "compare" in help_result.output
    assert "inspect" in help_result.output
    assert "rules" in help_result.output
    assert version_result.exit_code == 0, version_result.output
    assert version_result.output.strip() == "0.1.0"


def test_compare_command_returns_zero_for_low_risk_change(
    workbook_factory: WorkbookFactory,
) -> None:
    before = workbook_factory("safe-before.xlsx", _safe_before)
    after = workbook_factory("safe-after.xlsx", _safe_after)

    result = runner.invoke(app, ["compare", str(before), str(after)])

    assert result.exit_code == 0, result.output
    assert "Risk 1/100 (LOW)" in result.output
    assert "1 changed cells" in result.output


def test_compare_command_returns_one_at_high_risk_threshold(
    workbook_factory: WorkbookFactory,
) -> None:
    before = workbook_factory("risky-before.xlsx", _high_risk_before)
    after = workbook_factory("risky-after.xlsx", _high_risk_after)

    result = runner.invoke(app, ["compare", str(before), str(after)])

    assert result.exit_code == 1, result.output
    assert "(HIGH)" in result.output
    assert "1 formula overwrites" in result.output


def test_single_formula_overwrite_blocks_at_default_high_finding_threshold(
    workbook_factory: WorkbookFactory,
) -> None:
    before = workbook_factory("overwrite-before.xlsx", _high_risk_before)
    after = workbook_factory("overwrite-after.xlsx", _formula_overwrite_only_after)

    result = runner.invoke(app, ["compare", str(before), str(after)])

    assert result.exit_code == 1, result.output
    assert "Risk 30/100 (MEDIUM)" in result.output
    assert "1 formula overwrites" in result.output


def test_compare_command_returns_two_for_missing_file(tmp_path: Path) -> None:
    result = runner.invoke(
        app,
        ["compare", str(tmp_path / "missing-before.xlsx"), str(tmp_path / "missing-after.xlsx")],
    )

    assert result.exit_code == 2
    assert "Error:" in result.output
    assert "does not exist" in result.output
    assert "Traceback" not in result.output


def test_compare_command_returns_two_for_bad_configuration(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    before = workbook_factory("config-before.xlsx", _safe_before)
    after = workbook_factory("config-after.xlsx", _safe_after)
    config = tmp_path / "bad.yml"
    config.write_text(
        "rules:\n  - name: Missing range\n    type: formula_required\n",
        encoding="utf-8",
    )

    result = runner.invoke(
        app,
        ["compare", str(before), str(after), "--config", str(config)],
    )

    assert result.exit_code == 2
    assert "Invalid Tabulint configuration" in result.output
    assert "requires 'range'" in result.output
    assert "Traceback" not in result.output


def test_compare_command_writes_json_and_offline_html_reports(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    before = workbook_factory("report-before.xlsx", _safe_before)
    after = workbook_factory("report-after.xlsx", _safe_after)
    json_path = tmp_path / "reports" / "result.json"
    html_path = tmp_path / "reports" / "result.html"

    result = runner.invoke(
        app,
        [
            "compare",
            str(before),
            str(after),
            "--json",
            str(json_path),
            "--html",
            str(html_path),
        ],
    )

    assert result.exit_code == 0, result.output
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["tool_version"] == "0.1.0"
    assert payload["summary"]["risk_level"] == "LOW"
    assert payload["cell_changes"][0]["change_type"] == "text_changed"
    html = html_path.read_text(encoding="utf-8")
    assert "<!doctype html>" in html.lower()
    assert "Tabulint" in html
    assert "before" in html
    assert "after" in html
    assert "cdn." not in html.lower()
    assert str(json_path.resolve()) in result.output
    assert str(html_path.resolve()) in result.output


def test_compare_blocks_relocated_vba_project_end_to_end(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    before = workbook_factory("macro-before.xlsx", _safe_before)
    after = workbook_factory("macro-after.xlsx", _safe_before)
    add_vba_project(after, part_name="xl/custom/project.bin")
    config = tmp_path / "macro-policy.yml"
    config.write_text(
        "rules:\n  - name: No new macros\n    type: no_macro_added\n",
        encoding="utf-8",
    )
    json_path = tmp_path / "macro-result.json"

    result = runner.invoke(
        app,
        [
            "compare",
            str(before),
            str(after),
            "--config",
            str(config),
            "--json",
            str(json_path),
        ],
    )

    assert result.exit_code == 1, result.output
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["before_file"]["has_vba"] is False
    assert payload["after_file"]["has_vba"] is True
    assert payload["summary"]["macro_status_changed"] is True
    macro_change = next(
        change
        for change in payload["structure_changes"]
        if change["change_type"] == "macro_added"
    )
    assert macro_change["risk_level"] == "CRITICAL"
    macro_rule = next(
        rule for rule in payload["rule_results"] if rule["rule_type"] == "no_macro_added"
    )
    assert macro_rule["status"] == "FAILED"
    assert macro_rule["evidence"]["macro_added"] is True
    assert "macro_added" in {factor["risk_type"] for factor in payload["risk_factors"]}


def test_inspect_and_rules_validate_commands(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    workbook = workbook_factory("inspect.xlsx", _safe_before)
    snapshot_path = tmp_path / "snapshot.json"
    inspect_result = runner.invoke(
        app,
        ["inspect", str(workbook), "--json", str(snapshot_path)],
    )
    assert inspect_result.exit_code == 0, inspect_result.output
    assert "1 sheets" in inspect_result.output
    assert json.loads(snapshot_path.read_text(encoding="utf-8"))["sheets"][0]["name"] == "Main"

    config = tmp_path / "tabulint.yml"
    config.write_text("rules: []\n", encoding="utf-8")
    rules_result = runner.invoke(app, ["rules", "validate", str(config)])
    assert rules_result.exit_code == 0, rules_result.output
    assert "Valid configuration: 0 rules" in rules_result.output


@pytest.mark.parametrize(
    ("config_text", "expected_message"),
    [
        (
            "rules:\n"
            "  - name: Bad range\n"
            "    type: formula_required\n"
            "    range: Main!A0\n",
            "outside Excel worksheet limits",
        ),
        (
            "rules:\n"
            "  - name: Reversed bounds\n"
            "    type: numeric_range\n"
            "    target: Main!A1\n"
            "    min: 10\n"
            "    max: 1\n",
            "min cannot be greater than max",
        ),
        (
            "critical_cells:\n"
            "  - Main!A1:B2\n",
            "critical_cells item 1",
        ),
    ],
)
def test_rules_validate_rejects_semantically_invalid_locations_and_bounds(
    tmp_path: Path,
    config_text: str,
    expected_message: str,
) -> None:
    config = tmp_path / "invalid-semantic.yml"
    config.write_text(config_text, encoding="utf-8")

    result = runner.invoke(app, ["rules", "validate", str(config)])

    assert result.exit_code == 2
    assert expected_message in result.output
    assert "Traceback" not in result.output


def test_inspect_refuses_json_output_that_would_overwrite_input(
    workbook_factory: WorkbookFactory,
) -> None:
    workbook = workbook_factory("inspect-collision.xlsx", _safe_before)
    original = workbook.read_bytes()

    result = runner.invoke(
        app,
        ["inspect", str(workbook), "--json", str(workbook)],
    )

    assert result.exit_code == 2
    assert "Output path would overwrite an input file" in result.output
    assert workbook.read_bytes() == original


@pytest.mark.parametrize(
    ("output_option", "collision_target"),
    [
        ("--json", "before"),
        ("--html", "after"),
        ("--json", "config"),
    ],
)
def test_compare_refuses_outputs_that_collide_with_inputs_or_config(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
    output_option: str,
    collision_target: str,
) -> None:
    before = workbook_factory("collision-before.xlsx", _safe_before)
    after = workbook_factory("collision-after.xlsx", _safe_after)
    config = tmp_path / "collision-config.yml"
    config.write_text("rules: []\n", encoding="utf-8")
    targets = {"before": before, "after": after, "config": config}
    originals = {path: path.read_bytes() for path in targets.values()}

    result = runner.invoke(
        app,
        [
            "compare",
            str(before),
            str(after),
            "--config",
            str(config),
            output_option,
            str(targets[collision_target]),
        ],
    )

    assert result.exit_code == 2
    assert "Output path would overwrite an input file" in result.output
    assert all(path.read_bytes() == content for path, content in originals.items())


def test_compare_refuses_same_json_and_html_output_without_writing(
    workbook_factory: WorkbookFactory,
    tmp_path: Path,
) -> None:
    before = workbook_factory("same-output-before.xlsx", _safe_before)
    after = workbook_factory("same-output-after.xlsx", _safe_after)
    output = tmp_path / "same-output.report"
    before_bytes = before.read_bytes()
    after_bytes = after.read_bytes()

    result = runner.invoke(
        app,
        [
            "compare",
            str(before),
            str(after),
            "--json",
            str(output),
            "--html",
            str(output),
        ],
    )

    assert result.exit_code == 2
    assert "JSON and HTML outputs must use different paths" in result.output
    assert not output.exists()
    assert before.read_bytes() == before_bytes
    assert after.read_bytes() == after_bytes
