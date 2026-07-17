"""Built-in rule matrix and configuration validation tests."""

from __future__ import annotations

from typing import Any

import pytest
from openpyxl import Workbook

from tabulint.diff.workbook_diff import WorkbookDiffer
from tabulint.exceptions import ConfigurationError
from tabulint.models import RuleFailureStatus, RuleSpec, RuleStatus, TabulintConfig
from tabulint.parser.workbook import WorkbookParser
from tabulint.rules.builtin_rules import RuleContext, evaluate_builtin_rule
from tabulint.rules.engine import RuleEngine
from tabulint.rules.loader import load_config, validate_config_text
from tests.conftest import WorkbookFactory, add_vba_project


def _rule_before(workbook: Workbook) -> None:
    sheet = workbook.active
    sheet.title = "Main"
    sheet["A1"] = "=1+1"
    sheet["A2"] = 10


def _rule_after_clean(workbook: Workbook) -> None:
    sheet = workbook.active
    sheet.title = "Main"
    sheet["A1"] = "=1+1"
    sheet["A2"] = 11


def _rule_after_risky(workbook: Workbook) -> None:
    _rule_after_clean(workbook)
    main = workbook["Main"]
    main["B1"] = "='[External.xlsx]Data'!A1"
    hidden = workbook.create_sheet("Hidden Addition")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "review"


def _context(workbook_factory: WorkbookFactory, scenario: str) -> RuleContext:
    before_path = workbook_factory(f"{scenario}-before.xlsx", _rule_before)
    if scenario == "risky":
        after_path = workbook_factory("risky-after.xlsm", _rule_after_risky)
        add_vba_project(after_path)
    else:
        after_path = workbook_factory("clean-after.xlsx", _rule_after_clean)
    parser = WorkbookParser()
    before = parser.parse(before_path)
    after = parser.parse(after_path)
    structure, cells, _ = WorkbookDiffer().compare(before, after)
    return RuleContext(
        before=before,
        after=after,
        structure_changes=structure,
        cell_changes=cells,
        max_range_cells=100_000,
    )


RULE_CASES: list[tuple[str, dict[str, Any], str, RuleStatus]] = [
    (
        "formula-required-pass",
        {"name": "formula", "type": "formula_required", "range": "Main!A1"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "formula-required-fail",
        {"name": "formula", "type": "formula_required", "range": "Main!A1:A2"},
        "clean",
        RuleStatus.FAILED,
    ),
    (
        "allowed-range-pass",
        {
            "name": "scope",
            "type": "allowed_change_range",
            "ranges": ["Main!A1:A2"],
        },
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "allowed-range-fail",
        {
            "name": "scope",
            "type": "allowed_change_range",
            "ranges": ["Main!A1"],
        },
        "clean",
        RuleStatus.FAILED,
    ),
    (
        "external-links-pass",
        {"name": "links", "type": "no_external_links"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "external-links-fail",
        {"name": "links", "type": "no_external_links"},
        "risky",
        RuleStatus.FAILED,
    ),
    (
        "hidden-sheet-pass",
        {"name": "hidden", "type": "no_new_hidden_sheets"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "hidden-sheet-fail",
        {"name": "hidden", "type": "no_new_hidden_sheets"},
        "risky",
        RuleStatus.FAILED,
    ),
    (
        "macro-pass",
        {"name": "macro", "type": "no_macro_added"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "macro-fail",
        {"name": "macro", "type": "no_macro_added"},
        "risky",
        RuleStatus.FAILED,
    ),
    (
        "numeric-pass",
        {
            "name": "numeric",
            "type": "numeric_range",
            "target": "Main!A2",
            "min": 0,
            "max": 20,
        },
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "numeric-fail",
        {
            "name": "numeric",
            "type": "numeric_range",
            "target": "Main!A2",
            "min": 0,
            "max": 10,
        },
        "clean",
        RuleStatus.FAILED,
    ),
    (
        "required-sheet-pass",
        {"name": "required", "type": "required_sheet", "sheet": "main"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "required-sheet-fail",
        {"name": "required", "type": "required_sheet", "sheet": "Missing"},
        "clean",
        RuleStatus.FAILED,
    ),
    (
        "forbidden-sheet-pass",
        {"name": "forbidden", "type": "forbidden_sheet", "sheet": "Forbidden"},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "forbidden-sheet-fail",
        {"name": "forbidden", "type": "forbidden_sheet", "sheet": "Main"},
        "clean",
        RuleStatus.FAILED,
    ),
    (
        "changed-cell-limit-pass",
        {"name": "limit", "type": "max_changed_cells", "max": 1},
        "clean",
        RuleStatus.PASSED,
    ),
    (
        "changed-cell-limit-fail",
        {"name": "limit", "type": "max_changed_cells", "max": 0},
        "clean",
        RuleStatus.FAILED,
    ),
]


@pytest.mark.parametrize(
    ("case_id", "raw_spec", "scenario", "expected_status"),
    RULE_CASES,
    ids=[case[0] for case in RULE_CASES],
)
def test_each_builtin_rule_has_passing_and_failing_cases(
    workbook_factory: WorkbookFactory,
    case_id: str,
    raw_spec: dict[str, Any],
    scenario: str,
    expected_status: RuleStatus,
) -> None:
    del case_id
    result = evaluate_builtin_rule(
        RuleSpec.model_validate(raw_spec),
        _context(workbook_factory, scenario),
    )

    assert result.status is expected_status
    assert result.reason
    assert result.evidence


@pytest.mark.parametrize(
    "raw_spec",
    [
        {"name": "missing sheet", "type": "formula_required", "range": "Missing!A1"},
        {
            "name": "missing target",
            "type": "numeric_range",
            "target": "Main!Z99",
            "min": 0,
        },
        {
            "name": "missing allowed sheet",
            "type": "allowed_change_range",
            "ranges": ["Missing!A1"],
        },
    ],
)
def test_rule_engine_isolates_missing_sheet_and_cell_errors(
    workbook_factory: WorkbookFactory,
    raw_spec: dict[str, Any],
) -> None:
    context = _context(workbook_factory, "clean")
    config = TabulintConfig(rules=[RuleSpec.model_validate(raw_spec)])

    result = RuleEngine().evaluate(
        config,
        context.before,
        context.after,
        context.structure_changes,
        context.cell_changes,
    )[0]

    assert result.status is RuleStatus.ERROR
    assert result.evidence["error_type"] == "rule_evaluation_error"


def test_rule_engine_can_emit_configured_warning(
    workbook_factory: WorkbookFactory,
) -> None:
    context = _context(workbook_factory, "clean")
    spec = RuleSpec.model_validate(
        {
            "name": "advisory formula coverage",
            "type": "formula_required",
            "range": "Main!A1:A2",
            "failure_status": "WARNING",
        }
    )
    config = TabulintConfig(rules=[spec])

    result = RuleEngine().evaluate(
        config,
        context.before,
        context.after,
        context.structure_changes,
        context.cell_changes,
    )[0]

    assert spec.failure_status is RuleFailureStatus.WARNING
    assert result.status is RuleStatus.WARNING
    assert result.evidence["configured_failure_status"] == "WARNING"
    assert "configured as advisory" in result.reason


def test_numeric_formula_without_cached_value_is_skipped(
    workbook_factory: WorkbookFactory,
) -> None:
    def configure(workbook: Workbook) -> None:
        workbook.active.title = "Main"
        workbook.active["A1"] = "=1+1"

    parser = WorkbookParser()
    before = parser.parse(workbook_factory("formula-before.xlsx", configure))
    after = parser.parse(workbook_factory("formula-after.xlsx", configure))
    config = TabulintConfig(
        rules=[
            RuleSpec(
                name="formula numeric",
                type="numeric_range",
                target="Main!A1",
                min=0,
                max=10,
            )
        ]
    )

    result = RuleEngine().evaluate(config, before, after, [], [])[0]

    assert result.status is RuleStatus.SKIPPED
    assert result.evidence["value_source"] == "cached_formula_value"
    assert result.evidence["value"] is None
    assert "no cached value" in result.reason


@pytest.mark.parametrize(
    "text",
    [
        "rules: [",
        "- not-a-mapping",
        "rules:\n  - name: typo\n    type: not_a_rule",
        "rules:\n  - name: missing range\n    type: formula_required",
        "rules: []\nunexpected: true",
        "risk_weights:\n  formula_overwritten: 101",
    ],
)
def test_configuration_errors_are_actionable(text: str) -> None:
    with pytest.raises(ConfigurationError, match=r"(?i:configuration|yaml)"):
        validate_config_text(text)


def test_load_config_defaults_and_missing_path(tmp_path: Any) -> None:
    assert load_config().rules == []

    with pytest.raises(ConfigurationError, match="does not exist"):
        load_config(tmp_path / "missing.yml")


@pytest.mark.parametrize(
    "text",
    [
        "risk_weights:\n  formula_overwritten: true",
        "max_dependency_depth: true",
        "max_dependency_nodes: false",
        (
            "rules:\n"
            "  - name: numeric minimum\n"
            "    type: numeric_range\n"
            "    target: Main!A1\n"
            "    min: true"
        ),
        (
            "rules:\n"
            "  - name: numeric maximum\n"
            "    type: numeric_range\n"
            "    target: Main!A1\n"
            "    max: false"
        ),
        (
            "rules:\n"
            "  - name: changed cells\n"
            "    type: max_changed_cells\n"
            "    max: true"
        ),
    ],
)
def test_configuration_rejects_boolean_numeric_values(text: str) -> None:
    with pytest.raises(ConfigurationError, match="Invalid Tabulint configuration"):
        validate_config_text(text)


@pytest.mark.parametrize(
    "text",
    [
        'risk_weights:\n  formula_overwritten: "30"',
        'max_dependency_depth: "10"',
        'max_dependency_nodes: "100"',
        (
            "rules:\n"
            "  - name: strict numeric\n"
            "    type: numeric_range\n"
            "    target: Main!A1\n"
            '    min: "0"'
        ),
    ],
)
def test_configuration_rejects_coercible_numeric_strings(text: str) -> None:
    with pytest.raises(ConfigurationError, match="Invalid Tabulint configuration"):
        validate_config_text(text)


@pytest.mark.parametrize("limit", [".nan", ".inf", "-.inf"])
def test_numeric_range_limits_must_be_finite(limit: str) -> None:
    text = (
        "rules:\n"
        "  - name: finite numeric\n"
        "    type: numeric_range\n"
        "    target: Main!A1\n"
        f"    min: {limit}"
    )

    with pytest.raises(ConfigurationError, match=r"(?i)finite"):
        validate_config_text(text)


def test_risk_weight_names_accept_canonical_keys_and_public_aliases() -> None:
    config = validate_config_text(
        "risk_weights:\n"
        "  formula_overwritten: 31\n"
        "  formula_to_fixed_value: 27\n"
        "  format-changed: 3"
    )

    assert config.risk_weights == {
        "formula_overwritten": 31,
        "formula_to_fixed_value": 27,
        "format-changed": 3,
    }


def test_risk_weight_name_typo_is_rejected() -> None:
    with pytest.raises(
        ConfigurationError,
        match=r"unknown risk weight name.*formula_overwriten",
    ):
        validate_config_text("risk_weights:\n  formula_overwriten: 30")


@pytest.mark.parametrize(
    ("text", "duplicate_name"),
    [
        (
            "risk_weights: {}\n"
            "risk_weights:\n"
            "  formula_overwritten: 30",
            "risk_weights",
        ),
        (
            "risk_weights:\n"
            "  formula_overwritten: 30\n"
            "  formula_overwritten: 20",
            "formula_overwritten",
        ),
        (
            "rules:\n"
            "  - name: duplicate field\n"
            "    type: max_changed_cells\n"
            "    max: 1\n"
            "    max: 2",
            "max",
        ),
    ],
)
def test_yaml_duplicate_mapping_keys_are_rejected(
    text: str,
    duplicate_name: str,
) -> None:
    with pytest.raises(
        ConfigurationError,
        match=rf"Invalid YAML configuration: .*duplicate key '{duplicate_name}'",
    ):
        validate_config_text(text)
