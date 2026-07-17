"""Typer command-line interface for local and CI use."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Annotated, NoReturn

import typer
from pydantic import ValidationError

from tabulint import __version__
from tabulint.exceptions import ConfigurationError, TabulintError
from tabulint.models import RiskLevel, RuleStatus
from tabulint.parser.workbook import WorkbookParser
from tabulint.reports.html_report import write_html_report
from tabulint.reports.json_report import write_json_report
from tabulint.rules.loader import load_config
from tabulint.services.review_service import ReviewService

app = typer.Typer(
    name="tabulint",
    help="Deterministic semantic review and CI for Excel workbook changes.",
    no_args_is_help=True,
    pretty_exceptions_enable=False,
)
rules_app = typer.Typer(help="Validate and work with tabulint.yml rules.", no_args_is_help=True)
app.add_typer(rules_app, name="rules")

_RISK_ORDER = {
    RiskLevel.LOW: 0,
    RiskLevel.MEDIUM: 1,
    RiskLevel.HIGH: 2,
    RiskLevel.CRITICAL: 3,
}


def _paths_collide(left: Path, right: Path) -> bool:
    left_resolved = left.expanduser().resolve(strict=False)
    right_resolved = right.expanduser().resolve(strict=False)
    if left_resolved == right_resolved:
        return True
    try:
        return left.exists() and right.exists() and os.path.samefile(left, right)
    except OSError:
        return False


def _validate_output_paths(inputs: list[Path], outputs: list[Path]) -> None:
    for index, output in enumerate(outputs):
        if any(_paths_collide(output, input_path) for input_path in inputs):
            raise ValueError(f"Output path would overwrite an input file: {output}")
        if any(_paths_collide(output, other) for other in outputs[index + 1 :]):
            raise ValueError("JSON and HTML outputs must use different paths")


def _fail(message: str, code: int) -> NoReturn:
    typer.echo(f"Error: {message}", err=True)
    raise typer.Exit(code=code)


@app.command("compare")
def compare_command(
    before: Annotated[Path, typer.Argument(help="Workbook before the change (.xlsx/.xlsm).")],
    after: Annotated[Path, typer.Argument(help="Workbook after the change (.xlsx/.xlsm).")],
    config: Annotated[
        Path | None, typer.Option("--config", "-c", help="Path to tabulint.yml.")
    ] = None,
    json_output: Annotated[
        Path | None, typer.Option("--json", help="Write the structured JSON result.")
    ] = None,
    html_output: Annotated[
        Path | None, typer.Option("--html", help="Write a self-contained HTML report.")
    ] = None,
    fail_on: Annotated[
        RiskLevel | None,
        typer.Option("--fail-on", help="Return exit code 1 at or above this risk level."),
    ] = None,
    debug: Annotated[
        bool, typer.Option("--debug", help="Show an unexpected Python traceback.")
    ] = False,
) -> None:
    """Compare two workbooks and optionally produce JSON and HTML reports."""

    try:
        outputs = [path for path in (json_output, html_output) if path is not None]
        inputs = [before, after, *([config] if config is not None else [])]
        _validate_output_paths(inputs, outputs)
        service = ReviewService()
        result = service.review(before, after, config)
        validated_config = service.load_config(config)
        if json_output is not None:
            write_json_report(result, json_output)
        if html_output is not None:
            write_html_report(result, html_output)
    except (TabulintError, ValidationError, OSError, ValueError) as exc:
        _fail(str(exc), 2)
    except Exception as exc:
        if debug:
            raise
        _fail(f"internal failure ({type(exc).__name__}): {exc}", 3)

    typer.echo(
        f"Risk {result.summary.risk_score}/100 ({result.summary.risk_level.value}); "
        f"{result.summary.changed_cells} changed cells; "
        f"{result.summary.formula_overwrites} formula overwrites; "
        f"{result.summary.rules_failed} failed, "
        f"{result.summary.rules_warnings} warning, and "
        f"{result.summary.rules_errors} errored rules."
    )
    if json_output is not None:
        typer.echo(f"JSON report: {json_output.resolve()}")
    if html_output is not None:
        typer.echo(f"HTML report: {html_output.resolve()}")

    threshold = fail_on or validated_config.block_risk_level
    failed_rule = any(
        item.status in {RuleStatus.FAILED, RuleStatus.ERROR} for item in result.rule_results
    )
    high_finding = any(
        _RISK_ORDER[item.risk_level] >= _RISK_ORDER[threshold]
        for item in result.structure_changes
    ) or any(
        _RISK_ORDER[item.risk_level] >= _RISK_ORDER[threshold]
        for item in result.cell_changes
    ) or any(
        _RISK_ORDER[item.risk_level] >= _RISK_ORDER[threshold]
        for item in result.formula_changes
    )
    if (
        failed_rule
        or high_finding
        or _RISK_ORDER[result.summary.risk_level] >= _RISK_ORDER[threshold]
    ):
        raise typer.Exit(code=1)


@app.command("inspect")
def inspect_command(
    workbook: Annotated[Path, typer.Argument(help="Workbook to inspect (.xlsx/.xlsm).")],
    json_output: Annotated[
        Path | None, typer.Option("--json", help="Write the parsed workbook snapshot.")
    ] = None,
    debug: Annotated[bool, typer.Option("--debug")] = False,
) -> None:
    """Inspect workbook structure without comparing or calculating formulas."""

    try:
        if json_output is not None:
            _validate_output_paths([workbook], [json_output])
        snapshot = WorkbookParser().parse(workbook)
        if json_output is not None:
            json_output.parent.mkdir(parents=True, exist_ok=True)
            json_output.write_text(
                json.dumps(snapshot.model_dump(mode="json"), indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
    except (TabulintError, OSError, ValueError) as exc:
        _fail(str(exc), 2)
    except Exception as exc:
        if debug:
            raise
        _fail(f"internal failure ({type(exc).__name__}): {exc}", 3)

    typer.echo(
        f"{snapshot.file.name}: {len(snapshot.sheets)} sheets, "
        f"{snapshot.parsed_cell_count} materialized cells, "
        f"{len(snapshot.external_links)} external links, VBA={snapshot.has_vba}."
    )
    for sheet in snapshot.sheets:
        typer.echo(
            f"- {sheet.index + 1}. {sheet.name} [{sheet.state}] "
            f"cells={len(sheet.cells)} formulas="
            f"{sum(cell.formula is not None for cell in sheet.cells.values())}"
        )


@rules_app.command("validate")
def validate_rules_command(
    config: Annotated[Path, typer.Argument(help="Path to tabulint.yml.")],
) -> None:
    """Validate configuration syntax and all built-in rule fields."""

    try:
        parsed = load_config(config)
    except (ConfigurationError, ValidationError, OSError, ValueError) as exc:
        _fail(str(exc), 2)
    typer.echo(f"Valid configuration: {len(parsed.rules)} rules.")


@app.command("version")
def version_command() -> None:
    """Print the installed Tabulint version."""

    typer.echo(__version__)


def main() -> None:
    app()


if __name__ == "__main__":
    main()
