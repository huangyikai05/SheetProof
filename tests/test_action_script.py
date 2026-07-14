"""Regression tests for the GitHub Action workbook-pair policy."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from scripts import review_changed_workbooks as review_script


def _git(repo: Path, *arguments: str) -> str:
    completed = subprocess.run(
        ["git", *arguments],
        cwd=repo,
        check=True,
        text=True,
        encoding="utf-8",
        capture_output=True,
    )
    return completed.stdout.strip()


@pytest.fixture
def unpaired_changes() -> list[review_script.GitChange]:
    return [
        review_script.GitChange("A", None, "added.xlsx"),
        review_script.GitChange("D", "deleted.xlsm", None),
        review_script.GitChange("C100", "source.xlsx", "copy.xlsx"),
        review_script.GitChange("R100", "notes.txt", "renamed.xlsx"),
    ]


@pytest.mark.parametrize("allow_unpaired", [False, True])
def test_unpaired_workbook_changes_have_explicit_gate_outcomes(
    unpaired_changes: list[review_script.GitChange],
    allow_unpaired: bool,
) -> None:
    outcomes = [
        review_script._unpaired_outcome(change, allow_unpaired=allow_unpaired)
        for change in unpaired_changes
    ]

    assert all(outcome is not None for outcome in outcomes)
    concrete_outcomes = [outcome for outcome in outcomes if outcome is not None]
    expected_status = "SKIPPED" if allow_unpaired else "UNREVIEWABLE"
    expected_exit_code = 0 if allow_unpaired else 1
    assert [outcome.status for outcome in concrete_outcomes] == [expected_status] * 4
    assert [outcome.exit_code for outcome in concrete_outcomes] == [
        expected_exit_code
    ] * 4
    assert review_script._aggregate_exit_code(concrete_outcomes) == expected_exit_code


def test_aggregate_exit_code_uses_highest_failure_class() -> None:
    outcomes = [
        review_script.ReviewOutcome("SKIPPED", "added.xlsx", 0, "allowed"),
        review_script.ReviewOutcome("UNREVIEWABLE", "deleted.xlsx", 1, "unpaired"),
        review_script.ReviewOutcome("ERROR", "invalid.xlsx", 2, "invalid input"),
        review_script.ReviewOutcome("ERROR", "internal.xlsx", 3, "internal error"),
    ]

    assert review_script._aggregate_exit_code(outcomes[:1]) == 0
    assert review_script._aggregate_exit_code(outcomes[:2]) == 1
    assert review_script._aggregate_exit_code(outcomes[:3]) == 2
    assert review_script._aggregate_exit_code(outcomes) == 3


def test_diverged_pull_request_uses_merge_base_for_discovery_and_before_blob(
    tmp_path: Path,
) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _git(repo, "init", "-b", "main")
    _git(repo, "config", "user.name", "SheetProof tests")
    _git(repo, "config", "user.email", "sheetproof@example.invalid")

    workbook = repo / "forecast.xlsx"
    workbook.write_bytes(b"common ancestor workbook")
    _git(repo, "add", "forecast.xlsx")
    _git(repo, "commit", "-m", "common")
    common = _git(repo, "rev-parse", "HEAD")

    _git(repo, "checkout", "-b", "feature")
    workbook.write_bytes(b"pull request workbook")
    _git(repo, "commit", "-am", "change workbook in pull request")
    head = _git(repo, "rev-parse", "HEAD")

    _git(repo, "checkout", "main")
    workbook.write_bytes(b"new destination-branch workbook")
    _git(repo, "commit", "-am", "change workbook on destination branch")
    trusted_base = _git(repo, "rev-parse", "HEAD")

    comparison_base = review_script._merge_base(repo, trusted_base, head)
    assert comparison_base == common
    assert comparison_base != trusted_base
    assert review_script._changed_workbooks(repo, comparison_base, head) == [
        review_script.GitChange("M", "forecast.xlsx", "forecast.xlsx")
    ]

    before = tmp_path / "before.xlsx"
    review_script._materialize_blob(repo, comparison_base, "forecast.xlsx", before)
    assert before.read_bytes() == b"common ancestor workbook"


def test_output_directory_is_created_once_and_must_be_absent(tmp_path: Path) -> None:
    repo = tmp_path.resolve()
    output = review_script._safe_output_directory(repo, "reports")

    assert output == repo / "reports"
    assert output.is_dir()
    with pytest.raises(review_script.ReviewScriptError, match="must not already exist"):
        review_script._safe_output_directory(repo, "reports")


def test_output_directory_rejects_preseeded_directory(tmp_path: Path) -> None:
    repo = tmp_path.resolve()
    (repo / "reports").mkdir()
    (repo / "reports" / "from-pr.txt").write_text("untrusted", encoding="utf-8")

    with pytest.raises(review_script.ReviewScriptError, match="must not already exist"):
        review_script._safe_output_directory(repo, "reports")


def test_output_directory_rejects_leaf_symlink(tmp_path: Path) -> None:
    repo = tmp_path.resolve()
    target = repo / "untrusted-artifacts"
    target.mkdir()
    leaf = repo / "reports"
    try:
        leaf.symlink_to(target, target_is_directory=True)
    except OSError as exc:
        pytest.skip(f"directory symlinks are unavailable: {exc}")

    with pytest.raises(review_script.ReviewScriptError, match="must not already exist"):
        review_script._safe_output_directory(repo, "reports")


def test_output_directory_requires_existing_real_parent(tmp_path: Path) -> None:
    repo = tmp_path.resolve()

    with pytest.raises(review_script.ReviewScriptError, match="parent must be an existing"):
        review_script._safe_output_directory(repo, "missing/reports")


@pytest.mark.parametrize(
    "unsafe_path",
    ["../reports", "/reports", "reports\\nested", "reports\ninjected=value"],
)
def test_output_directory_rejects_nonportable_or_output_injection_paths(
    tmp_path: Path,
    unsafe_path: str,
) -> None:
    with pytest.raises(review_script.ReviewScriptError):
        review_script._safe_output_directory(tmp_path.resolve(), unsafe_path)


def test_required_workflow_is_base_owned_and_uploads_only_trusted_output() -> None:
    workflow = (
        Path(__file__).resolve().parents[1] / ".github" / "workflows" / "sheetproof.yml"
    ).read_text(encoding="utf-8")

    assert "pull_request_target:" in workflow
    assert "\n  pull_request:\n" not in workflow
    assert "permissions:\n  contents: read" in workflow
    assert "secrets." not in workflow
    assert "uses: ./.sheetproof-trusted/action" in workflow
    assert "ref: refs/pull/${{ github.event.pull_request.number }}/head" in workflow
    assert "steps.sheetproof.outputs.report-directory != ''" in workflow
    assert "path: .sheetproof-pr/${{ steps.sheetproof.outputs.report-directory }}/" in workflow
