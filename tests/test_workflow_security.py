"""Security invariants for GitHub Actions and the published composite action."""

from __future__ import annotations

import re
from pathlib import Path

import yaml

ROOT = Path(__file__).parents[1]
EXTERNAL_ACTION = re.compile(r"uses:\s+([^\s#]+)")
IMMUTABLE_ACTION = re.compile(r"[^/@\s]+/[^/@\s]+@[0-9a-f]{40}")


def _automation_files() -> list[Path]:
    return [
        *sorted((ROOT / ".github" / "workflows").glob("*.yml")),
        ROOT / "action" / "action.yml",
    ]


def test_external_actions_are_pinned_to_full_commit_shas() -> None:
    unpinned: list[str] = []

    for path in _automation_files():
        for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            match = EXTERNAL_ACTION.search(line)
            if match is None or match.group(1).startswith("./"):
                continue
            if IMMUTABLE_ACTION.fullmatch(match.group(1)) is None:
                unpinned.append(f"{path.relative_to(ROOT)}:{line_number}: {match.group(1)}")

    assert unpinned == []


def test_release_publish_job_uses_trusted_publishing_only() -> None:
    workflow = (ROOT / ".github" / "workflows" / "release.yml").read_text(
        encoding="utf-8"
    )

    assert "types: [published]" in workflow
    assert "environment:\n      name: pypi" in workflow
    assert "id-token: write" in workflow
    assert "pypa/gh-action-pypi-publish@" in workflow
    assert re.search(r"\$\{\{\s*secrets\.", workflow) is None
    assert "password:" not in workflow
    assert "pull_request_target" not in workflow


def test_pull_request_gate_keeps_head_checkout_as_data_only() -> None:
    workflow = (ROOT / ".github" / "workflows" / "sheetproof.yml").read_text(
        encoding="utf-8"
    )

    assert "pull_request_target:" in workflow
    assert "contents: read" in workflow
    assert "persist-credentials: false" in workflow
    assert "path: .sheetproof-trusted" in workflow
    assert "path: .sheetproof-pr" in workflow
    assert "uses: ./.sheetproof-trusted/action" in workflow
    assert re.search(r"\$\{\{\s*secrets\.", workflow) is None


def test_dependabot_tracks_python_and_github_actions_dependencies() -> None:
    configuration = yaml.safe_load(
        (ROOT / ".github" / "dependabot.yml").read_text(encoding="utf-8")
    )

    assert configuration["version"] == 2
    assert {update["package-ecosystem"] for update in configuration["updates"]} == {
        "github-actions",
        "pip",
    }
