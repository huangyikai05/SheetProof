#!/usr/bin/env python3
"""Review modified or renamed Excel workbooks between two Git commits.

The script is intentionally a thin CI adapter. It discovers Git changes, reads
the exact blobs from each commit into controlled temporary filenames, invokes
SheetProof's public CLI, and aggregates reports and exit codes. It never checks
out a workbook path supplied by a pull request.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any

WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}
CONFIG_SUFFIXES = {".yml", ".yaml"}
MAX_WORKBOOK_BYTES = 100 * 1024 * 1024
SAFE_REF_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._/-]{0,255}")


class ReviewScriptError(RuntimeError):
    """Expected orchestration error carrying a SheetProof-compatible exit code."""

    def __init__(self, message: str, exit_code: int = 2) -> None:
        super().__init__(message)
        self.exit_code = exit_code


@dataclass(frozen=True, slots=True)
class GitChange:
    """One record from ``git diff --name-status``."""

    status: str
    before_path: str | None
    after_path: str | None
    path_error: str | None = None

    @property
    def display_path(self) -> str:
        if self.before_path and self.after_path and self.before_path != self.after_path:
            return f"{self.before_path} -> {self.after_path}"
        return self.after_path or self.before_path or "<unknown>"


@dataclass(frozen=True, slots=True)
class ReviewOutcome:
    """A comparison or explicit skip rendered in the aggregate summary."""

    status: str
    workbook: str
    exit_code: int
    message: str
    risk_level: str = "—"
    risk_score: int | None = None
    changed_cells: int | None = None
    formula_overwrites: int | None = None
    failed_rules: int | None = None
    json_report: str | None = None
    html_report: str | None = None


def _run_git(
    repo: Path,
    arguments: Sequence[str],
    *,
    literal_pathspecs: bool = False,
) -> bytes:
    environment = os.environ.copy()
    environment["GIT_CONFIG_NOSYSTEM"] = "1"
    if literal_pathspecs:
        environment["GIT_LITERAL_PATHSPECS"] = "1"
    try:
        completed = subprocess.run(
            ["git", *arguments],
            cwd=repo,
            env=environment,
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        raise ReviewScriptError(f"Unable to run Git: {exc}") from exc
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", errors="replace").strip()[:2000]
        command = "git " + " ".join(arguments[:3])
        raise ReviewScriptError(f"{command} failed: {detail or 'unknown Git error'}")
    return completed.stdout


def _repo_root(raw_repository: str) -> Path:
    try:
        candidate = Path(raw_repository).expanduser().resolve(strict=True)
    except OSError as exc:
        raise ReviewScriptError(f"Unable to resolve repository path: {exc}") from exc
    if not candidate.is_dir():
        raise ReviewScriptError("repository path is not a directory")
    output = _run_git(candidate, ["rev-parse", "--show-toplevel"])
    try:
        return Path(output.decode("utf-8").strip()).resolve(strict=True)
    except (OSError, UnicodeError) as exc:
        raise ReviewScriptError(f"Unable to resolve the Git repository root: {exc}") from exc


def _resolve_commit(repo: Path, reference: str, label: str) -> str:
    reference = reference.strip()
    if (
        not SAFE_REF_RE.fullmatch(reference)
        or ".." in reference
        or "@{" in reference
        or "//" in reference
        or reference.endswith("/")
    ):
        raise ReviewScriptError(f"{label} is not a safe Git commit or reference")
    output = _run_git(repo, ["rev-parse", "--verify", f"{reference}^{{commit}}"])
    commit = output.decode("ascii", errors="strict").strip()
    if not re.fullmatch(r"[0-9a-fA-F]{40}|[0-9a-fA-F]{64}", commit):
        raise ReviewScriptError(f"{label} did not resolve to a commit object")
    return commit.lower()


def _merge_base(repo: Path, trusted_base: str, head: str) -> str:
    """Return the single common ancestor used for discovery and before-side blobs.

    Pull requests can lag behind the destination branch. Comparing the current
    destination tip directly with the pull-request head would then attribute
    destination-only changes to the pull request. A unique merge base preserves
    the same semantics as the pull-request file view while the policy remains
    pinned independently to the current trusted base commit.
    """

    output = _run_git(repo, ["merge-base", "--all", trusted_base, head])
    candidates = [line.strip().lower() for line in output.decode("ascii").splitlines()]
    if len(candidates) != 1 or not re.fullmatch(
        r"[0-9a-f]{40}|[0-9a-f]{64}", candidates[0] if candidates else ""
    ):
        raise ReviewScriptError(
            "trusted base and head must have exactly one valid Git merge base"
        )
    return candidates[0]


def _decode_git_path(raw_path: bytes) -> tuple[str, str | None]:
    try:
        path = raw_path.decode("utf-8", errors="strict")
    except UnicodeDecodeError:
        path = raw_path.decode("utf-8", errors="replace")
        return path, "Git path is not valid UTF-8 and cannot be handled portably"
    try:
        _validate_repo_path(path, "workbook path")
    except ReviewScriptError as exc:
        return path, str(exc)
    return path, None


def _parse_name_status(payload: bytes) -> list[GitChange]:
    fields = payload.split(b"\0")
    if fields and fields[-1] == b"":
        fields.pop()
    changes: list[GitChange] = []
    index = 0
    while index < len(fields):
        try:
            status = fields[index].decode("ascii", errors="strict")
        except UnicodeDecodeError as exc:
            raise ReviewScriptError("Git returned a malformed change status") from exc
        index += 1
        if not status:
            raise ReviewScriptError("Git returned an empty change status")
        kind = status[0]
        path_error: str | None = None
        before_path: str | None
        after_path: str | None
        if kind in {"R", "C"}:
            if index + 1 >= len(fields):
                raise ReviewScriptError("Git returned an incomplete rename/copy record")
            before_path, before_error = _decode_git_path(fields[index])
            after_path, after_error = _decode_git_path(fields[index + 1])
            index += 2
            path_error = before_error or after_error
        else:
            if index >= len(fields):
                raise ReviewScriptError("Git returned an incomplete change record")
            path, path_error = _decode_git_path(fields[index])
            index += 1
            before_path = None if kind == "A" else path
            after_path = None if kind == "D" else path
        changes.append(
            GitChange(
                status=status,
                before_path=before_path,
                after_path=after_path,
                path_error=path_error,
            )
        )
    return changes


def _changed_workbooks(repo: Path, base: str, head: str) -> list[GitChange]:
    payload = _run_git(
        repo,
        [
            "diff",
            "--no-ext-diff",
            "--name-status",
            "-z",
            "--find-renames",
            base,
            head,
            "--",
        ],
    )
    return [
        change
        for change in _parse_name_status(payload)
        if _is_workbook(change.before_path) or _is_workbook(change.after_path)
    ]


def _is_workbook(path: str | None) -> bool:
    return path is not None and PurePosixPath(path).suffix.lower() in WORKBOOK_SUFFIXES


def _validate_repo_path(path: str, label: str) -> str:
    if not path or "\0" in path or "\\" in path:
        raise ReviewScriptError(f"{label} is empty or contains a non-portable separator")
    if any(ord(character) < 32 or 0xD800 <= ord(character) <= 0xDFFF for character in path):
        raise ReviewScriptError(f"{label} contains a control or invalid Unicode character")
    parsed = PurePosixPath(path)
    if (
        parsed.is_absolute()
        or parsed.as_posix() != path
        or any(part in {"", ".", ".."} for part in parsed.parts)
    ):
        raise ReviewScriptError(f"{label} must be a normalized repository-relative path")
    return parsed.as_posix()


def _safe_output_directory(repo: Path, raw_path: str) -> Path:
    if not raw_path.strip():
        raise ReviewScriptError("output directory cannot be empty")
    portable_path = _validate_repo_path(raw_path.strip(), "output directory")
    requested = Path(*PurePosixPath(portable_path).parts)
    lexical = repo / requested
    try:
        lexical_absolute = Path(os.path.abspath(lexical))
        lexical_relative = lexical_absolute.relative_to(repo)
    except (OSError, ValueError) as exc:
        raise ReviewScriptError("output directory must stay inside the Git repository") from exc

    cursor = repo
    for part in lexical_relative.parts[:-1]:
        cursor = cursor / part
        if cursor.is_symlink():
            raise ReviewScriptError("output directory cannot contain symlink components")

    # The leaf is intentionally checked separately because ``Path.exists`` is
    # false for a dangling symlink. Reports must never share a directory that
    # came from the untrusted checkout: otherwise a pull request could seed
    # files that would later be uploaded as trusted review artifacts.
    if lexical_absolute.is_symlink() or lexical_absolute.exists():
        raise ReviewScriptError("output directory must not already exist")

    parent = lexical_absolute.parent
    if not parent.exists() or not parent.is_dir():
        raise ReviewScriptError("output directory parent must be an existing directory")

    try:
        resolved = lexical_absolute.resolve(strict=False)
        relative = resolved.relative_to(repo)
    except (OSError, ValueError) as exc:
        raise ReviewScriptError("output directory must stay inside the Git repository") from exc
    if resolved == repo or not relative.parts:
        raise ReviewScriptError("output directory cannot be the repository root")

    try:
        # ``exist_ok=False`` is the final race-resistant ownership check. Only
        # this trusted helper creates the leaf uploaded by the workflow.
        resolved.mkdir(mode=0o700, exist_ok=False)
    except FileExistsError as exc:
        raise ReviewScriptError("output directory must not already exist") from exc
    except OSError as exc:
        raise ReviewScriptError(f"unable to create output directory: {exc}") from exc
    return resolved


def _blob_entry(repo: Path, commit: str, path: str) -> tuple[str, str] | None:
    path = _validate_repo_path(path, "Git object path")
    payload = _run_git(
        repo,
        ["ls-tree", "-z", "--full-tree", commit, "--", path],
        literal_pathspecs=True,
    )
    for record in payload.split(b"\0"):
        if not record:
            continue
        try:
            metadata, raw_name = record.split(b"\t", 1)
            mode, object_type, object_id = metadata.decode("ascii").split(" ", 2)
            name = raw_name.decode("utf-8", errors="strict")
        except (UnicodeError, ValueError) as exc:
            raise ReviewScriptError(f"Malformed Git tree entry for {path}") from exc
        if name != path:
            continue
        if object_type != "blob" or mode == "120000":
            raise ReviewScriptError(f"{path} is not a regular file in commit {commit[:12]}")
        return object_id, mode
    return None


def _materialize_blob(
    repo: Path,
    commit: str,
    source_path: str,
    destination: Path,
    *,
    max_bytes: int | None = MAX_WORKBOOK_BYTES,
) -> None:
    entry = _blob_entry(repo, commit, source_path)
    if entry is None:
        raise ReviewScriptError(
            f"{source_path} does not exist as a regular file in commit {commit[:12]}"
        )
    object_id, _ = entry
    size_payload = _run_git(repo, ["cat-file", "-s", object_id])
    try:
        size = int(size_payload.decode("ascii").strip())
    except (UnicodeError, ValueError) as exc:
        raise ReviewScriptError(f"Git returned an invalid blob size for {source_path}") from exc
    if max_bytes is not None and size > max_bytes:
        raise ReviewScriptError(
            f"{source_path} is {size:,} bytes; the configured CI limit is {max_bytes:,} bytes"
        )
    content = _run_git(repo, ["cat-file", "blob", object_id])
    if len(content) != size:
        raise ReviewScriptError(f"Git blob size changed while reading {source_path}")
    try:
        destination.write_bytes(content)
    except OSError as exc:
        raise ReviewScriptError(f"Unable to create a temporary Git blob: {exc}") from exc


def _artifact_prefix(index: int, change: GitChange) -> str:
    source = change.after_path or change.before_path or "workbook"
    stem = PurePosixPath(source).stem
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", stem).strip("-._").lower() or "workbook"
    slug = slug[:48]
    digest = hashlib.sha256(change.display_path.encode("utf-8")).hexdigest()[:10]
    return f"{index:03d}-{slug}-{digest}"


def _command_detail(completed: subprocess.CompletedProcess[str]) -> str:
    text = (completed.stderr or completed.stdout or "").strip()
    return text[:2000] or f"SheetProof exited with code {completed.returncode}"


def _run_sheetproof(arguments: Sequence[str]) -> subprocess.CompletedProcess[str]:
    # Never execute the CLI with the pull-request checkout as the current directory:
    # an untrusted ``sheetproof`` package there would shadow the trusted installation.
    trusted_source = Path(__file__).resolve().parents[1]
    environment = os.environ.copy()
    environment["PYTHONUTF8"] = "1"
    environment["PYTHONNOUSERSITE"] = "1"
    environment.pop("PYTHONPATH", None)
    environment.pop("PYTHONHOME", None)
    try:
        return subprocess.run(
            [sys.executable, "-m", "sheetproof", *arguments],
            cwd=trusted_source,
            env=environment,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        raise ReviewScriptError(f"Unable to run SheetProof: {exc}", exit_code=3) from exc


def _prepare_config(
    repo: Path,
    policy_commit: str,
    raw_config: str,
    temp_root: Path,
) -> tuple[Path | None, str]:
    if not raw_config.strip():
        return None, "No policy path configured; built-in defaults were used."
    config_path = _validate_repo_path(raw_config.strip(), "configuration path")
    if PurePosixPath(config_path).suffix.lower() not in CONFIG_SUFFIXES:
        raise ReviewScriptError("configuration path must end in .yml or .yaml")
    if _blob_entry(repo, policy_commit, config_path) is None:
        return None, (
            f"Policy `{config_path}` was not present at trusted commit "
            f"`{policy_commit[:12]}`; built-in defaults were used."
        )
    destination = temp_root / "sheetproof-policy.yml"
    _materialize_blob(repo, policy_commit, config_path, destination, max_bytes=1024 * 1024)
    validation = _run_sheetproof(["rules", "validate", str(destination)])
    if validation.returncode != 0:
        code = validation.returncode if validation.returncode in {2, 3} else 2
        raise ReviewScriptError(_command_detail(validation), exit_code=code)
    return destination, (
        f"Policy `{config_path}` from trusted commit `{policy_commit[:12]}` "
        "was validated and applied."
    )


def _compare_change(
    repo: Path,
    base: str,
    head: str,
    change: GitChange,
    index: int,
    temp_root: Path,
    output_dir: Path,
    config_path: Path | None,
    fail_on: str,
) -> ReviewOutcome:
    if change.path_error:
        return ReviewOutcome("ERROR", change.display_path, 2, change.path_error)
    before_source = change.before_path
    after_source = change.after_path
    if before_source is None or after_source is None:
        raise ReviewScriptError("Internal comparison selection error", exit_code=3)

    before_suffix = PurePosixPath(before_source).suffix.lower()
    after_suffix = PurePosixPath(after_source).suffix.lower()
    before_file = temp_root / f"before-{index:03d}{before_suffix}"
    after_file = temp_root / f"after-{index:03d}{after_suffix}"
    _materialize_blob(repo, base, before_source, before_file)
    _materialize_blob(repo, head, after_source, after_file)

    prefix = _artifact_prefix(index, change)
    json_path = output_dir / f"{prefix}.json"
    html_path = output_dir / f"{prefix}.html"
    command = [
        "compare",
        str(before_file),
        str(after_file),
        "--json",
        str(json_path),
        "--html",
        str(html_path),
        "--fail-on",
        fail_on,
    ]
    if config_path is not None:
        command.extend(["--config", str(config_path)])
    completed = _run_sheetproof(command)
    exit_code = completed.returncode if completed.returncode in {0, 1, 2, 3} else 3

    if exit_code not in {0, 1} or not json_path.is_file() or not html_path.is_file():
        return ReviewOutcome(
            "ERROR",
            change.display_path,
            exit_code if exit_code in {2, 3} else 3,
            _command_detail(completed),
            json_report=json_path.name if json_path.is_file() else None,
            html_report=html_path.name if html_path.is_file() else None,
        )
    try:
        payload: dict[str, Any] = json.loads(json_path.read_text(encoding="utf-8"))
        summary = payload["summary"]
        risk_score = int(summary["risk_score"])
        risk_level = str(summary["risk_level"])
        changed_cells = int(summary["changed_cells"])
        formula_overwrites = int(summary["formula_overwrites"])
        failed_rules = int(summary["rules_failed"])
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        return ReviewOutcome(
            "ERROR",
            change.display_path,
            3,
            f"Generated JSON report was invalid: {exc}",
            json_report=json_path.name,
            html_report=html_path.name,
        )

    return ReviewOutcome(
        "PASS" if exit_code == 0 else "BLOCKED",
        change.display_path,
        exit_code,
        (completed.stdout or "Review completed").strip()[:1000],
        risk_level=risk_level,
        risk_score=risk_score,
        changed_cells=changed_cells,
        formula_overwrites=formula_overwrites,
        failed_rules=failed_rules,
        json_report=json_path.name,
        html_report=html_path.name,
    )


def _unpaired_outcome(
    change: GitChange,
    *,
    allow_unpaired: bool,
) -> ReviewOutcome | None:
    kind = change.status[0]
    if change.path_error:
        return ReviewOutcome("ERROR", change.display_path, 2, change.path_error)
    status = "SKIPPED" if allow_unpaired else "UNREVIEWABLE"
    exit_code = 0 if allow_unpaired else 1
    if kind == "A":
        return ReviewOutcome(
            status,
            change.display_path,
            exit_code,
            "Added workbook has no base-side version to compare; paired review is required.",
        )
    if kind == "D":
        return ReviewOutcome(
            status,
            change.display_path,
            exit_code,
            "Deleted workbook has no head-side version to compare; paired review is required.",
        )
    if kind == "R" and not (
        _is_workbook(change.before_path) and _is_workbook(change.after_path)
    ):
        return ReviewOutcome(
            status,
            change.display_path,
            exit_code,
            "Rename crosses the supported workbook boundary; paired review is required.",
        )
    if kind == "C":
        return ReviewOutcome(
            status,
            change.display_path,
            exit_code,
            "Copied workbook is unpaired and cannot receive a semantic comparison.",
        )
    if kind not in {"M", "R"}:
        return ReviewOutcome(
            status,
            change.display_path,
            exit_code,
            f"Git status {change.status} is not a reviewable workbook pair.",
        )
    return None


def _markdown(value: object) -> str:
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"([\\`*_\[\]{}()#+.!|~-])", r"\\\1", text)
    return html.escape(text, quote=True).replace("\n", "<br>")


def _summary_markdown(
    trusted_base: str,
    comparison_base: str,
    head: str,
    config_note: str,
    outcomes: Sequence[ReviewOutcome],
    output_name: str,
) -> str:
    compared = sum(outcome.status in {"PASS", "BLOCKED"} for outcome in outcomes)
    blocked = sum(outcome.status == "BLOCKED" for outcome in outcomes)
    unreviewable = sum(outcome.status == "UNREVIEWABLE" for outcome in outcomes)
    skipped = sum(outcome.status == "SKIPPED" for outcome in outcomes)
    errors = sum(outcome.status == "ERROR" for outcome in outcomes)
    lines = [
        "# SheetProof workbook review",
        "",
        f"Trusted base `{_markdown(trusted_base[:12])}` · "
        f"Merge base `{_markdown(comparison_base[:12])}` · "
        f"Head `{_markdown(head[:12])}` · "
        f"Compared **{compared}** · Blocked **{blocked}** · Skipped **{skipped}** · "
        f"Unreviewable **{unreviewable}** · Errors **{errors}**",
        "",
        config_note,
        "",
    ]
    if outcomes:
        lines.extend(
            [
                "| Result | Workbook | Risk | Changed cells | Formula overwrites | "
                "Failed rules | Reports / reason |",
                "| --- | --- | ---: | ---: | ---: | ---: | --- |",
            ]
        )
        for outcome in outcomes:
            risk = (
                f"{outcome.risk_score}/100 {outcome.risk_level}"
                if outcome.risk_score is not None
                else outcome.risk_level
            )
            reports = ", ".join(
                name for name in (outcome.json_report, outcome.html_report) if name is not None
            )
            detail = reports or outcome.message
            lines.append(
                "| "
                + " | ".join(
                    _markdown(value)
                    for value in (
                        outcome.status,
                        outcome.workbook,
                        risk,
                        outcome.changed_cells if outcome.changed_cells is not None else "—",
                        (
                            outcome.formula_overwrites
                            if outcome.formula_overwrites is not None
                            else "—"
                        ),
                        outcome.failed_rules if outcome.failed_rules is not None else "—",
                        detail,
                    )
                )
                + " |"
            )
    else:
        lines.append("No `.xlsx` or `.xlsm` changes were found between these commits.")
    lines.extend(
        [
            "",
            f"Artifacts are stored under `{_markdown(output_name)}`. Unpaired workbook "
            "changes block by default because a semantic comparison needs both sides.",
            "",
        ]
    )
    return "\n".join(lines)


def _aggregate_exit_code(outcomes: Sequence[ReviewOutcome]) -> int:
    codes = {outcome.exit_code for outcome in outcomes}
    if 3 in codes:
        return 3
    if 2 in codes:
        return 2
    if 1 in codes:
        return 1
    return 0


def _append_file_from_environment(variable: str, content: str) -> None:
    target = os.environ.get(variable)
    if not target:
        return
    try:
        with Path(target).open("a", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            if not content.endswith("\n"):
                stream.write("\n")
    except OSError as exc:
        print(f"Warning: unable to write {variable}: {exc}", file=sys.stderr)


def _write_action_outputs(
    output_dir: Path,
    repo: Path,
    outcomes: Sequence[ReviewOutcome],
    exit_code: int,
) -> None:
    compared = sum(outcome.status in {"PASS", "BLOCKED"} for outcome in outcomes)
    skipped = sum(outcome.status == "SKIPPED" for outcome in outcomes)
    content = (
        f"report-directory={output_dir.relative_to(repo).as_posix()}\n"
        f"compared-count={compared}\n"
        f"skipped-count={skipped}\n"
        f"review-exit-code={exit_code}\n"
    )
    _append_file_from_environment("GITHUB_OUTPUT", content)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Review modified/renamed Excel workbooks between two Git commits."
    )
    parser.add_argument(
        "--repository",
        default=".",
        help="Local Git checkout containing the base and head commits.",
    )
    parser.add_argument(
        "--base",
        required=True,
        help="Current trusted base commit or safe Git reference.",
    )
    parser.add_argument("--head", default="HEAD", help="Head commit SHA or safe Git reference.")
    parser.add_argument(
        "--config-ref",
        default="",
        help="Trusted commit containing the policy; defaults to --base.",
    )
    parser.add_argument(
        "--config",
        default="sheetproof.yml",
        help="Optional repository-relative policy path from --config-ref; empty disables it.",
    )
    parser.add_argument(
        "--output-dir",
        default="sheetproof-reports",
        help="New report directory inside the repository; it must not already exist.",
    )
    parser.add_argument(
        "--fail-on",
        choices=("LOW", "MEDIUM", "HIGH", "CRITICAL"),
        default="HIGH",
        help="Block at or above this risk level.",
    )
    parser.add_argument(
        "--allow-unpaired",
        action="store_true",
        help="Advisory mode: report unpaired workbook changes as skipped instead of blocking.",
    )
    return parser


def run(arguments: argparse.Namespace) -> int:
    repo = _repo_root(arguments.repository)
    output_dir = _safe_output_directory(repo, arguments.output_dir)
    trusted_base = _resolve_commit(repo, arguments.base, "base reference")
    head = _resolve_commit(repo, arguments.head, "head reference")
    comparison_base = _merge_base(repo, trusted_base, head)
    policy_reference = arguments.config_ref or arguments.base
    policy_commit = _resolve_commit(repo, policy_reference, "configuration reference")
    outcomes: list[ReviewOutcome] = []
    config_note = "Policy preparation did not complete."

    with tempfile.TemporaryDirectory(prefix="sheetproof-ci-") as temp_name:
        temp_root = Path(temp_name)
        try:
            config_path, config_note = _prepare_config(
                repo,
                policy_commit,
                arguments.config,
                temp_root,
            )
            changes = _changed_workbooks(repo, comparison_base, head)
            comparison_index = 0
            for change in changes:
                unpaired = _unpaired_outcome(
                    change,
                    allow_unpaired=arguments.allow_unpaired,
                )
                if unpaired is not None:
                    outcomes.append(unpaired)
                    continue
                comparison_index += 1
                try:
                    outcomes.append(
                        _compare_change(
                            repo,
                            comparison_base,
                            head,
                            change,
                            comparison_index,
                            temp_root,
                            output_dir,
                            config_path,
                            arguments.fail_on,
                        )
                    )
                except ReviewScriptError as exc:
                    outcomes.append(
                        ReviewOutcome(
                            "ERROR",
                            change.display_path,
                            exc.exit_code,
                            str(exc),
                        )
                    )
                except Exception as exc:  # keep a summary artifact for an isolated CI failure
                    outcomes.append(
                        ReviewOutcome(
                            "ERROR",
                            change.display_path,
                            3,
                            f"Unexpected comparison failure ({type(exc).__name__}): {exc}",
                        )
                    )
        except ReviewScriptError as exc:
            outcomes.append(ReviewOutcome("ERROR", "Policy / Git input", exc.exit_code, str(exc)))
        except Exception as exc:  # keep a summary artifact for an orchestration failure
            outcomes.append(
                ReviewOutcome(
                    "ERROR",
                    "Policy / Git input",
                    3,
                    f"Unexpected orchestration failure ({type(exc).__name__}): {exc}",
                )
            )

    exit_code = _aggregate_exit_code(outcomes)
    summary = _summary_markdown(
        trusted_base,
        comparison_base,
        head,
        config_note,
        outcomes,
        output_dir.relative_to(repo).as_posix(),
    )
    summary_path = output_dir / "summary.md"
    try:
        summary_path.write_text(summary, encoding="utf-8", newline="\n")
    except OSError as exc:
        raise ReviewScriptError(f"Unable to write aggregate summary: {exc}", exit_code=3) from exc
    _append_file_from_environment("GITHUB_STEP_SUMMARY", summary)
    _write_action_outputs(output_dir, repo, outcomes, exit_code)
    print(summary)
    print(f"SheetProof reports: {summary_path}")
    return exit_code


def main(argv: Sequence[str] | None = None) -> int:
    arguments = _build_parser().parse_args(argv)
    try:
        return run(arguments)
    except ReviewScriptError as exc:
        message = f"SheetProof CI orchestration failed: {exc}"
        print(message, file=sys.stderr)
        _append_file_from_environment(
            "GITHUB_STEP_SUMMARY",
            f"# SheetProof workbook review\n\n**ERROR:** {_markdown(message)}\n",
        )
        return exc.exit_code
    except Exception as exc:  # defensive boundary for a CI-facing script
        message = f"SheetProof CI orchestration failed internally ({type(exc).__name__}): {exc}"
        print(message, file=sys.stderr)
        _append_file_from_environment(
            "GITHUB_STEP_SUMMARY",
            f"# SheetProof workbook review\n\n**ERROR:** {_markdown(message)}\n",
        )
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
