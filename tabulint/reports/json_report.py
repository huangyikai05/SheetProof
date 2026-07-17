"""Serialize the canonical review result as human-readable JSON."""

from __future__ import annotations

import json
from pathlib import Path

from tabulint.exceptions import ReportGenerationError
from tabulint.models import ReviewResult


def render_json(result: ReviewResult) -> str:
    """Return a UTF-8 friendly, deterministic JSON representation of ``result``."""

    try:
        payload = result.model_dump(mode="json")
        return json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    except (TypeError, ValueError) as exc:
        raise ReportGenerationError(f"Unable to render JSON report: {exc}") from exc


def write_json_report(result: ReviewResult, path: str | Path) -> Path:
    """Write a JSON report to ``path`` and return its resolved output path."""

    target = Path(path).expanduser()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(render_json(result), encoding="utf-8", newline="\n")
    except OSError as exc:
        raise ReportGenerationError(f"Unable to write JSON report to {target}: {exc}") from exc
    return target.resolve()
