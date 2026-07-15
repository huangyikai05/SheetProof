"""Stable public API for deterministic workbook reviews."""

from __future__ import annotations

from pathlib import Path

from sheetproof._version import __version__
from sheetproof.models import ReviewResult

__all__ = ["ReviewResult", "__version__", "compare_workbooks"]


def compare_workbooks(
    before_path: str | Path,
    after_path: str | Path,
    config_path: str | Path | None = None,
) -> ReviewResult:
    """Compare two workbooks through SheetProof's canonical review service.

    The service import is intentionally local so internal modules can read the
    package version without creating a package-initialization cycle.
    """

    from sheetproof.services.review_service import ReviewService

    return ReviewService().review(before_path, after_path, config_path)
