"""Shared factories for small, source-generated workbook fixtures."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

WorkbookConfigurator = Callable[[Workbook], None]
WorkbookFactory = Callable[[str, WorkbookConfigurator | None], Path]


@pytest.fixture
def workbook_factory(tmp_path: Path) -> WorkbookFactory:
    """Create a workbook under pytest's temporary directory.

    Tests intentionally generate every OOXML input at runtime.  No opaque binary
    fixtures are checked into the repository.
    """

    def create(name: str, configure: WorkbookConfigurator | None = None) -> Path:
        path = tmp_path / name
        workbook = Workbook()
        if configure is not None:
            configure(workbook)
        workbook.save(path)
        workbook.close()
        return path

    return create


def add_vba_project(path: Path, payload: bytes = b"SheetProof test VBA marker") -> None:
    """Add an inert VBA package member for deterministic detection tests.

    SheetProof detects the presence of ``xl/vbaProject.bin`` and never executes
    its bytes.  The payload is deliberately not executable VBA.
    """

    with ZipFile(path, mode="a", compression=ZIP_DEFLATED) as archive:
        archive.writestr("xl/vbaProject.bin", payload)
