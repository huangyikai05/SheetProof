"""Build metadata and artifact tests that do not rely on checked-in ``dist`` files."""

from __future__ import annotations

import shutil
import subprocess
import sys
import tarfile
import tomllib
from pathlib import Path
from zipfile import ZipFile

import pytest

from sheetproof import __version__

PROJECT_ROOT = Path(__file__).parents[1]


def test_pyproject_uses_canonical_dynamic_version_and_release_metadata() -> None:
    pyproject = tomllib.loads((PROJECT_ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    project = pyproject["project"]
    setuptools = pyproject["tool"]["setuptools"]
    assert "version" not in project
    assert project["dynamic"] == ["version"]
    assert setuptools["dynamic"]["version"]["attr"] == "sheetproof._version.__version__"
    assert project["license"] == "MIT"
    assert project["license-files"] == ["LICENSE"]
    assert project["urls"]["Changelog"].endswith("/CHANGELOG.md")
    runtime_dependencies = project["dependencies"]
    assert "click>=8.3.3,<9" in runtime_dependencies
    assert "Pygments>=2.20,<3" in runtime_dependencies
    assert "typer>=0.26.8,<1" in runtime_dependencies
    assert {item.split(">=", 1)[0] for item in project["optional-dependencies"]["release"]} == {
        "build",
        "setuptools",
        "twine",
    }
    assert setuptools["packages"]["find"]["namespaces"] is False
    assert setuptools["include-package-data"] is False
    assert "py.typed" in setuptools["package-data"]["sheetproof"]
    assert "templates/*.j2" in setuptools["package-data"]["sheetproof.reports"]


def test_fresh_build_contains_version_license_template_and_type_marker(tmp_path: Path) -> None:
    pytest.importorskip("build", reason="artifact verification requires the release extra")
    source = tmp_path / "source"
    artifacts = tmp_path / "artifacts"
    source.mkdir()
    artifacts.mkdir()

    for filename in (
        "AGENTS.md",
        "CHANGELOG.md",
        "CONTRIBUTING.md",
        "LICENSE",
        "MANIFEST.in",
        "README.md",
        "ROADMAP.md",
        "SECURITY.md",
        "pyproject.toml",
    ):
        shutil.copy2(PROJECT_ROOT / filename, source / filename)
    source_directories = (
        ".streamlit",
        "action",
        "docs",
        "examples",
        "scripts",
        "sheetproof",
        "tests",
        "web",
    )
    for directory in source_directories:
        shutil.copytree(
            PROJECT_ROOT / directory,
            source / directory,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "generated"),
        )

    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "build",
            "--no-isolation",
            "--outdir",
            str(artifacts),
        ],
        cwd=source,
        capture_output=True,
        check=False,
        text=True,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr

    wheels = list(artifacts.glob("*.whl"))
    sdists = list(artifacts.glob("*.tar.gz"))
    assert len(wheels) == 1
    assert len(sdists) == 1

    with ZipFile(wheels[0]) as wheel:
        wheel_names = set(wheel.namelist())
        metadata_name = next(name for name in wheel_names if name.endswith(".dist-info/METADATA"))
        metadata = wheel.read(metadata_name).decode("utf-8")
        assert f"Version: {__version__}" in metadata
        assert "License-Expression: MIT" in metadata
        assert "Project-URL: Changelog," in metadata
        assert "sheetproof/py.typed" in wheel_names
        assert "sheetproof/reports/templates/report.html.j2" in wheel_names
        assert any(name.endswith(".dist-info/licenses/LICENSE") for name in wheel_names)

    with tarfile.open(sdists[0], mode="r:gz") as sdist:
        sdist_names = set(sdist.getnames())
        assert any(name.endswith("/CHANGELOG.md") for name in sdist_names)
        assert any(name.endswith("/ROADMAP.md") for name in sdist_names)
        assert any(name.endswith("/docs/demo-safe.md") for name in sdist_names)
        assert any(
            name.endswith("/examples/generate_demo_workbooks.py") for name in sdist_names
        )
        assert any(name.endswith("/web/app.py") for name in sdist_names)
        assert any(name.endswith("/action/action.yml") for name in sdist_names)
        assert any(
            name.endswith("/scripts/review_changed_workbooks.py") for name in sdist_names
        )
        assert any(name.endswith("/sheetproof/py.typed") for name in sdist_names)
        assert any(
            name.endswith("/sheetproof/reports/templates/report.html.j2")
            for name in sdist_names
        )
