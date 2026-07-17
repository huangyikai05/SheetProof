# Tabulint agent guide

This file defines the repository constraints for human contributors and coding agents. The
central rule is: **programs validate; AI may explain evidence but must never decide whether a
workbook is correct.**

## Architecture

- `tabulint/parser/`: bounded, non-executing workbook and formula fact extraction.
- `tabulint/diff/`: structure, semantic cell, formula, and copied-pattern differences.
- `tabulint/graph/`: bounded dependency graph construction and impact traversal.
- `tabulint/rules/`: safe YAML loading and deterministic built-in rules.
- `tabulint/risk/`: configurable, deduplicated, capped risk contributions.
- `tabulint/reports/`: JSON and autoescaped, offline HTML generated only from `ReviewResult`.
- `tabulint/services/`: the single orchestration path shared by interfaces.
- `tabulint/cli.py`: Typer adapter and documented exit-code contract.
- `web/app.py`: local Streamlit adapter; no analysis logic belongs here.
- `action/` and `scripts/`: GitHub Action adapter and base/head workbook orchestration.
- `examples/`: synthetic rules and reproducible demo-workbook generator.
- `tests/`: dynamically generated fixtures plus unit and integration coverage.

Dependency direction flows toward the typed models and core services. CLI, web, reports, and
CI must not maintain alternate verdict logic.

## Development commands

```bash
python -m venv .venv
python -m pip install -e ".[dev,web]"
python examples/generate_demo_workbooks.py
tabulint version
tabulint compare examples/generated/before.xlsx examples/generated/after_safe.xlsx \
  --config examples/tabulint.example.yml --json build/report.json --html build/report.html
streamlit run web/app.py
```

## Required verification

```bash
pytest
ruff check .
mypy tabulint
python -m compileall -q tabulint web scripts examples
```

Run focused tests while iterating, then the full suite. A change to a core rule, parser fact,
diff classification, dependency limit, risk weight, report contract, or CLI exit code must add
or update tests that cover success, failure, and relevant edge cases.

## Code conventions

- Target Python 3.11+, use type annotations, and keep functions single-purpose.
- Keep the canonical contracts in Pydantic models and reject misspelled configuration fields.
- Prefer deterministic ordering and stable evidence over heuristic confidence scores.
- Preserve unsupported inputs in results and identify the limitation; do not silently ignore
  parsing or analysis failures.
- Use `pathlib`, standard temporary directories, argument-list subprocess calls, and explicit
  path containment checks for untrusted paths.
- Keep examples synthetic, reports offline, and user-facing claims synchronized with tests.

## Security boundaries

- Never execute VBA, formulas, embedded objects, shell text, or workbook-controlled commands.
- Never fetch external workbook links or upload workbook contents.
- Never deserialize YAML with an unsafe loader.
- Never render workbook content without contextual escaping.
- Never remove parser resource limits or bounded graph traversal without a documented threat
  analysis and replacement limits.
- GitHub workflows use read-only permissions and must not expose secrets to pull-request code.
  Required PR gates must execute the Action, helper, package, and policy from a separate trusted
  base checkout; never use `./action`, scripts, or policy files from the PR-head checkout.

## Prohibited changes

- Do not use random numbers, hidden mock verdicts, or hard-coded demo results.
- Do not fabricate cached formula values or claim to emulate Excel's calculation engine.
- Do not use AI text, model confidence, or a language-model response as validation evidence or
  a merge gate.
- Do not duplicate business decisions in the CLI, Streamlit page, templates, or CI wrapper.
- Do not weaken, skip, or delete tests merely to make a change pass.
- Do not add accounts, billing, cloud storage, databases, agents, or unrelated infrastructure
  to the MVP.
- Do not commit proprietary, sensitive, or source-unknown workbooks.
