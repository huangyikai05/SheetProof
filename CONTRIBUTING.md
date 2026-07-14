# Contributing to SheetProof

Thank you for helping make spreadsheet review more reliable. SheetProof follows one
non-negotiable design rule: deterministic code produces every finding and gate decision;
human- or AI-written prose may explain evidence, but may not replace it.

## Before opening a change

- Search existing issues and pull requests.
- Use a focused issue for behavior changes or new rules. Security reports belong in the
  private process described in [SECURITY.md](SECURITY.md), not a public issue.
- Keep the MVP boundaries in mind: local `.xlsx`/`.xlsm` inspection, deterministic
  comparison, rules, scoring, reports, CLI, web demo, and CI integration.

## Development setup

SheetProof requires Python 3.11 or newer.

```bash
python -m venv .venv
python -m pip install --upgrade pip
python -m pip install -e ".[dev,web]"
```

Activate the virtual environment before running commands. On POSIX shells use
`source .venv/bin/activate`; on PowerShell use `.\.venv\Scripts\Activate.ps1`.

## Required checks

Run the complete local verification before requesting review:

```bash
pytest
ruff check .
mypy sheetproof
python examples/generate_demo_workbooks.py
sheetproof compare examples/generated/before.xlsx examples/generated/after_safe.xlsx \
  --config examples/sheetproof.example.yml \
  --json build/safe.json \
  --html build/safe.html
```

Tests must generate their own small workbooks. Do not commit workbooks copied from a company,
customer, or other sensitive source.

## Code and test expectations

- Add type annotations and keep modules narrowly responsible.
- Put deterministic review logic in the core package, not in the CLI, Streamlit page, or
  report template.
- Preserve raw evidence and report unsupported analysis explicitly. Never invent a formula
  result when Excel has not stored a usable cached value.
- Any change to parsing, diff classification, formula analysis, built-in rules, exit codes,
  or risk weights requires focused unit tests and, where appropriate, an integration test.
- Do not weaken or delete a failing test to make a change pass.
- Keep HTML output offline and autoescaped.

## Pull requests

Create a branch with a descriptive name such as `feat/formula-rule` or `fix/xlsm-detection`.
Small conventional commits are encouraged, for example `feat: detect table range changes`.

A pull request should explain:

- the problem and the deterministic behavior added or changed;
- important design and security decisions;
- files and public interfaces affected;
- commands run and their real results;
- known limitations or follow-up work.

The pull request template contains the complete checklist. A maintainer may ask for smaller
scope when unrelated changes are bundled together.

## Documentation

Update README examples when CLI, configuration, action inputs, output fields, or limitations
change. Do not describe planned capabilities as available. Add user-visible changes under
`Unreleased` in [CHANGELOG.md](CHANGELOG.md).

By contributing, you agree that your contribution is licensed under the MIT License.
