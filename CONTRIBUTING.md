# Contributing to Tabulint

Thank you for helping improve deterministic spreadsheet review. Tabulint has one
non-negotiable design rule:

> Programs validate. AI may explain evidence, but AI text, confidence, or model responses must
> never become validation evidence or a merge-gate decision.

## Before opening a change

- Search existing issues and pull requests.
- Keep changes focused on local workbook inspection, comparison, deterministic rules, risk
  scoring, reports, interfaces, or CI integration.
- Discuss large behavior or public-contract changes in an issue first.
- Report vulnerabilities through the private process in [SECURITY.md](SECURITY.md), never a
  public issue.
- Never upload or commit a real workbook containing personal, financial, customer, employer, or
  other confidential data. Build the smallest synthetic reproduction instead.

## Clone and install

Tabulint requires Python 3.11 or newer.

On macOS or Linux:

~~~bash
git clone https://github.com/huangyikai05/Tabulint.git
cd Tabulint
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
~~~

On PowerShell:

~~~powershell
git clone https://github.com/huangyikai05/Tabulint.git
Set-Location Tabulint
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
~~~

On Command Prompt:

~~~bat
git clone https://github.com/huangyikai05/Tabulint.git
cd Tabulint
py -3.11 -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
~~~

If PowerShell execution policy prevents activation, use Command Prompt or invoke
<code>.\.venv\Scripts\python.exe</code> directly. The important detail is that subsequent
installation and test commands use the new virtual environment, not the system interpreter.

Install the optional web dependencies only when needed:

~~~bash
python -m pip install -e ".[dev,web]"
~~~

## Create a branch

Start from an up-to-date default branch and use a focused name:

~~~bash
git switch main
git pull --ff-only
git switch -c feat/short-description
~~~

Examples include <code>fix/xlsm-detection</code>, <code>feat/formula-rule</code>, and
<code>docs/windows-installation</code>. Do not combine unrelated refactors and behavior changes.

## Run the required checks

Run focused tests while iterating, then the complete project checks before requesting review:

~~~bash
pytest
ruff check .
mypy tabulint
python -m compileall -q tabulint web scripts examples
~~~

For a user-facing comparison or report change, also regenerate and run both synthetic demos:

~~~bash
python examples/generate_demo_workbooks.py

tabulint compare examples/generated/before.xlsx examples/generated/after_safe.xlsx \
  --config examples/tabulint.example.yml \
  --json build/safe-review.json \
  --html build/safe-review.html

tabulint compare examples/generated/before.xlsx examples/generated/after_risky.xlsx \
  --config examples/tabulint.example.yml \
  --json build/risky-review.json \
  --html build/risky-review.html
~~~

The risky command is expected to return exit code <code>1</code>. Do not mask an unexpected
<code>2</code> input error or <code>3</code> internal error.

## Add tests and workbook fixtures

Tests must create small workbooks dynamically with openpyxl, normally through helpers in
<code>tests/conftest.py</code> or a narrowly scoped helper in the test module.

- Keep every fixture synthetic and minimal.
- Add only the workbook feature needed to exercise the behavior.
- Assert typed evidence: change type, location, before/after facts, reason, limit state, and exit
  code as relevant.
- Cover the successful case, invalid or unsupported input, and important boundaries.
- Include Unicode sheet/file names, paths with spaces, and platform path behavior when relevant.
- Do not commit customer workbooks, downloaded mystery files, or sanitized-looking files whose
  provenance is unclear.

A change to a parser fact, diff classification, formula analysis, dependency bound, risk weight,
report contract, rule behavior, or CLI exit code must add or update focused tests. Never weaken,
skip, or delete a test merely to make a change pass.

## Change the rule engine

Built-in rules live in <code>tabulint/rules/</code> and their canonical configuration contracts
live in the typed models.

When adding or changing a rule:

1. Define strict configuration fields and reject misspellings.
2. Implement deterministic evaluation from parsed review evidence.
3. Return a typed status, severity, reason, location, and supporting evidence.
4. Distinguish <code>SKIPPED</code> limitations from evaluator <code>ERROR</code> failures.
5. Add loader, evaluator, edge-case, and integration tests.
6. Update the README policy example and changelog when the behavior is user-visible.

Do not call a model, network service, spreadsheet application, macro, or workbook-controlled
command to decide a rule result.

## Code expectations

- Target Python 3.11+ and add type annotations.
- Keep functions single-purpose and deterministic ordering stable.
- Put business decisions in the core service path, not in CLI, Streamlit, templates, or CI
  wrappers.
- Preserve unsupported inputs and limitations in the result; never silently drop an analysis
  failure.
- Never invent cached formula values or claim to emulate Excel calculation.
- Keep dependency traversal and parser resource use bounded.
- Use safe YAML loading, path containment, argument-list subprocess calls, and temporary
  directories for untrusted input.
- Keep HTML offline and autoescaped.

Read [AGENTS.md](AGENTS.md) for the complete architecture, security boundaries, and prohibited
changes.

## Documentation

Update documentation whenever CLI flags, configuration, public Python APIs, Action inputs,
report fields, exit codes, security boundaries, or limitations change. Commands and output
counts must be actually reproducible. Do not advertise planned capabilities, a PyPI publication,
or a Trusted Publisher configuration before it is confirmed.

Record user-visible changes under <code>Unreleased</code> in [CHANGELOG.md](CHANGELOG.md).

## Commits

Small conventional commits are encouraged:

~~~text
feat: detect table range changes
fix: preserve unsupported formula evidence
test: cover Unicode workbook paths
docs: clarify Action artifact privacy
~~~

Write commit messages about the behavior changed, not the tool used to edit it.

## Open a pull request

A pull request should explain:

- the problem and deterministic behavior added or changed;
- public interfaces and files affected;
- security and trust-boundary decisions;
- tests added;
- exact commands run and their real results;
- known limitations and follow-up work.

Complete the repository pull-request template. Maintainers may ask to split unrelated work or add
boundary coverage. A new feature is incomplete without its tests and corresponding documentation.

By contributing, you agree that your contribution is licensed under the [MIT License](LICENSE).
