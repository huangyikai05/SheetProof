# Tabulint

[![CI](https://github.com/huangyikai05/Tabulint/actions/workflows/ci.yml/badge.svg)](https://github.com/huangyikai05/Tabulint/actions/workflows/ci.yml)
[![Tested on Python 3.11 and 3.12](https://img.shields.io/badge/tested-Python%203.11%20%7C%203.12-3776AB?logo=python&logoColor=white)](https://github.com/huangyikai05/Tabulint/blob/main/pyproject.toml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/huangyikai05/Tabulint/blob/main/LICENSE)

**Deterministic spreadsheet change review and CI for Excel workbooks.**

Tabulint is an open-source spreadsheet review and CI tool that detects risky Excel changes,
formula overwrites, structural modifications, dependency impacts, and business-rule violations
before they reach production.

Tabulint 是一个开源的 Excel 变更审查与持续集成工具，用于在文件投入使用前发现公式覆盖、
结构变化、依赖影响和业务规则违规。

Install the current public source and compare two workbooks:

~~~bash
python -m pip install "tabulint @ git+https://github.com/huangyikai05/Tabulint.git@main"
tabulint compare before.xlsx after.xlsx --html report.html
~~~

The generated risky demo produces this real result:

~~~text
Tabulint Review

Risk score:             100/100
Risk level:             CRITICAL
Changed cells:          5
Changed formulas:       3
Formula overwrites:     1
Hidden sheets added:    1
External links added:   1
Rules:                  4 failed, 1 warning, 1 skipped
Exit code:              1
~~~

The exact CLI summary is:

~~~text
Risk 100/100 (CRITICAL); 5 changed cells; 1 formula overwrites; 4 failed, 1 warning, and 0 errored rules.
~~~

[Run the safe demo](https://github.com/huangyikai05/Tabulint/blob/main/docs/demo-safe.md) ·
[Run the risky demo](https://github.com/huangyikai05/Tabulint/blob/main/docs/demo-risky.md) ·
[Record the 30–60 second demo](https://github.com/huangyikai05/Tabulint/blob/main/docs/demo-script.md) ·
[Read the v0.1.0 release notes](https://github.com/huangyikai05/Tabulint/blob/main/docs/releases/v0.1.0.md)

> Install published releases from PyPI. To test unreleased changes, install the public
> <code>main</code> branch or clone the repository as described below.

## How it works

~~~mermaid
flowchart LR
    B["before.xlsx"] --> S["Tabulint"]
    A["after.xlsx"] --> S
    P["optional tabulint.yml"] --> S
    S --> D["Semantic diff"]
    S --> F["Formula overwrite and pattern detection"]
    S --> G["Bounded dependency impact"]
    S --> R["Deterministic business rules"]
    D --> E["Typed review evidence"]
    F --> E
    G --> E
    R --> E
    E --> J["JSON"]
    E --> H["Offline HTML"]
    E --> C["CLI / CI exit code"]
~~~

Programs produce every finding, risk contribution, rule result, and exit code. AI-generated
text is never used as validation evidence or as a merge decision.

## Why Tabulint

An Excel workbook is a ZIP-based document, not a line-oriented source file. A normal Git diff
cannot explain that a formula became a fixed value, a total range became shorter, a hidden sheet
appeared, or a critical downstream cell may be affected.

Tabulint converts workbook facts into stable, reviewable evidence:

- semantic cell and workbook changes instead of raw XML noise;
- explicit formula-overwrite, range-reduction, and copied-pattern findings;
- bounded upstream and downstream dependency evidence;
- project-specific YAML rules;
- explainable, capped risk contributions;
- one typed result shared by the CLI, reports, web page, Python API, and CI.

Tabulint is a review aid. It does not prove that a workbook is correct and does not replace
professional financial, security, legal, or operational review.

## Key features

- **Bounded non-executing parser:** reads <code>.xlsx</code> and <code>.xlsm</code> facts while
  enforcing archive, cell, merge, formula-expansion, and graph limits.
- **Semantic comparison:** classifies values, text, blanks, formulas, types, styles, names,
  tables, validations, merges, hidden rows/columns, sheet visibility, VBA presence, and external
  link indicators.
- **Formula review:** detects formula replacement, range reduction, additions, and broken
  row/column copy patterns without calculating formulas.
- **Dependency impact:** reports direct and bounded downstream relationships, paths, cycles,
  critical-cell impact, and truncation state.
- **Strict YAML policy:** validates nine built-in deterministic rule types and rejects duplicate
  keys, unknown fields, and invalid values.
- **Explainable risk:** deduplicates named risk contributions and caps the total at 100.
- **Portable reports:** writes structured JSON and self-contained, autoescaped offline HTML.
- **Multiple adapters:** Typer CLI, source-checkout Streamlit page, public Python API, and a
  read-only GitHub Action use the same review service.

## Quick start

Clone the repository to generate both reproducible demo cases:

~~~bash
git clone https://github.com/huangyikai05/Tabulint.git
cd Tabulint
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e .
python examples/generate_demo_workbooks.py

tabulint compare examples/generated/before.xlsx examples/generated/after_safe.xlsx \
  --config examples/tabulint.example.yml \
  --json build/safe-review.json \
  --html build/safe-review.html
~~~

The safe comparison exits <code>0</code> with <code>2/100 LOW</code>. The risky comparison below
intentionally exits <code>1</code>; that exit is the expected gate result, not a crash:

~~~bash
tabulint compare examples/generated/before.xlsx examples/generated/after_risky.xlsx \
  --config examples/tabulint.example.yml \
  --json build/risky-review.json \
  --html build/risky-review.html
~~~

PowerShell setup:

~~~powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
~~~

Command Prompt setup:

~~~bat
py -3.11 -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -e .
~~~

If PowerShell blocks script activation, either use Command Prompt or run the virtual
environment interpreter directly, for example
<code>.\.venv\Scripts\python.exe -m pip install -e .</code>.

## Installation

Requirements:

- Python 3.11 or newer (release CI currently tests Python 3.11 and 3.12);
- a platform supported by Python and openpyxl;
- local access to the workbooks being reviewed.

Install a published release from PyPI:

~~~bash
python -m pip install tabulint
~~~

To test unreleased changes from the public repository:

~~~bash
python -m pip install "tabulint @ git+https://github.com/huangyikai05/Tabulint.git@main"
~~~

From a clone:

~~~bash
python -m pip install .
~~~

Optional local web interface:

~~~bash
python -m pip install ".[web]"
~~~

Contributor environment:

~~~bash
python -m pip install -e ".[dev,web]"
~~~

Verify the installed entry point with <code>tabulint version</code>,
<code>tabulint --help</code>, and <code>tabulint compare --help</code>.

The canonical package installation command is <code>python -m pip install tabulint</code>.

## CLI usage

### Compare two workbooks

~~~bash
tabulint compare BEFORE.xlsx AFTER.xlsx \
  [--config tabulint.yml] \
  [--json report.json] \
  [--html report.html] \
  [--fail-on LOW|MEDIUM|HIGH|CRITICAL]
~~~

The configuration controls the default blocking level. <code>--fail-on</code> overrides it.
A failed or errored rule always blocks.

### Inspect one workbook

~~~bash
tabulint inspect workbook.xlsx
tabulint inspect workbook.xlsm --json build/snapshot.json
~~~

Inspection records workbook facts. It does not recalculate formulas or run VBA.

### Validate a policy

~~~bash
tabulint rules validate tabulint.yml
tabulint version
~~~

### Exit codes

| Code | Meaning |
| ---: | --- |
| 0 | Review completed without a configured blocking condition. |
| 1 | Risk or a finding reached the blocking threshold, or a rule failed/errored. Reports are still written. |
| 2 | Workbook, path, output, or policy input was invalid. |
| 3 | An unexpected internal error occurred. Use <code>--debug</code> locally for a traceback. |

## Web interface

Install the web extra and start Streamlit from the repository root:

~~~bash
python -m pip install -e ".[web]"
streamlit run web/app.py
~~~

The page accepts before/after workbooks and an optional policy, then provides summary tables and
JSON/HTML downloads. Each upload is limited to 25 MiB. Files are copied to a private temporary
directory for the review and removed afterward. The web page contains no alternate verdict logic.

## YAML rules

Configuration is optional. It is loaded with a duplicate-key-rejecting safe YAML loader and
validated by strict Pydantic models.

~~~yaml
rules:
  - name: Cash-flow formulas must remain formulas
    type: formula_required
    range: "现金流!B5:B20"
    severity: high

  - name: Forecast edits stay in the approved area
    type: allowed_change_range
    ranges:
      - "预测!D5:H30"
    severity: high

  - name: Do not add external links
    type: no_external_links
    severity: critical

  - name: Do not add hidden sheets
    type: no_new_hidden_sheets
    severity: high

critical_cells:
  - "现金流!B22"

block_risk_level: HIGH
max_dependency_depth: 10
max_dependency_nodes: 10000
~~~

A numeric rule that targets a formula without a usable cached value is reported as
<code>SKIPPED</code>; Tabulint never invents the value or pretends to emulate Excel.
See [the complete demo policy](https://github.com/huangyikai05/Tabulint/blob/main/examples/tabulint.example.yml).

## GitHub Actions

The repository provides:

- a normal <code>CI</code> workflow for tests, Ruff, Mypy, and compilation on supported Python
  versions;
- a composite Tabulint action and a read-only workbook pull-request gate;
- a release workflow using PyPI Trusted Publishing.

The required workbook gate deliberately loads its implementation, helper, and policy from a
separate trusted base checkout. Pull-request workbooks are treated as Git-object input; PR code
is not imported or executed. Added or deleted workbooks are <code>UNREVIEWABLE</code> and block
by default because a semantic comparison needs both sides. The action supports an explicit
advisory mode for unpaired files.

The PyPI release workflow requires a protected <code>pypi</code> GitHub Environment and a
matching PyPI Trusted Publisher. It exchanges GitHub's short-lived OIDC identity for upload
authority, so no PyPI token belongs in the repository.

See [the composite action](https://github.com/huangyikai05/Tabulint/blob/main/action/action.yml)
and [the workbook workflow](https://github.com/huangyikai05/Tabulint/blob/main/.github/workflows/tabulint.yml).

### Action artifact privacy

Tabulint itself does not upload workbooks or telemetry. A GitHub Actions workflow may,
however, upload generated JSON and HTML reports as artifacts. Those reports can contain derived
workbook content such as cell addresses, formulas, before/after values, sheet names, external-link
names, and rule evidence. Repository owners should treat report artifacts according to the source
workbook's sensitivity, restrict repository and artifact access, choose an appropriate retention
period, and avoid public CI for confidential workbooks.

## Example reports

The repository stores generators rather than proprietary workbook fixtures. Run
<code>python examples/generate_demo_workbooks.py</code> to create:

| Case | Risk | Evidence summary | Exit |
| --- | --- | --- | ---: |
| [Safe update](https://github.com/huangyikai05/Tabulint/blob/main/docs/demo-safe.md) | 2/100 LOW | 1 changed cell; no formula, structure, link, or blocking-rule finding | 0 |
| [Risky update](https://github.com/huangyikai05/Tabulint/blob/main/docs/demo-risky.md) | 100/100 CRITICAL | 5 changed cells; 3 formula changes; 1 overwrite; 1 hidden sheet; 1 external link; 4 failed rules | 1 |

The commands write machine-readable JSON and a self-contained HTML report to <code>build/</code>.
Generated reports are intentionally not presented as static golden verdicts; regenerate them
from the synthetic source script and current deterministic implementation.

## Python API

The supported top-level helper returns the same typed <code>ReviewResult</code> used by every
interface:

~~~python
from tabulint import compare_workbooks

result = compare_workbooks(
    before_path="before.xlsx",
    after_path="after.xlsx",
    config_path="tabulint.yml",
)

print(result.summary.risk_score)
print(result.summary.risk_level.value)
for finding in result.formula_changes:
    print(finding.location, finding.change_type)
~~~

The helper performs deterministic review only. Callers should use typed fields and enum values,
not parse human-readable descriptions to implement a gate.

## Security model

- VBA presence is detected; VBA is never executed, decompiled, verified, or removed.
- Formulas are read as text; they are never executed or recalculated.
- External workbook links are recorded as evidence; they are never opened or fetched.
- Workbook-controlled shell text, embedded objects, and commands are never executed.
- YAML uses a safe loader with strict validation.
- OOXML paths and resource budgets are validated before materialization.
- HTML is autoescaped and contains no online CDN dependency.
- Core findings and gate decisions come from deterministic program logic, never a language model.

Read [SECURITY.md](https://github.com/huangyikai05/Tabulint/blob/main/SECURITY.md) before
processing untrusted inputs and use its private reporting process for vulnerabilities.

## Privacy

Local CLI, Python, and Streamlit reviews require no account, API key, database, or cloud service.
Tabulint does not upload workbook contents or emit telemetry. Files remain on the machine unless
the user or an integration explicitly copies or uploads an output. In particular, review the
[Action artifact privacy](#action-artifact-privacy) boundary before enabling CI for sensitive
workbooks.

## Known limitations

- Tabulint cannot fully simulate Excel's calculation engine. Cached values may be absent or
  stale, and Tabulint never fabricates them.
- Formula parsing is not a complete Excel grammar. Some complex formulas receive reference-level
  analysis only; dynamic references such as <code>INDIRECT</code> cannot be reliably resolved.
- VBA is detected only; macro behavior is not analyzed.
- Charts, slicers, Power Query, data models, embedded objects, digital signatures,
  conditional-formatting semantics, and every style detail are not fully compared.
- ZIP64 and multi-disk OOXML packages are outside the current security profile.
- External-link detection does not guarantee that every connection type is found.
- Added/deleted workbook pairs are unreviewable by semantic diff and fail closed in the default
  CI gate.
- A low risk score does not prove correctness, and a high score does not prove malicious intent.
  The score prioritizes human review; it is not a probability or professional audit conclusion.

## Roadmap

See [ROADMAP.md](https://github.com/huangyikai05/Tabulint/blob/main/ROADMAP.md) for the public,
non-date-bound roadmap. Proposed post-MVP work includes formula graph visualization, better
formula parsing, optional PR comments, deterministic rule plugins, performance work, and
carefully bounded evidence-explanation integrations. Roadmap items are proposals, not release
promises.

## Contributing

Read [CONTRIBUTING.md](https://github.com/huangyikai05/Tabulint/blob/main/CONTRIBUTING.md) and
[AGENTS.md](https://github.com/huangyikai05/Tabulint/blob/main/AGENTS.md). Core changes must
include focused tests and preserve the rule that programs validate while AI may only explain
evidence. Use synthetic minimal workbooks; never submit confidential, proprietary, or
source-unknown Excel files.

## License

Tabulint is available under the
[MIT License](https://github.com/huangyikai05/Tabulint/blob/main/LICENSE).
