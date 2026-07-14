# SheetProof

**Deterministic semantic review and CI for Excel workbook changes.**

SheetProof compares a before and after `.xlsx` or `.xlsm` workbook, records evidence in one
typed result, assigns an explainable risk score, evaluates optional business rules, and writes
JSON and offline HTML review reports. It works without an API key and does not ask an AI model
to decide whether a change is safe.

> Programs validate. AI may explain.

SheetProof is an alpha MVP. Read [Known limitations](#13-known-limitations) before using it for
important decisions.

## 1. Project overview

Spreadsheet files are ZIP-based documents rather than line-oriented source files. A normal Git
diff cannot explain that a formula became a fixed value, a hidden sheet appeared, or a named
range moved. SheetProof turns those workbook facts into reviewable, machine-readable evidence
and a CI-compatible exit code.

The current release is entirely local: one review takes two files and an optional YAML policy.
The CLI, HTML report, Streamlit page, and GitHub Action all call the same deterministic review
service and consume the same Pydantic result.

## 2. Why SheetProof

Human and automated edits can both introduce subtle spreadsheet errors. Common examples include
overwriting one formula in a copied column, shortening a total range, adding an external link,
or changing content outside an approved forecast area. These changes may still leave a workbook
that opens normally.

SheetProof provides evidence for pull-request review:

- semantic changes instead of raw XML noise;
- explicit formula-overwrite and copied-pattern findings;
- bounded downstream dependency impact;
- project-specific policy checks;
- a reproducible score with named point contributions;
- local reports and merge gating without a paid service.

It is a review aid, not a guarantee that a workbook is financially, legally, or operationally
correct.

## 3. Core features

- **Bounded workbook inspection:** sheet names/order/visibility, cells, raw values, formulas,
  cached values, types, hidden rows and columns, merged cells, names, data validation, freeze
  panes, table ranges, external-link indicators, VBA presence, and key style summaries.
- **Structural diff:** added/deleted/reordered or newly hidden sheets, visibility, hidden rows
  and columns, merges, names, validations, freeze panes, tables, external links, and VBA
  presence changes.
- **Semantic cell diff:** distinguishes numbers, text, dates, errors, blanks, formulas, type
  changes, clearing, formula addition, and formula replacement.
- **Formula analysis:** extracts A1 references for common functions, identifies reduced formula
  ranges, retains unsupported formulas with `unsupported_formula_analysis`, and never evaluates
  formula results.
- **Formula-pattern review:** flags a fixed value or blank that breaks a nearby row/column copy
  pattern and retains the neighboring formula evidence.
- **Dependency impact:** current-sheet and cross-sheet cell/range references, direct upstream and
  downstream cells, bounded downstream counts, path examples, critical-cell impact, cycles, and
  truncation flags.
- **YAML policy:** nine deterministic built-in rule types with `PASSED`, `FAILED`, `WARNING`,
  `SKIPPED`, or `ERROR` results.
- **Explainable risk:** configurable, deduplicated point contributions capped at 100.
- **One report contract:** JSON, offline autoescaped HTML, CLI, web, and CI use `ReviewResult`.
- **Safe local interfaces:** Typer CLI, bounded Streamlit uploads, and a read-only GitHub Actions
  workflow. No workbook is uploaded by SheetProof and no API key is required.

## 4. Quick start

From a clone of this repository:

```bash
python -m venv .venv
python -m pip install --upgrade pip
python -m pip install -e ".[dev,web]"
python examples/generate_demo_workbooks.py

sheetproof compare \
  examples/generated/before.xlsx \
  examples/generated/after_safe.xlsx \
  --config examples/sheetproof.example.yml \
  --json build/safe-review.json \
  --html build/safe-review.html
```

On PowerShell, the multiline continuation character is a backtick rather than `\`; the same
command can also be entered on one line. Open `build/safe-review.html` locally to inspect the
self-contained report.

The demo generator also creates `after_risky.xlsx`. Comparing it is expected to return exit code
`1` because it includes formula overwrites and policy violations:

```bash
sheetproof compare examples/generated/before.xlsx examples/generated/after_risky.xlsx \
  --config examples/sheetproof.example.yml \
  --json build/risky-review.json \
  --html build/risky-review.html
```

## 5. Installation

Requirements:

- Python 3.11 or newer;
- a supported desktop/server platform for Python and openpyxl;
- source access to this repository (the alpha is not documented here as a published PyPI
  package).

Runtime-only editable installation:

```bash
python -m pip install -e .
```

Include the local web demo:

```bash
python -m pip install -e ".[web]"
```

Contributor installation:

```bash
python -m pip install -e ".[dev,web]"
```

Docker is not required. Verify the installed command with `sheetproof version`.

## 6. CLI usage

### Compare workbooks

```bash
sheetproof compare BEFORE.xlsx AFTER.xlsx \
  [--config sheetproof.yml] \
  [--json report.json] \
  [--html report.html] \
  [--fail-on LOW|MEDIUM|HIGH|CRITICAL]
```

`--fail-on` overrides `block_risk_level` from the configuration. Without either an explicit
flag or configuration file, the default gate is `HIGH`. A failed or errored rule always blocks.

Exit codes are stable CI contracts:

| Code | Meaning |
| ---: | --- |
| `0` | Review completed and did not meet a blocking condition. |
| `1` | A rule failed/errored or risk reached the configured blocking level. Reports are still written. |
| `2` | Workbook, path, or configuration input was invalid. |
| `3` | An unexpected internal error occurred. Use `--debug` locally for a traceback. |

### Inspect one workbook

```bash
sheetproof inspect workbook.xlsx
sheetproof inspect workbook.xlsm --json build/snapshot.json
```

Inspection records workbook facts; it does not calculate formulas or run macros.

### Validate policy and print version

```bash
sheetproof rules validate sheetproof.yml
sheetproof version
```

Use `sheetproof --help` or `sheetproof COMMAND --help` for complete option help.

## 7. Web page

Install the `web` extra and start Streamlit from the repository root:

```bash
python -m pip install -e ".[web]"
streamlit run web/app.py
```

The page accepts before/after workbooks and an optional policy, then shows summary, high-risk,
cell, formula-overwrite, structure, and rule tables with JSON/HTML downloads. Each upload is
limited to 25 MiB. Files are copied to a private temporary directory for the review and removed
afterward. The web page is only an adapter around `ReviewService`.

## 8. `sheetproof.yml`

Configuration is optional and parsed with a duplicate-key-rejecting loader derived
from PyYAML's `SafeLoader`, then validated by strict Pydantic models.
Misspelled fields or invalid rule shapes are rejected.

```yaml
rules:
  - name: Cash-flow formulas must remain formulas
    type: formula_required
    range: "现金流!B5:B40"
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

  - name: Do not add VBA
    type: no_macro_added
    severity: critical

  - name: Margin is plausible when a cached value exists
    type: numeric_range
    target: "利润表!F18"
    min: 0
    max: 0.6
    severity: medium

  - name: Profit sheet exists
    type: required_sheet
    sheet: "利润表"

  - name: Scratch sheet is forbidden
    type: forbidden_sheet
    sheet: "临时导入"

  - name: Keep the review focused
    type: max_changed_cells
    max: 25
    failure_status: WARNING

risk_weights:
  formula_overwritten: 40

critical_cells:
  - "现金流!B22"

block_risk_level: HIGH
max_dependency_depth: 10
max_dependency_nodes: 10000
```

For `numeric_range`, a formula cell without a cached result is `SKIPPED` with a reason; SheetProof
does not calculate or invent the value. Set `failure_status: WARNING` on any rule that should report
an advisory finding instead of a blocking `FAILED`; unexpected evaluator failures remain `ERROR`.
See
[`examples/sheetproof.example.yml`](examples/sheetproof.example.yml) for a complete demo policy.

## 9. GitHub Actions

The repository includes `.github/workflows/sheetproof.yml` and a reusable composite action in
`action/action.yml`. The required workflow uses base-owned `pull_request_target`, an explicitly
read-only `contents` token, no repository secrets, checkout with persisted credentials disabled,
and no PR comment permission. Its trust boundary is deliberate: the current base commit is
checked out to `.sheetproof-trusted` for the Action and Python implementation, while the
base repository's generated `refs/pull/<number>/head` is checked out separately to
`.sheetproof-pr` only as Git-object input. No pull-request program or workflow is executed.

Minimal local-action usage after checkout:

```yaml
permissions:
  contents: read

steps:
  - name: Check out trusted implementation
    uses: actions/checkout@v4
    with:
      ref: ${{ github.event.pull_request.base.sha }}
      path: .sheetproof-trusted
      fetch-depth: 1
      persist-credentials: false

  - name: Check out pull request objects as data only
    uses: actions/checkout@v4
    with:
      repository: ${{ github.repository }}
      ref: refs/pull/${{ github.event.pull_request.number }}/head
      path: .sheetproof-pr
      fetch-depth: 0
      persist-credentials: false

  - name: Review changed workbooks
    uses: ./.sheetproof-trusted/action
    with:
      base-ref: ${{ github.event.pull_request.base.sha }}
      head-ref: ${{ github.event.pull_request.head.sha }}
      config-ref: ${{ github.event.pull_request.base.sha }}
      repository: .sheetproof-pr
      config: sheetproof.yml
      fail-on: HIGH
      output-dir: sheetproof-reports
```

The trusted helper computes the unique merge base of the **current** trusted base and head, uses
that merge base for change discovery and every before-side workbook blob, and follows both paths
of a Git rename. This prevents new destination-branch commits from being misattributed to a
behind-base pull request. It reads both sides directly from the data checkout's Git object
database into controlled temporary filenames; it does not import Python from the PR checkout.
The policy is independently read from `config-ref`, which defaults to the current `base-ref`, so a
pull request cannot weaken its own gate by editing `sheetproof.yml`, and a recently updated base
policy takes effect even when the PR fork point is older. To intentionally test a proposed policy,
use a separate non-gating workflow; do not change `config-ref` to an untrusted PR SHA on a
required check.

Action inputs are:

| Input | Default | Meaning |
| --- | --- | --- |
| `base-ref` | required | Current trusted base; its merge base with `head-ref` supplies before-side blobs. |
| `head-ref` | `HEAD` | Head commit used for after-side workbook blobs. |
| `repository` | `.` | Git checkout containing both commits; reports are written inside it. |
| `config` | `sheetproof.yml` | Repository-relative policy path; empty disables the file. |
| `config-ref` | `base-ref` | Trusted commit from which `config` is read. |
| `fail-on` | `HIGH` | Risk level that blocks (`LOW`, `MEDIUM`, `HIGH`, or `CRITICAL`). |
| `output-dir` | `sheetproof-reports` | New report directory relative to `repository`; the leaf must not exist. |
| `allow-unpaired` | `false` | Advisory mode: report added, deleted, or otherwise unpaired workbooks as `SKIPPED` instead of blocking as `UNREVIEWABLE`. |
| `python-version` | `3.11` | Python version selected by `actions/setup-python`. |

A newly added workbook has no base version and a deleted workbook has no head version; those
cases are explicitly marked `UNREVIEWABLE` and block by default because a semantic comparison
requires both sides. Rename-to/from a non-workbook is handled the same way. Set the Action's
`allow-unpaired: true` only for advisory workflows to report them as `SKIPPED`. Multiple
comparisons are aggregated: exit `1` blocks on findings or unreviewable changes,
while input/internal errors preserve codes `2`/`3`.

Before analysis, the trusted helper rejects a pre-existing output leaf, a leaf symlink, and any
symlink in its path, then creates the leaf itself. The workflow uploads only the directory path
published by that helper, including when findings block the gate; it never uploads a static path
that a PR could pre-seed. A missing trusted policy uses built-in defaults and reports that fact.
The default workflow runs only when a PR changes an `.xlsx` or `.xlsm`.

## 10. JSON output

Every interface is based on one `ReviewResult`. Top-level fields are:

```json
{
  "tool_version": "0.1.0",
  "reviewed_at": "2026-01-01T00:00:00Z",
  "before_file": {},
  "after_file": {},
  "summary": {},
  "structure_changes": [],
  "cell_changes": [],
  "formula_changes": [],
  "dependency_impacts": [],
  "rule_results": [],
  "risk_factors": [],
  "limitations": [],
  "errors": []
}
```

Each change contains a type, risk level, location, before/after facts or formulas, human-readable
description, and evidence. Dependency entries expose limits and cycle/truncation state. Rule
results include status, severity, reason, location, and evidence. Treat fields as typed evidence;
do not parse the prose description to implement a gate.

## 11. Risk scoring

Risk is deterministic and additive, with duplicate `(risk type, location)` findings counted once
and the total capped at 100. Default contributions include:

| Finding | Points |
| --- | ---: |
| VBA added | 40 |
| External link added | 35 |
| Formula overwritten/deleted | 30 |
| Hidden sheet added | 20 |
| Formula range reduced | 20 |
| Formula changed | 15 |
| Hidden rows/columns changed | 10 |
| 25 or more changed cells | 10 |
| Formula added | 8 |
| Value/text/style changes | 2 / 1 / 1 |

Bands are `LOW` 0–19, `MEDIUM` 20–49, `HIGH` 50–79, and `CRITICAL` 80–100.
`risk_weights` can override named weights with integers from 0 through 100. Every applied item is
returned in `risk_factors` with its points, location, description, and source evidence.

Risk scoring prioritizes review; it is not a probability, confidence score, or absolute verdict.

## 12. Privacy and security

- Analysis is local and requires no account, API key, paid API, database, or cloud service.
- SheetProof does not upload workbook contents or telemetry.
- VBA is detected from package contents but never executed.
- Formulas are read as text and are never executed or recalculated.
- External workbook links are recorded but never opened or fetched.
- The parser streams and validates the OOXML central directory before materializing it, then
  enforces 100 MiB file, 512 MiB expanded-archive, 10,000 archive-entry, 1,000,000 materialized
  cell, and 100,000-cell per-merge limits. ZIP64 and multi-disk packages are rejected.
- HTML uses Jinja autoescaping and has no online CDN dependency.
- The included PR workflow has read-only repository permission and does not auto-comment.
- The PR workflow runs the Action, Python package, helper script, and policy from the base commit;
  pull-request code and policy are not used to decide their own gate.

An `.xlsx`/`.xlsm` can contain content outside SheetProof's current model. Do not treat a clean
report as malware clearance. Use least-privilege execution and read [SECURITY.md](SECURITY.md)
before processing untrusted inputs.

## 13. Known limitations

- **The current version cannot fully simulate Excel's calculation engine.** It does not
  recalculate formulas; cached results may be missing or stale.
- **Some complex formulas receive reference-level analysis only.** Unsupported constructs retain
  their raw formula and are marked `unsupported_formula_analysis` rather than silently ignored.
- Formula parsing is not a complete Excel grammar. Dynamic references such as `INDIRECT` cannot
  be resolved reliably from text alone. Formula reference expansion is capped per formula and at
  200,000 generated references per review; dependency traversal has separate depth/node bounds.
- **VBA is detected only and is never executed.** SheetProof does not decompile, analyze, verify,
  or remove macro code.
- openpyxl does not preserve or expose every Excel feature. Charts, slicers, Power Query, data
  models, embedded objects, digital signatures, conditional-formatting semantics, and every style
  detail are not fully compared by this MVP.
- ZIP64 and multi-disk OOXML packages are outside the MVP security profile even when their
  logical workbook contents might otherwise be small.
- External-link detection is evidence collection, not a guarantee that every possible link or
  connection type was found.
- GitHub CI semantically compares modified and renamed workbook pairs. Added/deleted workbooks
  fail closed as `UNREVIEWABLE` by default because one side is absent; advisory workflows may set
  `allow-unpaired: true`.
- **SheetProof does not replace a professional financial audit, security scanner, or business
  approval process.**
- **The risk score is a review aid, not an absolute conclusion.** A low score does not prove
  correctness, and a high score does not prove malicious intent.

## 14. Roadmap

Potential post-MVP work, subject to design and tests:

- a more complete formula AST and richer formula-range explanations;
- optional LibreOffice headless recalculation with explicit trust boundaries;
- workbook history and batch review;
- visual dependency graphs;
- opt-in GitHub PR comments with tightly scoped permission;
- custom deterministic Python rule plugins and reusable policy templates;
- MCP and evidence-only AI explanation interfaces.

These items are not implemented promises. Accounts, billing, cloud file storage, and autonomous
workbook modification are outside the current direction.

## 15. Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) and [AGENTS.md](AGENTS.md). Changes to parsing,
diff logic, formulas, dependency analysis, built-in rules, scoring, or exit-code behavior must
include tests. Use synthetic workbooks and never attach sensitive business data to an issue.

Bug reports and focused feature proposals are welcome through the provided issue forms. Security
reports follow [SECURITY.md](SECURITY.md). Pull requests should use the repository template and
record the exact checks run.

## 16. License

SheetProof is available under the [MIT License](LICENSE).
