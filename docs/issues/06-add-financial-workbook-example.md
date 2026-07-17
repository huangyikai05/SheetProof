# Issue draft: Add financial workbook example

Suggested labels: `documentation`, `good first issue`, `help wanted`

## Background

The current synthetic demos explain general safe and risky changes. A compact financial-model
example could make formula overwrites, narrowed ranges, assumption changes, and business rules more
recognizable while avoiding real company data and any claim of professional audit assurance.

## Goal

Add a script-generated, fully synthetic financial workbook pair plus rules and documentation that
demonstrate existing Tabulint behavior end to end.

## Suggested implementation

- Generate a small assumptions, forecast, and summary model with invented entities and values.
- Provide one safe revision and one risky revision containing deliberate, documented changes.
- Reuse public APIs/CLI and the existing report pipeline; do not hard-code expected verdicts into
  the generator.
- Add a YAML policy that demonstrates allowed ranges and deterministic formula/range checks.
- Explain every finding, the current calculation limitations, and how to regenerate the files.

## Acceptance criteria

- [ ] One command deterministically generates the baseline, safe, and risky synthetic workbooks.
- [ ] The safe and risky comparisons run successfully with documented commands.
- [ ] Documentation records actual, reproducible findings and does not claim financial assurance.
- [ ] Tests assert key workbook facts and expected deterministic rule/diff evidence.
- [ ] No real brand, customer, employee, account, transaction, or proprietary template is present.
- [ ] Generated reports remain offline and workbook external links are never fetched.

## Non-goals

- A comprehensive accounting model, valuation recommendation, or professional financial audit.
- Reproducing Excel's calculation engine or fabricating cached formula values.
- Bundling a real or merely anonymized corporate workbook.
