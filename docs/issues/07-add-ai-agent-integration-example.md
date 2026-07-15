# Issue draft: Add AI agent integration example

Suggested labels: `documentation`, `help wanted`

## Background

An AI assistant can help a reviewer navigate SheetProof's structured evidence, but it must not
inspect private workbooks by default, invent facts, change risk scores, or decide whether a workbook
is correct. A narrowly scoped example would make that boundary concrete for integrators.

## Goal

Document a provider-neutral integration in which deterministic SheetProof analysis runs first and
an optional agent receives only a bounded, sanitized report summary to explain cited evidence.

## Suggested implementation

- Run `compare_workbooks` or the CLI before any optional model call and retain its exit code and
  `ReviewResult` as the canonical outcome.
- Define a minimal allowlist of report fields suitable for explanation; omit cell values, paths,
  formulas, and other workbook content by default.
- Require explanations to cite finding/rule identifiers and clearly label unsupported statements.
- Supply an offline fake explainer for tests and a provider-neutral pseudocode adapter for docs.
- Include privacy, prompt-injection, cost, logging, and data-retention warnings.

## Acceptance criteria

- [ ] The example works offline through deterministic analysis with no API key.
- [ ] Tests prove explainer output cannot alter the result, risk score, rule status, or exit code.
- [ ] Only explicitly allowlisted, bounded fields are passed to the explainer adapter.
- [ ] The documentation distinguishes program evidence from optional generated prose.
- [ ] No model SDK becomes a core dependency and automated tests make no network calls.
- [ ] A synthetic prompt-injection-like cell value is never treated as an instruction.

## Non-goals

- Adding a chatbot, multi-agent system, or model-based merge gate to the MVP.
- Uploading workbooks or complete reports to a hosted model by default.
- Using model confidence or generated text as validation evidence.
