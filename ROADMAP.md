# Tabulint roadmap

This roadmap communicates direction, not promised scope or delivery dates. Priorities may change
in response to real workbook compatibility reports, security findings, and maintainer capacity.
Every shipped capability must preserve deterministic evidence, bounded processing, and the rule
that AI text never decides workbook correctness.

## v0.1.x

The v0.1 series focuses on making the existing review engine dependable for early adopters:

- fix reproducible parser, diff, rule, report, CLI, web, and Action defects;
- expand <code>.xlsx</code> and <code>.xlsm</code> compatibility tests with synthetic fixtures;
- improve Windows, macOS, Linux, and source/PyPI installation documentation;
- incorporate first-release feedback without breaking typed report contracts unnecessarily;
- profile and fix bounded performance bottlenecks on larger synthetic workbooks;
- clarify unsupported Excel features and produce better limitation evidence;
- refine examples, policy templates, and troubleshooting guidance.

## v0.2.0 candidates

These are design candidates and will ship only with explicit threat analysis, tests, and stable
typed contracts:

- a visual formula dependency graph built from the existing bounded graph evidence;
- opt-in GitHub pull-request review comments with least-privilege permissions;
- a more complete formula AST and clearer range-change explanations;
- deterministic custom Python rule plugins with a constrained loading and trust model;
- large-workbook parsing, comparison, and report performance improvements;
- richer compatibility fixtures for financial and operational workbook patterns.

## Future exploration

- an optional LibreOffice recalculation backend with clear process and file trust boundaries;
- an MCP verification server that exposes typed evidence without allowing AI to become a verdict
  source;
- an evidence-only AI explanation layer that cites deterministic findings and cannot affect
  scores, rules, or exit codes;
- bounded batch review and workbook-history comparison;
- reusable policy templates for common review domains.

## Suggested public issues

Good first public discussions include:

1. Visual formula dependency graph
2. GitHub pull-request review comments
3. Custom Python rule plugins
4. Improve complex formula parsing
5. Benchmark large synthetic workbooks
6. Add a synthetic financial workbook example
7. Add an AI agent evidence-integration example
8. Improve Windows installation documentation

Issue descriptions should state the background, goal, suggested approach, acceptance criteria,
and non-goals. Documentation-only or narrowly scoped fixture work may be labeled
<code>good first issue</code>; design-heavy work should use <code>help wanted</code> only after
maintainers document the relevant contracts and security boundaries.

## Explicit non-goals

The current direction does not include user accounts, billing, cloud workbook storage, automatic
workbook modification, an enterprise administration backend, or language-model verdicts.

