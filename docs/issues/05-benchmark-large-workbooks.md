# Issue draft: Benchmark large workbooks

Suggested labels: `performance`, `help wanted`

## Background

Tabulint protects untrusted-input processing with parser and graph bounds, but maintainers need a
repeatable way to observe runtime, peak memory, report size, and truncation behavior as synthetic
workbooks grow. One-off measurements are not comparable and should not become unsupported marketing
claims.

## Goal

Create a reproducible, local benchmark suite for representative synthetic workbook sizes and change
patterns, with enough environment metadata to compare runs responsibly.

## Suggested implementation

- Generate deterministic workbooks from fixed parameters rather than committing large binaries.
- Cover dense values, copied formulas, cross-sheet dependencies, style changes, and bounded graph
  traversal at small, medium, and opt-in large sizes.
- Record wall-clock time, peak memory where portable, workbook size, changed-cell count, and any
  parser/graph limit reached.
- Keep the largest cases out of required pull-request CI; run a small smoke benchmark in CI and
  document an explicit command for full local runs.
- Store machine-readable results with Python, OS, CPU, and Tabulint revision metadata.

## Acceptance criteria

- [ ] A documented command regenerates all benchmark inputs from synthetic data.
- [ ] Two runs with the same parameters generate equivalent workbook structure and findings.
- [ ] Output records timing, input dimensions, relevant limits, version, and environment metadata.
- [ ] CI has a fast functional smoke test without a fragile timing threshold.
- [ ] The full benchmark completes within documented resource bounds or reports a bounded stop.
- [ ] No proprietary workbook, personal data, or source-unknown fixture is committed.

## Non-goals

- Removing safety limits to improve a benchmark score.
- Promising production throughput from one developer machine.
- Making microbenchmark timing a flaky required PR gate.
