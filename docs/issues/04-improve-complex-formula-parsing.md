# Issue draft: Improve complex formula parsing

Suggested labels: `enhancement`, `help wanted`

## Background

The current parser provides bounded reference-level analysis and explicitly preserves unsupported
inputs. Complex Excel syntax can contain quoted sheet names, external references, structured table
references, unions, intersections, and nested expressions that are difficult to classify with
simple token matching.

## Goal

Improve deterministic extraction of formula references and copied-pattern structure without
evaluating formulas or pretending to reproduce Excel's calculation engine.

## Suggested implementation

- Establish a synthetic formula corpus grouped by supported and intentionally unsupported syntax.
- Introduce or extend a bounded tokenizer/AST that retains source spans and normalized references.
- Add support incrementally, beginning with quoted sheet names, absolute/mixed references, ranges,
  and nested calls before considering structured references or dynamic arrays.
- Return explicit limitation evidence for malformed, external, or unsupported constructs.
- Enforce input length, nesting depth, token-count, and traversal limits.

## Acceptance criteria

- [ ] The supported syntax matrix and normalization rules are documented.
- [ ] New cases produce stable references without changing formula text or calculating a value.
- [ ] Unsupported or malformed formulas remain visible in results with a specific limitation.
- [ ] Tests cover Unicode/quoted names, mixed references, ranges, nested calls, and adversarial depth.
- [ ] Existing formula overwrite and copied-pattern tests remain unchanged or are deliberately
      updated with equivalent deterministic evidence.
- [ ] Benchmarks show the parser remains bounded on large and deeply nested formulas.

## Non-goals

- Full Excel calculation semantics, cached-value fabrication, or formula execution.
- Fetching external workbooks or resolving external-link values.
- Heuristic or AI-based guesses presented as parsed facts.
