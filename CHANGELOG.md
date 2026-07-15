# Changelog

All notable user-visible changes to SheetProof are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). SheetProof uses
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

No user-visible changes have been recorded after the v0.1.0 release candidate.

## 0.1.0 - Unreleased

This section describes the first public release candidate. A calendar date will be added only
when the release is actually published.

### Added

- Bounded, non-executing inspection of <code>.xlsx</code> and <code>.xlsm</code> packages.
- Semantic cell, formula, copied-pattern, and workbook-structure comparison.
- Formula-overwrite, formula-range-reduction, and broken copied-pattern detection.
- Bounded formula dependency and downstream impact analysis with cycle and truncation evidence.
- Strict YAML business-rule engine with nine deterministic built-in rule types.
- Configurable, deduplicated, capped, and explainable risk scoring.
- Canonical typed review results with JSON and self-contained offline HTML reports.
- Typer CLI commands for comparison, inspection, policy validation, and version output.
- Public Python helper returning the same typed review result as other interfaces.
- Local Streamlit interface with bounded uploads and temporary-file cleanup.
- Composite GitHub Action and read-only pull-request workbook gate.
- Reproducible synthetic safe and risky workbook generator.
- CI and Trusted Publishing release-workflow preparation for the first package release.
- Public safe/risky demo guides, recording script, roadmap, and release notes.

### Security

- VBA is detected but never executed.
- External workbook links are recorded but never fetched.
- Formulas are read as text and are never executed or recalculated.
- YAML uses a duplicate-key-rejecting safe loader and strict typed validation.
- Report content is autoescaped and requires no online assets.
- OOXML paths, archive expansion, entry count, materialized cells, merges, formula expansion,
  and dependency traversal are bounded.
- The required pull-request gate uses a base-owned implementation and policy, read-only
  permissions, merge-base semantics, and an exclusively created artifact directory.
- Core findings and merge-gate decisions are deterministic and never delegated to AI text.

### Limitations

- SheetProof does not implement or emulate Excel's full calculation engine; cached formula values
  may be missing or stale.
- Formula parsing is not a complete Excel grammar. Some complex and dynamic formulas receive only
  reference-level or unsupported-analysis evidence.
- VBA behavior, charts, slicers, Power Query, data models, embedded objects, signatures, and every
  formatting semantic are not fully analyzed.
- ZIP64 and multi-disk OOXML packages are rejected by the current security profile.
- External-link detection cannot guarantee coverage of every Excel connection mechanism.
- Added or deleted workbooks cannot be semantically paired and fail closed in the default CI gate.
- Risk scores prioritize review and are not probabilities, correctness proofs, or professional
  financial-audit conclusions.

[Unreleased]: https://github.com/huangyikai05/SheetProof/commits/main
