# Changelog

All notable user-visible changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project
intends to use [Semantic Versioning](https://semver.org/spec/v2.0.0.html) after its first
published release.

## [Unreleased]

### Added

- Deterministic `.xlsx` and `.xlsm` workbook inspection with bounded OOXML parsing.
- Semantic workbook, cell, formula, copied-formula-pattern, and structure comparisons.
- Bounded dependency impact analysis with cycle and truncation evidence.
- Nine built-in YAML rules and explainable, configurable risk scoring.
- Canonical Pydantic review results with JSON and self-contained HTML reports.
- Typer CLI commands for comparison, inspection, rule validation, and version output.
- Local Streamlit demonstration UI with bounded uploads and temporary-file cleanup.
- Generated safe and risky demo workbooks.
- Composite GitHub Action and read-only pull-request workflow for modified or renamed workbooks,
  with the implementation and gate policy loaded from a separate trusted base checkout.

### Security

- VBA is detected but never executed, external links are never fetched, YAML uses a safe
  loader, and report content is HTML-escaped.
- OOXML central directories are stream-counted before allocation; worksheet parts are resolved
  through OPC metadata regardless of extension; unsafe paths, non-finite numbers, and resource
  limit violations fail closed.
- Required PR gates use a base-owned, read-only `pull_request_target` workflow, trusted policy,
  merge-base comparison semantics, and an exclusively created artifact directory.
- YAML rejects duplicate keys, implicit numeric booleans, and unknown risk-weight names.
