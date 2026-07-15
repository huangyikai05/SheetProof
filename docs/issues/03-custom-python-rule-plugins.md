# Issue draft: Custom Python rule plugins

Suggested labels: `enhancement`, `help wanted`

## Background

SheetProof intentionally loads YAML safely and evaluates deterministic built-in rules. Some users
need organization-specific checks that cannot be expressed in the current schema. A plugin API can
support trusted, installed Python packages, but it must never turn workbook content or a PR checkout
into executable code.

## Goal

Define a small, versioned interface for deterministic Python rules supplied by explicitly installed
and trusted packages, while preserving canonical evidence and result models.

## Suggested implementation

- Discover plugins through Python package entry points under a SheetProof-specific group.
- Pass immutable or read-only typed workbook facts into a narrow rule protocol and require typed,
  serializable evidence in return.
- Require explicit configuration to enable a named plugin; reject unknown options and duplicate
  rule identifiers.
- Keep built-in and plugin rule results on the same orchestration path with stable ordering.
- Document that plugins are executable dependencies and must never be loaded from a workbook,
  arbitrary file path, downloaded URL, or untrusted PR checkout.

## Acceptance criteria

- [ ] A separately installed synthetic test plugin can register and run one deterministic rule.
- [ ] Missing, duplicated, disabled, malformed, and incompatible plugins fail with clear results.
- [ ] Plugin evidence serializes through the existing JSON and HTML report contracts.
- [ ] Plugin ordering is stable and risk contributions still follow existing deduplication/caps.
- [ ] Tests prove workbook text and configuration cannot select an arbitrary import or file path.
- [ ] Versioning and trust guidance are documented for plugin authors and users.

## Non-goals

- Sandboxing malicious Python packages or claiming that third-party plugins are safe.
- Installing plugins automatically, fetching code from URLs, or executing workbook macros/formulas.
- Allowing a plugin or language model to bypass the canonical verdict and risk pipelines.
