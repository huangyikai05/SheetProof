# Issue draft: GitHub Pull Request review comments

Suggested labels: `enhancement`, `help wanted`

## Background

The current GitHub integration keeps the required workbook gate read-only and publishes reports as
artifacts. Reviewers would benefit from an optional concise PR comment, but granting write access to
a workflow that handles pull-request data requires a documented threat model and strict separation
from untrusted code.

## Goal

Design an opt-in comment publisher that summarizes an already-produced, trusted `ReviewResult`
without weakening the existing read-only gate or introducing an alternate verdict path.

## Suggested implementation

- Keep workbook analysis in the trusted base checkout and pass only a bounded, validated summary to
  a separate publisher step or reusable workflow.
- Grant `pull-requests: write` only to the publisher job; retain `contents: read`, no repository
  secrets, and no execution of code, actions, scripts, or policy from the PR head.
- Update one bot comment identified by a stable marker instead of creating a new comment on every
  synchronization.
- Escape Markdown, cap comment size, and link to the complete artifact when evidence is omitted.
- Document behavior for forks, deleted artifacts, permission denial, and concurrent updates.

## Acceptance criteria

- [ ] A security design documents trust boundaries before write permissions are introduced.
- [ ] The existing read-only Tabulint gate remains usable without comment permissions.
- [ ] The comment is derived from validated report fields and matches the CLI/JSON verdict.
- [ ] Re-running a PR updates one marked comment and cannot overwrite human comments.
- [ ] Tests cover Markdown injection, oversized evidence, forks, missing reports, and API failures.
- [ ] No secret or write-scoped token is exposed to or used while executing PR-head content.

## Non-goals

- Editing workbook files, pushing commits, approving reviews, or automatically merging a PR.
- Letting an AI-generated explanation determine pass/fail status.
- Replacing downloadable JSON and offline HTML evidence with a short comment.
