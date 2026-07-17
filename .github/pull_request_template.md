## Goal

Describe the workbook-review problem and the deterministic outcome of this change.

## Implementation

- Architecture/components changed:
- Public CLI, configuration, model, or report changes:
- Security boundaries affected:

## Evidence and tests

List the exact commands run and their real output. Do not mark a command as passed unless it was
actually run.

```text
pytest
ruff check .
mypy tabulint
```

## Local demonstration

Provide commands that generate synthetic workbooks and reproduce the behavior. Do not attach
proprietary or sensitive workbooks.

## Known limitations

State unsupported inputs, formula-analysis limits, resource bounds, and follow-up work.

## Checklist

- [ ] Verdicts and gates come from deterministic program logic, not AI text or confidence.
- [ ] Core parser/diff/formula/rule/risk changes include focused tests.
- [ ] No macro, formula, embedded command, or external link is executed or fetched.
- [ ] Reports preserve evidence and escape workbook-controlled content.
- [ ] README/config examples and `CHANGELOG.md` match implemented behavior.
- [ ] Tests use synthetic workbooks and contain no secrets or private data.
- [ ] `pytest`, `ruff check .`, and `mypy tabulint` were run, or failures are explained above.
