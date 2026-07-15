# Risky demo: formula and structure regressions

This synthetic case deliberately combines formula damage, a hidden worksheet, an external-link
indicator, and policy violations. It exists to demonstrate evidence and CI blocking; it is not a
malware sample and its external workbook is never fetched.

## What changed

| Location | Change |
| --- | --- |
| <code>现金流!B13</code> | Replaces a copied formula with the fixed value <code>12500</code>. |
| <code>现金流!B22</code> | Shortens <code>=SUM(B5:B20)</code> to <code>=SUM(B5:B19)</code>. |
| <code>说明!B2</code> | Changes the owner text to <code>Unapproved automation</code>. |
| <code>隐藏调整!A1</code> | Adds a formula containing an <code>external.xlsx</code> reference. |
| <code>隐藏调整!A2</code> | Adds explanatory text. |
| <code>隐藏调整</code> | Adds the worksheet in hidden state. |

The generator creates all three workbooks locally:

~~~bash
python examples/generate_demo_workbooks.py
~~~

## What SheetProof detected

| Metric | Result |
| --- | --- |
| Risk | **100/100 CRITICAL** |
| Changed cells | **5** |
| Changed formulas | **3** |
| Formula overwrites | **1** |
| Formula range reductions | **1** |
| Formula additions | **1** |
| Hidden sheets added | **1** |
| External links added | **1** |
| Rule results | **3 passed, 4 failed, 1 warning, 1 skipped, 0 errors** |
| CLI exit code | **1** |

The formula change list contains:

1. <code>现金流!B13</code>: <code>formula_overwritten</code>, with neighboring copied-pattern
   evidence.
2. <code>现金流!B22</code>: <code>formula_range_reduced</code>, identifying excluded cell
   <code>B20</code>.
3. <code>隐藏调整!A1</code>: <code>formula_added</code>, preserving unsupported external-reference
   analysis evidence.

Four deterministic rules fail: required cash-flow formulas, allowed change range, no external
links, and no newly hidden sheets. The maximum-changed-cells rule reports one advisory warning.
The cached-value-dependent margin rule reports one explicit skip.

## Why the score is critical

The risk contributions are 20 points for the hidden sheet, 35 for the external link, 30 for the
formula overwrite, 20 for the reduced formula range, 8 for the added formula, and 1 point for each
of two text changes. Their uncapped total exceeds 100, so the documented scoring model caps the
result at **100**.

The score is an explainable review priority, not a probability and not proof of malicious intent.

## Reproduce it

~~~bash
sheetproof compare examples/generated/before.xlsx examples/generated/after_risky.xlsx \
  --config examples/sheetproof.example.yml \
  --json build/risky-review.json \
  --html build/risky-review.html
~~~

Expected terminal output:

~~~text
Risk 100/100 (CRITICAL); 5 changed cells; 1 formula overwrites; 4 failed, 1 warning, and 0 errored rules.
JSON report: <absolute path>/build/risky-review.json
HTML report: <absolute path>/build/risky-review.html
~~~

The command intentionally exits <code>1</code>. That is the expected policy gate result. Exit
<code>2</code> would indicate invalid input and exit <code>3</code> an internal failure.

## Report fragment

~~~json
{
  "risk_score": 100,
  "risk_level": "CRITICAL",
  "changed_cells": 5,
  "changed_formulas": 3,
  "formula_overwrites": 1,
  "added_hidden_sheets": 1,
  "added_external_links": 1,
  "rules_passed": 3,
  "rules_failed": 4,
  "rules_warnings": 1,
  "rules_skipped": 1,
  "rules_errors": 0
}
~~~

Open <code>build/risky-review.html</code> to inspect the linked structural, cell, formula, rule,
dependency, limitation, and risk evidence. The HTML is self-contained and autoescaped.

## External-link safety boundary

The workbook package contains a detectable reference to <code>external.xlsx</code>. SheetProof
records the link name from local workbook metadata but does not resolve, open, download, or
calculate the external workbook.

