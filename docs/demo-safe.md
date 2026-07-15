# Safe demo: approved forecast value update

This synthetic case demonstrates a low-risk value edit that does not alter formulas, workbook
structure, external links, or a protected rule target.

## What changed

The generator copies <code>before.xlsx</code> and changes one approved forecast input:

| Location | Before | After |
| --- | ---: | ---: |
| <code>预测!D5</code> | 1100 | 1125 |

The cell is inside the policy's approved input area, <code>预测!D5:H30</code>.

## What SheetProof detected

The current deterministic review produces:

| Metric | Result |
| --- | --- |
| Risk | **2/100 LOW** |
| Changed cells | **1** |
| Changed formulas | **0** |
| Formula overwrites | **0** |
| Hidden sheets added | **0** |
| External links added | **0** |
| Rule results | **8 passed, 0 failed, 0 warnings, 1 skipped, 0 errors** |
| CLI exit code | **0** |

The numeric-range rule for <code>利润表!F18</code> is skipped because the formula has no usable
cached value. That limitation is explicit: SheetProof does not recalculate or invent a value.

## Why the score is low

The only risk contribution is the ordinary value change at <code>预测!D5</code>, worth 2 points
under the default weights. The edit does not break a formula, shorten a range, change workbook
structure, add a link, or violate the demo policy.

Low risk means “prioritize less review,” not “proved correct.”

## Reproduce it

From a repository clone with SheetProof installed:

~~~bash
python examples/generate_demo_workbooks.py

sheetproof compare examples/generated/before.xlsx examples/generated/after_safe.xlsx \
  --config examples/sheetproof.example.yml \
  --json build/safe-review.json \
  --html build/safe-review.html
~~~

Expected terminal output:

~~~text
Risk 2/100 (LOW); 1 changed cells; 0 formula overwrites; 0 failed, 0 warning, and 0 errored rules.
JSON report: <absolute path>/build/safe-review.json
HTML report: <absolute path>/build/safe-review.html
~~~

The path prefix is machine-specific. The command exits <code>0</code>.

## Report fragment

The JSON and offline HTML reports are generated from the same typed result. The relevant summary
is equivalent to:

~~~json
{
  "risk_score": 2,
  "risk_level": "LOW",
  "changed_cells": 1,
  "changed_formulas": 0,
  "formula_overwrites": 0,
  "added_hidden_sheets": 0,
  "added_external_links": 0,
  "rules_passed": 8,
  "rules_failed": 0,
  "rules_warnings": 0,
  "rules_skipped": 1,
  "rules_errors": 0
}
~~~

Open <code>build/safe-review.html</code> locally. It is self-contained and does not load workbook
content from a CDN or remote report service.

