# Security policy

SheetProof processes files that may be untrusted. Security defects in archive validation,
path handling, report escaping, configuration parsing, or CI integration are therefore
treated seriously.

## Supported versions

The project is currently an alpha. Security fixes are applied to the latest code on the
default branch and to the newest published `0.1.x` release, when one exists. Older snapshots
may not receive backports.

## Reporting a vulnerability

Do not open a public issue with exploit details or a sensitive workbook. Use GitHub's private
vulnerability reporting feature on the repository's **Security** tab. If that feature is not
available, contact a maintainer privately and ask for a secure reporting channel before
sending attachments.

Include only the minimum information needed to reproduce the problem:

- affected version or commit;
- operating system and Python version;
- impact and attack prerequisites;
- a minimal synthetic reproducer, if safe;
- suggested mitigation, if known.

Never send real financial, customer, employee, or credential-bearing workbooks. Maintainers
will acknowledge a usable report, investigate it, coordinate a fix and disclosure, and credit
the reporter if requested. Response times cannot be guaranteed while the project is in alpha.

## Security boundaries

SheetProof is designed to:

- run locally without API keys or uploads;
- read only `.xlsx` and `.xlsm` OOXML archives within configured resource limits;
- stream-count ZIP central-directory records before Python materializes archive entries, and
  reject ZIP64, multi-disk, unsafe-path, duplicate-path, oversized, and over-compressed packages;
- detect VBA presence without executing VBA;
- record external links without fetching them;
- parse configuration with PyYAML's safe loader;
- autoescape workbook-controlled content in HTML reports;
- use read-only GitHub Actions permissions by default;
- run the CI Action, Python implementation, helper script, and policy from a separate trusted
  base-commit checkout rather than from pull-request head code.

The included base-owned `pull_request_target` workflow checks the current trusted base into
`.sheetproof-trusted` and the base repository's generated pull-request head ref into
`.sheetproof-pr`. It invokes `./.sheetproof-trusted/action`, passes the current base SHA as
`config-ref`, and treats the second checkout only as a Git object source and artifact destination.
The helper computes the base/head merge base for discovery and before-side blobs, while policy
continues to come from the current base. It runs Python from its trusted source directory with
inherited `PYTHONPATH`/`PYTHONHOME` removed, preventing a PR-local `sheetproof` package from
shadowing the trusted implementation. Do not collapse these checkouts, execute PR code, add
secrets, or point `config-ref` at an untrusted PR commit in a required gate.

The report directory is also part of the boundary. The trusted helper requires its leaf not to
exist, rejects symlink components, and creates the leaf itself. Artifact upload is conditional on
the helper publishing that validated path; a static, PR-preseeded directory is never uploaded.

Formula analysis and dependency traversal use explicit per-formula, per-review, depth, and node
budgets. When a budget is exhausted, the result records truncation and unsupported analysis
instead of silently treating the formula as fully analyzed.

SheetProof does not sandbox Python itself, emulate Excel, remove malware, validate digital
signatures, or certify a workbook as safe. Run it with least privilege and treat reports as
review evidence, not as a substitute for malware analysis or professional audit.

## Public security discussions

Hardening ideas that do not reveal an exploitable weakness are welcome as normal issues.
Please avoid attaching proprietary workbooks; generate a small synthetic case instead.
