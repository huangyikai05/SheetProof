# Issue draft: Improve Windows installation documentation

Suggested labels: `documentation`, `good first issue`, `help wanted`

## Background

Windows users often alternate between Command Prompt and PowerShell, whose activation syntax,
line continuation, quoting, and executable lookup differ. Clear Windows-native commands would reduce
installation reports caused by pasting PowerShell syntax into `cmd.exe` or using an unsupported
Python interpreter.

## Goal

Add a short, tested Windows installation and troubleshooting guide for Python 3.11+ covering both
PowerShell and Command Prompt.

## Suggested implementation

- Show separate copyable blocks for `py -3.11 -m venv .venv`, activation, package installation,
  `tabulint version`, and a synthetic demo comparison.
- Explain when to use `python -m tabulint` or the virtual-environment executable if the console
  entry point is not on `PATH`.
- Cover paths containing spaces and non-ASCII characters without requiring administrator access.
- Include focused troubleshooting for execution policy, missing `py`, stale virtual environments,
  and invoking a different Python than the one used for installation.
- Verify all project-specific commands against the released package or an isolated wheel install.

## Acceptance criteria

- [ ] PowerShell and Command Prompt instructions are clearly separated and independently copyable.
- [ ] The guide installs into a virtual environment without administrator privileges.
- [ ] Version, help, inspect, and compare examples match the current CLI.
- [ ] At least one documented example uses a path containing spaces or Chinese characters.
- [ ] Troubleshooting commands identify the active Python and `tabulint` executable.
- [ ] A Windows reviewer or Windows CI smoke run verifies the commands before merge.

## Non-goals

- Shipping a Windows installer, executable bundle, Excel add-in, or GUI rewrite.
- Supporting Python versions below the package's declared minimum.
- Documenting GitHub authentication or unrelated shell administration.
