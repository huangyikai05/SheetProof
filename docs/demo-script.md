# Tabulint demo recording script

This script is designed for a clear 30–60 second terminal-and-report recording. No GIF is checked
in yet: record the real application after the release branch is verified, then review the
resulting media for readable text, secrets, local usernames, and unrelated desktop content.

## Preparation

Use a clean clone, a terminal large enough to show one-line summaries, and a browser with no
sensitive tabs. Create and activate the virtual environment before the take. If network
installation is slow, warm pip's dependency cache first, but keep the real Tabulint install
command in the recording.

PowerShell preparation:

~~~powershell
git clone https://github.com/huangyikai05/Tabulint.git
Set-Location Tabulint
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
~~~

macOS/Linux preparation:

~~~bash
git clone https://github.com/huangyikai05/Tabulint.git
cd Tabulint
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
~~~

Delete old <code>build/risky-review.json</code> and <code>build/risky-review.html</code> before
the take so the report visibly comes from the recorded command.

## Shot list

### 0–8 seconds: identify the tool

Show the repository README heading and say:

> Tabulint compares Excel workbooks with deterministic rules. It never runs formulas, macros,
> or external links.

### 8–16 seconds: install Tabulint

~~~bash
python -m pip install -e .
~~~

Do not claim a PyPI install in the recording until the package page is confirmed.

### 16–23 seconds: generate synthetic workbooks

~~~bash
python examples/generate_demo_workbooks.py
~~~

Briefly show the three printed paths: <code>before.xlsx</code>, <code>after_safe.xlsx</code>, and
<code>after_risky.xlsx</code>.

### 23–35 seconds: run the risky comparison

Enter this as one command to keep the recording readable:

~~~bash
tabulint compare examples/generated/before.xlsx examples/generated/after_risky.xlsx --config examples/tabulint.example.yml --json build/risky-review.json --html build/risky-review.html
~~~

Pause on the real summary:

~~~text
Risk 100/100 (CRITICAL); 5 changed cells; 1 formula overwrites; 4 failed, 1 warning, and 0 errored rules.
~~~

The process exits <code>1</code> by design. Do not edit the recording to imply success status
<code>0</code>.

### 35–54 seconds: open the offline report

PowerShell or Command Prompt:

~~~powershell
Start-Process .\build\risky-review.html
~~~

macOS:

~~~bash
open build/risky-review.html
~~~

Linux:

~~~bash
xdg-open build/risky-review.html
~~~

Show, in this order:

1. <code>100/100 CRITICAL</code> summary;
2. the formula overwrite at <code>现金流!B13</code>;
3. the shortened total range at <code>现金流!B22</code>;
4. the newly hidden <code>隐藏调整</code> worksheet;
5. the external-link and failed-rule evidence.

### 54–60 seconds: close on the trust boundary

Return to the report summary or README and say:

> The evidence, risk score, and exit code come from deterministic program logic. The workbook
> stays local unless you explicitly upload the generated report.

## Recording checklist

- Use the current generated fixtures and do not hard-code or overlay different numbers.
- Keep the terminal exit status visible if the recorder supports it.
- Do not show access tokens, environment variables, browser profiles, usernames, or private file
  paths.
- Verify the final GIF or video is legible at GitHub README width.
- Do not add the README media link until the final file exists at
  <code>docs/assets/tabulint-demo.gif</code> and has been reviewed.
- Prefer a short MP4 linked from release notes if a readable GIF would be excessively large.
