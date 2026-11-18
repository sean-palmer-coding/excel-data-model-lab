# Native and portable text ordering

The portable collation resource captures Windows NLS version **394243** on
Windows build **22621**. Its complete identity and hashes are recorded in
`src/excel_model_lab/data/nls-en-v1.json`. Windows uses the installed native NLS
implementation; portable resource tests do not promise identical ordering across
every Windows NLS release.

On November 9, 2026, GitHub's Windows Server 2025 runner (build 26100) reported
NLS **394498** for both supported locales. The diagnostic workflow identified
U+30AB9 as the first changed scalar in the seeded random corpus:

| Implementation | Sort key (hex) |
| --- | --- |
| Captured NLS 394243 | `b7b7013f0127010100` |
| Native NLS 394498 | `fe44b7b7013e3f012627010100` |

The other five characters in that failing string had identical individual keys.
This is an operating-system weight change, not evidence of a portable key
assembly defect. Diagnostic evidence is in
[workflow run 42583358374](https://github.com/sean-palmer-coding/excel-model-lab/actions/runs/42583358374).
Run `python tools/diagnose_nls.py` on Windows to report the host identity and
reproduce the first mismatch, including individual character keys.

The full Unicode randomized comparison requires the captured NLS identity and
reports an explicit skip when the installed identity differs. Failure to read
the native identity is a test failure. The basic/combining, mixed-script,
contextual kana, and Jamo native comparisons remain enabled on every Windows
version. Captured key and resource-integrity checks run on all platforms,
including a fixed U+30AB9 regression for the pinned portable resource.
