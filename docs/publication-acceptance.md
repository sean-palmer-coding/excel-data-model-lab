# 0.5.0rc1 release acceptance

Validated locally on Windows / Python 3.12.14 on November 18, 2026.

- Full source-distribution tests against the installed release wheel: **864 passed,
  50 skipped, zero failures**, in 501.95 seconds. Skips require historical fixtures
  intentionally excluded from redistribution.
- Twine checks pass for wheel and source archive.
- Runtime modules and bundled resources are byte-identical to the previously
  native-verified build; see [native evidence](review-fixes-native-evidence.json).
- The two included native-authoring fixture shells have personal author and
  local-path metadata removed. Every other ZIP entry, including model data,
  is byte-identical to its original.
- The five fixtures with unrecorded redistribution rights are excluded.
- Artifact scans found no recognizable PyPI/GitHub tokens or private keys.
  Pattern scanning cannot establish the absence of every possible secret.
- GitHub Actions were not used. Workflows require manual dispatch.

These checks do not extend the documented feature/compatibility boundaries.

## Published distribution

Published to https://pypi.org/project/excel-model-lab/0.5.0rc1/.
A no-cache installation from the public PyPI index passed import, model creation,
save, reopen and table readback (20, 30, 40). The existing GitHub repository
remains private pending the owner's archive/public-repository decision.

## Project rename

Published as **Excel Data Model Lab**, distribution `excel-data-model-lab`,
version `0.5.0rc1`. Imports remain `excel_model_lab` and `excelmodel`.
The renamed wheel passes 57 focused tests and public-PyPI installation with
create/save/reopen/readback. Runtime differences are installation guidance and
newline normalization only. The original publication remains available.

https://pypi.org/project/excel-data-model-lab/0.5.0rc1/
