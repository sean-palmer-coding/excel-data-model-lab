# 0.5.0rc2 acceptance (2026-11-19)

Multi-table creation now assigns each cloned table a distinct internal count
measure name, while preserving ordinary captions.

- Full source-distribution suite against the installed wheel on local Windows,
  Python 3.12.14: **865 passed, 50 expected historical-fixture skips, zero failures**
  in 556.16 seconds. GitHub Actions were not used.
- Wheel/source archive pass Twine validation. Installed dependencies pass pip check.
- Native Excel: the installed rc2 wheel created two related tables and a measure
  without source-data worksheets. DAX returned total 100 and Camping 40, unchanged
  after save/reopen.
- The prior six-table demonstration also verified 24,000 sales, relationship
  filters, calculated profit, CUBE formulas and populated reports in native Excel.
- Artifacts exclude the five legacy fixtures with unrecorded redistribution
  rights. Pattern scans found no recognizable credentials or private keys.

The source records can reside entirely inside the embedded Data Model.
Worksheet copies are optional and are not implicitly linked to the model.

Published to PyPI as `excel-data-model-lab==0.5.0rc2`. A no-cache public-index
installation passed the multi-table regression; both artifact hashes match
the locally verified release.
