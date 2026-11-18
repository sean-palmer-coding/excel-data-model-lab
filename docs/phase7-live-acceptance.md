# Phase 7 live endpoint acceptance

Verified November 17, 2026. Development version **0.5.0.dev3**.
The prior unavailable-endpoint blocker is resolved for local SSAS Tabular with
Windows integrated authentication. Earlier checkpoints remain unchanged.

## Native endpoint and workbook checks

- [x] SQL Server Analysis Services 17.0.25.218, Tabular mode, instance
  `localhost\ExcelModelLab`; no embedded passwords or cloud credentials.
- [x] Official ADOMD.NET NuGet package 19.118.3, net472 client via pythonnet 3.2.1.
  Package SHA-256: `fb16292dcc7c9b37d6c21b64e8cd6bab4a336df561d3c1ef437476aaa109afff`.
- [x] Synthetic Sales model: Amount 10/20/30, labels Alpha/Beta/Gamma, Total Amount 60.
- [x] Actual Python queries return exact totals, labels, BLANK as None and empty rows.
- [x] Actual provider tests reject row-limit overflow, multiple EVALUATE result sets
  and invalid DAX; a later successful query verifies continued usability.
- [x] Nullable integer 9007199254740993 remains exact in both result rows and DataFrame.
- [x] Excel recognizes the external OLAP connection, calculates 60, saves and reopens.
- [x] Changing the server partition to 10/20/90 and reprocessing produces 120 in Python.
  The workbook file is unchanged during that server operation; no embedded model
  is copied into it. Excel initially retains 60, deliberate refresh retrieves 120,
  and save/reopen retains 120.
- [x] Native fixture run uses one owned Excel workbook at a time and closes its app.
- [x] Full local regression: **868 passed**, zero skips/failures, 542.41 seconds.
- [x] Clean installed-wheel checks: **88 passed**, all six optional extras,
  composition/CUBE/connection/DAX/CLI smoke, and all 69 runtime files matching
  source, wheel and source archive. The clean installed wheel also queried the
  actual endpoint for 120 and preserved the nullable large integer exactly.
- [x] All nine installed-wheel CI jobs pass: Windows/Linux/macOS,
  Python 3.10/3.12/3.14.

Evidence: [native results](phase7-live-native-evidence.json). Excel reports version
16.0, build 20625.0; xlwings 0.33.22. The final retained test database is
`ExcelModelLab_Phase7_Final_20261117`, with total 120. The earlier investigation
fixture `ExcelModelLab_Phase7_Test_20261117` is also synthetic and retained.

## Reproduce

Run `scripts/verify_external_semantic_endpoint.py` with `--server`, a **new**
`--database` name, `--tom-directory`, `--adomd-assembly`, and a **new** `--output`
directory. The script requires Windows Excel and the optional desktop/query/
workbook dependencies. It creates and retains a synthetic database and workbook;
it refuses to overwrite an existing database or output directory. TOM and ADOMD
client binaries are installed separately and are not bundled in this package.

## Scope

This verifies local SSAS Tabular with Windows integrated authentication, one DAX
result set, external connection authoring, CUBE querying and explicit refresh.
It does not certify Power BI/Azure authentication, external PivotTable authoring,
general DirectQuery configuration or a public remote-model deployment API.
The synthetic deployment helper is a verification tool, not such an API.
ExternalQueryResult.to_dataframe() uses object columns to avoid implicit numeric
rounding; callers choose conversions explicitly. Timeout configuration remains
unit-tested; no claim of forcibly interrupting every native provider stall is made.

Tested runtime revision: `85e81b3a20ad94ce69d9d534baed64221a438f5c`.

[CI run 44237807963](https://github.com/sean-palmer-coding/excel-model-lab/actions/runs/44237807963)
passed all nine jobs. Existing platform/local-fixture skips apply to installed
CI; the local source run had none. Final documentation-only changes preserve
the tested runtime. The supported Phase 7 milestone is complete; all previously
agreed implementation phases now have acceptance evidence within their stated
boundaries. PyPI publication remains separate.
