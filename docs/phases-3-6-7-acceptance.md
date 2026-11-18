# Phases 3, 6 and 7 acceptance

Historical checkpoint: the Phase 7 endpoint blocker mentioned below was resolved
on November 17, 2026; see [live acceptance](phase7-live-acceptance.md).

Development version: **0.5.0.dev1**, branch `phases-3-6-7`.
Delivery order is 3, 6, 7, 4, 5. The [implementation plan](next-phases-plan.md)
sets the workstreams and gates. The frozen 0.4.0rc1 checkpoint is unchanged.

## Phase 3 — analytical workbook composition

- [x] XlsxWriter composition and attachment to finalized worksheet workbooks.
- [x] Preserve worksheet cells, tables, names, styles, drawings and existing
  charts. Reject connection conflicts and unsupported model relationships.
- [x] Native attachment uses Excel's `powerPivotData` relationship. This fixes
  the earlier `create(workbook=...)` path that Excel treated as an empty model.
- [x] Validated CUBEMEMBER, CUBEVALUE, CUBESET, CUBESETCOUNT and CUBERANKEDMEMBER
  helpers. CUBEMEMBERPROPERTY is rejected for embedded models because Excel
  does not support it for Power Pivot models.
- [x] Native Excel 16.0.20625.20044: measure/member captions, totals 100/40/60,
  set count 2, ranked member Beta, worksheet sum 100, literal leading-equals
  text, three charts, report total 100 and save/reopen all pass.
  Evidence: workspace `work/phase3-dashboard-native-v7.xlsx.json`.
- [x] Ordinary file-open check, without COM Open/Initialize/Refresh/Update:
  CUBE cells calculate to their expected values automatically. The new
  PivotTable remains blank and its PivotChart has zero series until a layout
  update, matching the documented Phase 2 boundary. Existing ordinary charts
  retain one series each. Evidence: `work/phase3-cli-open-verified-v2.xlsx.json`.
- [x] Actual connected slicer selection 40, timeline selection 30, clear 100,
  and reopen pass in `work/phase3-controls-native.xlsx.json`.
- [x] 100,000-row/100-group composed model: CUBE/report total 4,999,950,000 and
  selected group 49,950,000 survive native save/reopen. Authoring measured
  15.29 seconds, 47,355,118 peak Python allocation bytes, 780,310 workbook bytes.
  These are one local synthetic measurement, not general performance guarantees.
  Evidence: `work/phase3-scale-100000.benchmark.json` and
  `work/phase3-scale-native-verified.xlsx.json`.
- [x] Final local integrated suite: 801 passed, zero skips or failures.
- [x] Clean installed-wheel optional-feature checks and source payload parity.
- [x] All nine installed-package CI jobs pass: Windows/Linux/macOS,
  Python 3.10/3.12/3.14.

## Phase 6 — optional bounded DAX evaluation

- [x] Audit and pin the published pbix-mcp 0.9.153 evaluator; keep it optional
  and unloaded when importing the core package.
- [x] Read-only model snapshot adapter; no fabricated Excel caches or source
  mutation. Explicit unsupported/error handling instead of false BLANKs.
- [x] Independent native oracle: 14 scalar cases for aggregation, BLANK,
  iterators, CALCULATE, filtering, relationships and totals. Native results
  match expected values before and after saving/reopening.
  Evidence: `work/phase6/native-oracle-v2-results/dax-oracle-v2.json`.
- [x] Final evaluator guards, 17 native comparisons and reproducible 1,000/
  100,000-row benchmarks recorded in [evaluation evidence](phase6-evaluation-evidence.json)
  and [backend audit](phase6-backend-audit.json). The published wheel is pinned;
  its runtime files are checked rather than trusting a matching version string.
- [x] Final local integrated suite: 801 passed, zero skips or failures.
- [x] Clean installed-wheel optional-feature checks and source payload parity.
- [x] All nine installed-package CI jobs pass: Windows/Linux/macOS,
  Python 3.10/3.12/3.14.

This is a verified expression subset, not a complete DAX engine or an Excel
formula evaluator. Strict `==` syntax was rejected by the tested embedded Excel
model; backend-only results do not establish Excel compatibility.

## Phase 7 — external connection and compatibility work

- [x] Typed external semantic-model connections and atomic package insertion,
  separate from embedded import/refresh sources. Credential-free definitions.
- [x] Optional ADOMD query adapter with row limits and explicit errors; mock
  transport tests establish adapter behavior, not endpoint interoperability.
- [x] Native Excel recognizes the external connection as OLAP, cube Model,
  COM connection type 1, while retaining the embedded model. No external query
  was attempted. Evidence: `work/phase7/native-metadata-result.json`.
- [x] Reconcile prior native SQL import/refresh evidence and documentation.
- [x] Broader scale coverage from the composed 100,000-row model above.
- [ ] Actual external query/refresh/save/reopen against Analysis Services.
- [x] Final local integrated suite: 801 passed, zero skips or failures.
- [x] Clean installed-wheel optional-feature checks and source payload parity.
- [x] All nine installed-package CI jobs pass: Windows/Linux/macOS,
  Python 3.10/3.12/3.14.

**External endpoint gate is blocked:** no SSAS or Power BI engine is installed.
The official Microsoft-signed SQL Server 2025 Developer bootstrapper was
downloaded, but Windows refuses to execute it without OS administrator elevation
(including `/Help`). Sandbox escalation does not provide that Windows privilege.
A normal `RunAs` retry reached Windows consent, but that prompt is not
exposed through the available desktop automation; the caller was interrupted.
The protected consent process could not be canceled (access denied). No server
installation was performed. The artifact is
`work/sql-analysis-services/SQL2025-Developer.exe`. An installed
test SSAS instance or an authorized reachable endpoint is required to close this
gate. Metadata-only validation is not live-connection certification. Phase 7
must not be declared complete while this gate is open.

Phases 4 and 5 remain planned; no desktop agent/session API is claimed here.

## Integrated verification

Runtime revision: `5524e08a9827a58350be6dbc2e52a27ae3516ba8`.
The local suite completed in 500.88 seconds. Clean wheel installation checked
composition, CUBE cache boundaries, external connection add/inspect, 17 scalar
evaluation cases, CLI entry points, all 66 runtime payload files, and archive
licenses/provenance/exclusions. Native evidence is summarized with hashes in
[the portable record](phases-3-6-7-native-evidence.json).

[CI run 43002859173](https://github.com/sean-palmer-coding/excel-model-lab/actions/runs/43002859173)
passed all nine jobs. Platform/local-fixture skips are allowed by the existing
suite; the local source run had none. Final documentation-only changes do not
change the tested runtime revision.
