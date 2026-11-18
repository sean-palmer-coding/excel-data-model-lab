# Implementation status — November 18, 2026

Current release candidate: **0.5.0rc1**, published on PyPI.
See [publication acceptance](docs/publication-acceptance.md) for the clean local
run (864 passed, 50 expected skips) and public-index installation verification. Review fixes include legal internal IDs
for punctuated captions, context-aware column cloning, and synthetic regression
fixtures. Local source/installed-package and native Excel verification are
recorded in [review acceptance](docs/review-fixes-acceptance.md).

Release validation runs locally. Hosted Actions are manual-only. The older CI
results below describe historical checkpoints, not a new run for this release.
The five legacy fixtures with unrecorded redistribution rights are excluded
from release packages and the prepared public tree. Personal author/local-path
metadata has been removed from the included synthetic workbook fixtures.

## Historical development checkpoints

The previous development checkpoint was **0.5.0.dev3**. Phase 7 now passes local SSAS Tabular
live-query/Excel refresh/save-reopen acceptance; see
[the current record](docs/phase7-live-acceptance.md). Nullable large integers are
preserved in query DataFrames, and failed connection setup releases its client.
All **868 local tests**, **88 installed-package tests**, and **nine CI jobs** pass.
All agreed implementation phases are complete within their documented scope.
Code is on `phase7-live-verification`; PyPI publication remains separate.

The **0.5.0.dev2** checkpoint added phases 4 and 5: optional Windows
desktop sessions and structured agent plans/checkpoints/file handoff. Native
Excel acceptance passed, including changed model/Pivot/CUBE totals after handoff
and reopen. All **866 local tests**, **66 installed-package tests**, and **nine CI
jobs** passed. Phases 4 and 5 are complete within their documented boundaries.
Code is on `phases-4-5`. See [current acceptance](docs/phases-4-5-acceptance.md).

The previous **0.5.0.dev1** checkpoint covers phases 3, 6 and 7. Workbook composition
has native dashboard/CUBE/control/save-reopen and 100,000-row acceptance.
Optional evaluation is bounded by independently checked expression semantics.
External connection authoring had native metadata recognition; its live endpoint
gate was blocked at that checkpoint and is now resolved by dev3. See
[the historical acceptance record](docs/phases-3-6-7-acceptance.md)
for its original completed and blocked gates. All 801 local tests and nine installed-
package CI jobs pass. Phases 3 and 6 are complete within their documented
supported boundaries; the Phase 7 gate recorded there is superseded by dev3. Code is on
`phases-3-6-7`. The checkpoint below remains separate.

The **0.4.0rc1 Phase 2 milestone is complete**: 712 local tests and all nine
cross-platform CI jobs pass. The source adds native model-backed PivotTables,
PivotCharts, member filters, label sorting, chart formatting, slicers, timelines,
and guarded report/control inspection, edits and removal. Headless authoring
requires no Excel installation. Native Windows Excel verifies report values,
working controls, and save/reopen. See the current
[Phase 2 acceptance record](docs/phase2-acceptance.md) for verification status.

Fresh reports need a layout update in the tested Excel build; ordinary open or
Refresh All does not populate them. The [report guide](docs/phase2-reports.md)
documents the UI action and optional Windows helper. Unprocessed or stale
calculated-column member inventories are rejected. Empty stored text (`""`) is
also rejected explicitly; callers may choose `None` for BLANK.

Code is on the private GitHub repository's `phase2-model-reports` branch.
PyPI publication is deferred. The evidence below describes the separately
frozen model-authoring checkpoint and is not the Phase 2 test count.

## Frozen model-authoring checkpoint — November 8, 2026

The **supported expanded MVP is complete**. This is a locally verified
`0.3.1rc1` candidate; no package or repository has been published.

## Verification

- Frozen Windows source suite: **512 passed**, 472.17 seconds.
- Full Linux installed-wheel suite: **506 passed, 6 skipped**, 417.73 seconds.
  The skips are four Windows NLS comparisons and two Windows ACL checks.
- Wheel, source distribution and wheel rebuilt from the source distribution
  contain byte-identical package code/resources. Licenses, provenance and
  exclusion of unidentified legacy fixtures were checked.
- Clean Windows installed-wheel regression: **148 passed**, 161.78 seconds.
  Final artifact hashes and source commit are in the distribution verification report.
- Native evidence and hashes are in `docs/native-acceptance-2026-11-08.json`.
  `LOCAL_HANDOFF.md` preserves the investigation history, including superseded
  failures; those historical entries are not the current implementation status.

## Implemented and exercised

The unified reader exposes tables, projections, chunked/memory-mapped reads,
PBIXRay metadata and Excel-specific metadata. Authoring covers new/existing
models, tables and stored columns; replacement, growth, shrinkage, clearing and
repopulation; relationships and validation; hierarchies, sorting, visibility and
formatting; measures and KPIs; stored date-table keys and perspective lifecycle;
calculated-column CRUD with six explicit types; and dependency-aware structural
edits with report-preservation guards. JSON specifications and CLI inspection,
authoring and validation are included.

Calculated relationships, custom sorts and dependent hierarchies rebuild after
input changes through Excel's native processing. Known relationship keys are
validated immediately; pending DAX results produce an explicit deferred
validation warning rather than invented uniqueness results.

Python CSV/Parquet/caller-managed SQL refresh and native linked-table/OLE DB
bindings are available. Native two-source SQL refresh, rebind and detach passed
save/reopen and exact Python readback. The final preservation smoke test returned
totals 75 and 300 after automatic session authentication and a fresh Excel reopen.
Source/schema extensions survive rebinding; ambiguous schema edits fail atomically.

Native coverage also includes exact numeric boundaries, one million rows with
100 distinct groups, tiny dictionaries, Unicode relationship equality, nullable
integer/Currency/decimal/text lifecycles, empty calculated tables, all six
calculated types, measure removal, generalized KPI definitions, perspective
membership and date keys with alternate-key relationships. A combined workbook
returned four sales rows, total 75, calculated total 150, four distinct audit IDs
including a blank, and three calendar rows; every stored table and calculated
result matched independent Python readback after save/reopen.

## Explicit boundaries

- The writer needs neither Excel nor .NET. Excel performs DAX calculation;
  newly authored or changed calculations have unprocessed caches.
- Calculated date-table keys remain guarded. The tested computed-date key
  causes Excel's circular-dependency error even with corrected native key
  cardinality. This is not a claim about every possible DAX expression/version.
- Integer authoring accepts -9223372036854775807 through 9223372036854775806;
  Currency accepts -922337203685477.5807 through 922337203685477.5806.
  Adjacent storage endpoints are readable but rejected for new authoring after
  demonstrated native aggregate failures.
- Pinned text collation covers supported en-US/en-GB characters, including tested
  Japanese/Jamo behavior. Unsupported locales/codepoints and unknown storage
  shapes fail explicitly. Datetime authoring is timezone-naive to whole seconds.
- High-level authoring retains dataframes; bounded codecs do not make it an
  out-of-core authoring engine. Identifier and report-dependency guards remain.
- macOS and Excel 2024 are not certified by these Windows/WSL checks.
- DAX execution, worksheet calculation, Power Query authoring, PBIX writing and
  general worksheet authoring remain separate projects. Model-backed report
  layout authoring is now included in Phase 2 as described above.

The project is not claimed to be the first or only model writer. The research
record in `LOCAL_HANDOFF.md` identifies other projects and distinguishes their
published claims from independently verified Excel compatibility.
