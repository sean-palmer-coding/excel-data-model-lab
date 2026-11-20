# Excel Data Model Lab release notes

## 0.5.0rc2 (2026-11-19)

- Fix multi-table model creation: cloned tables now have unique internal count
  measure names. Previously Excel could reject a model containing multiple
  newly created tables even though each table loaded on its own.
- Preserve ordinary table/column captions while rebinding this engine identity.
- Add a three-table regression test. Native Excel verification includes a fresh
  two-table model and the six-table, 24,000-sale demonstration workbook.
- Update project links to the public Excel Data Model Lab repository.


The distribution is now named `excel-data-model-lab`. Python imports and CLI
commands remain compatible. The earlier `excel-model-lab` publication is legacy.

## 0.5.0rc1 (2026-11-18)

First public release candidate, incorporating model authoring/reading, reports,
composition, bounded evaluation, connections, and desktop/agent workflows.

### Code review fixes

- Stored/calculated column and hierarchy captions containing periods now use legal,
  separate internal IDs. Hierarchy levels reference column IDs, and rebuilds
  retain hierarchy identities and display captions.
- Integer-mutation regression tests generate synthetic input, so they run
  from the source distribution without excluded historical workbooks.
- Column cloning updates identity fields without rewriting metadata constants.
  Columns named `Key`, `Regular`, or `Dictionary` no longer corrupt cloned
  stored or calculated columns.
- Removing a hierarchy or column no longer deletes the storage of a sibling
  whose name extends it with a period (`A` / `A.B`, `No` / `No.`).
- New PivotTable caches no longer reuse a cacheId held by an existing slicer
  or timeline cache.
- `set_integer_value` uses exact arithmetic and rejects values outside the
  column's existing range, since it does not rebuild statistics or indexes.
- Column remove/configure/clone and `rename_column` address natively renamed
  columns by their physical ID; `add_table(like=...)` no longer rewrites
  captions such as `Sales Amount`.
- Slicers on decimal columns are rejected when added rather than at save.
- Spec relationships accept dotted table names; sort-only column entries no
  longer call `configure_column`.
- The DAX validator no longer treats `"-"` or `[+]` as operators.
- Slicer sheet names match case-insensitively; lowercase PivotTable cells are
  written as uppercase references.
- ABF reading rejects negative file offsets/sizes and caps the decompressed
  model part (`MAX_MODEL_PART_BYTES`); `abf-extract` refuses colliding names.
- `tools/verify_model_reports.ps1` and `tools/verify_model_report_controls.ps1`
  now require `.xlsx` and force-disable macros and events, like the other tools.

## 0.5.0.dev3 — verified local Analysis Services connections

- Verify Windows-integrated SSAS Tabular queries and Excel external CUBE
  calculation, server data changes, explicit refresh and save/reopen (60 to 120).
- Preserve nullable large integers and provider values in `to_dataframe()` with
  object columns; callers choose any subsequent dtype conversion explicitly.
- Dispose an ADOMD connection when assigning its connection string fails.
- Add an opt-in synthetic endpoint acceptance script. It refuses an existing
  database/output directory and retains its test fixtures for inspection.

Validation: 868 local tests, 88 installed-package tests and all nine CI jobs pass,
plus actual source/installed-client SSAS queries and native Excel refresh.
See [Phase 7 live acceptance](docs/phase7-live-acceptance.md). The former endpoint
blocker is resolved for local Windows-integrated SSAS. Cloud authentication,
external PivotTable generation and general remote deployment remain unverified.

## 0.5.0.dev2 — desktop sessions and agent workflows

- Add optional Windows xlwings sessions for exact-path/PID attachment, bounded
  range editing, calculation, model report updates, and explicit save/close.
- Add structured agent plans, checkpoints, stale-plan/user-change detection,
  and staged close/edit/publish/reopen with explicit recovery information.
- Invalidate saved embedded-model CUBE tuple results when model bytes change,
  preserving external caches and requesting recalculation so old totals do not persist.
- Keep Excel/xlwings optional; headless authoring does not start desktop engines.
- Require an exclusive interaction window during workflow application; these
  optimistic guards do not make simultaneous interactive edits transactional.

Verification: 866 local tests, 66 installed-package tests, all nine CI jobs and
native Excel desktop/handoff/save-reopen acceptance passed. See the
[Phase 4/5 acceptance record](docs/phases-4-5-acceptance.md). Phase 7 live
endpoint certification remains separately blocked; no PyPI publication is implied.

## 0.5.0.dev1 — workbook composition, optional evaluation and connections

- Compose XlsxWriter worksheets with an embedded model, reports and validated
  CUBE formulas; preserve existing worksheet assets and attach models atomically.
- Correct the model relationship used when attaching to an existing workbook.
- Expose five supported embedded-model CUBE function families; reject the
  unsupported CUBEMEMBERPROPERTY operation rather than emitting a failing formula.
- Add an optional, version-pinned, read-only DAX evaluator with explicit bounded
  expression/type support. It does not write calculated Excel caches.
- Add typed external semantic connections and an optional bounded ADOMD query
  adapter. Actual live connectivity remains uncertified pending a test endpoint.
- Keep all new dependencies optional and core import free of engine startup.

This development snapshot is not a publication or completion claim for Phase 7.
See [acceptance](docs/phases-3-6-7-acceptance.md) and the individual phase guides.

## 0.4.0rc1 — verified Phase 2 release candidate

Current source adds model-backed Excel reports to the model-authoring API.
The supported milestone passes 712 local tests, all nine installed-package CI
jobs, and native Excel acceptance recorded in
[the acceptance record](docs/phase2-acceptance.md). It does not supersede
the separately frozen 0.3.1rc1 artifacts or imply publication.

- Author OLAP PivotTables with row, column and page fields, explicit measures,
  multiple measures, tabular/compact/outline layouts and grand totals.
- Author linked column, bar and line PivotCharts with styles, legend settings,
  axis titles, number formats, value labels and background colors.
- Select explicit member filters and ascending/descending label sorts.
- Author real model-backed slicers and date timelines, including connections
  to multiple PivotTables and preservation of existing charts.
- Inspect saved reports and controls, edit supported properties and layouts,
  append reports, and remove reports with dependency guards. Unknown native
  metadata and unsupported anchor/geometry edits fail explicitly.
- Preserve worksheet namespaces and unrelated parts during supported edits;
  validate staged report/control changes before writing output atomically.
- Reject zero-length stored text (`""`) explicitly: native Excel rejects the
  current writer's empty-string storage. Choose `None` explicitly for a model
  BLANK or supply nonempty text. Values are not silently normalized. Reading
  existing models and authoring DAX expressions that return `""` are unaffected.

Native Excel acceptance includes six reports and two charts, connected controls
filtering both reports from 100 to 40 (Alpha) and 30 (January), clearing to 100,
and save/reopen. The expanded example with its East-only report also passed.
Appending/editing reports preserved working controls; edited axes and selected
filters returned the expected descending IDs and total 40.

Python authors package metadata without Excel. Excel supplies DAX calculation,
report values and chart series when their layouts are updated. On the tested
Excel build, neither opening a fresh report nor ordinary Refresh/Refresh All
populates its layout. Use **Defer Layout Update → Update** in PivotTable Fields,
or the optional `tools/update_model_reports.ps1` helper, which calls the tested
`PivotTable.Update()` operation. Automatic rendering on open is not guaranteed;
this limitation remains part of the supported release contract.
Headless calculation, arbitrary MDX, measure-value sorting, Power Query
authoring and general worksheet editing are outside this release scope.
See [the report guide](docs/phase2-reports.md) for APIs and boundaries.

## 0.3.1rc1 — frozen model-authoring checkpoint

Local release candidate for the completed supported model-authoring MVP.
The frozen Windows suite passed 512 tests; the installed-wheel Linux suite
passed 506 with six Windows-only skips. Native acceptance includes the combined
feature fixture and SQL refresh/save/reopen. Nothing has been published.

## Model authoring and reading

- Create and edit embedded Excel models, tables and stored columns, including
  empty tables, growth, shrinkage, clearing and repopulation.
- Read tables and Excel-specific metadata through one API, including projected
  and chunked reads, calculated expressions, processing states, date keys,
  hierarchies and perspectives.
- Author calculated integer, decimal, Currency, datetime, Boolean and text
  columns. Explicit types survive native processing; saved Currency values,
  dates and Booleans read back with their correct Python semantics.
- Create and edit relationships, including calculated keys. Known keys are
  validated immediately; pending calculated-key validation is explicitly
  deferred to Excel. The library does not evaluate DAX.
- Author hierarchies, custom sorts, measures, generalized KPIs, stored date-table
  keys and perspectives. Calculated sorts and dependent hierarchies rebuild
  after input changes. Perspective membership and lifecycle survive reopening.
- Preserve unrelated workbook contents and reject unsafe structural edits using
  dependency checks and object-scoped validation errors.

## Native compatibility fixes

- Corrected empty calculated caches, date-key cardinality and owning
  measure-group granularity, including relationships through alternate keys.
- Compact unused dictionary members so native DISTINCTCOUNT reflects actual
  values. Verified a million-row model with exactly 100 text groups.
- Corrected tiny dictionaries, blank-member lookup segmentation, wide-integer
  dictionary flags and all-null dictionary creation/repopulation for text,
  integer, Currency and decimal columns.
- Reject exact signed-64-bit endpoints that produced blank native aggregates.
  Supported authoring ranges are integer -9223372036854775807 through
  9223372036854775806 and Currency -922337203685477.5807 through
  922337203685477.5806, inclusive. Existing endpoint values remain readable.
- Consolidated native OLE DB source views and corrected rebind/detach handling
  after Excel renames storage files. Multiple-source SQL refresh, rebind,
  detach, save and fresh-instance reopen passed with automatic session
  credentials. Unrelated source and schema extensions are preserved.
- Verified an integrated workbook combining a marked date table, stored and
  calculated-key relationships, nullable Currency and wide integers,
  calculated columns, Unicode filtering, sorting, hierarchy, KPI and perspective.

## Distribution and boundaries

Original project code is MIT licensed. The provenance inventory identifies
bundled resources; five historical binary fixtures without complete origin
records remain excluded from source distributions.

Calculated date-table keys remain guarded: the tested computed-date key causes
Excel's circular-dependency error even with corrected key metadata. Unsupported
locales/codepoints and storage shapes fail explicitly. High-level authoring
retains dataframes; streaming codecs do not make it an out-of-core writer.

DAX execution, worksheet calculation, Power Query authoring, PBIX writing and
new report/Pivot layout authoring remain outside this MVP. Windows Excel and
Linux Python results do not certify macOS or Excel 2024. Detailed native evidence
and hashes are recorded in `docs/native-acceptance-2026-11-08.json`.
