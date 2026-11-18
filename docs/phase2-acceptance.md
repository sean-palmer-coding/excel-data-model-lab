# Phase 2 acceptance checklist

Status as of November 11, 2026: **the supported Phase 2 milestone is complete**.
Native report/control gates passed on Excel **16.0.20625.20044**, all 712 local
tests passed, and all nine installed-package CI jobs passed. This is a verified
local **0.4.0rc1** release candidate, not a PyPI publication or a universal Excel
compatibility claim. The frozen 0.3.1rc1 checkpoint remains separate.

## Implemented and checked locally

- [x] Native model-backed PivotTable authoring: row/column/page fields, explicit
  measures, multiple measures, tabular/compact/outline layouts, grand totals.
- [x] Native PivotCharts: column/bar/line, linked report identity, generated
  series after Excel refresh, save/reopen.
- [x] Six-report/two-chart native matrix with exact totals; filtered chart
  default-member metadata corrected and tested.
- [x] Selected-member filters and ascending/descending label-sort APIs with
  validation against current model values.
- [x] Chart style, legend, axes, number format, value labels, and color APIs.
- [x] Report inspection and guarded edits/removal for recognized saved layouts.
- [x] Real OLAP slicer and timeline authoring with dedicated caches and views.
- [x] Controls connected to two PivotTables: initial totals `100`/`100`, Alpha
  slicer selection `40`/`40`, January timeline selection `30`/`30`, clear
  `100`/`100`, save/reopen `100`/`100`; each control had one view and two report
  connections. Evidence: workspace
  `work/phase2-rest/public-connected-controls-final.xlsx.json`.
- [x] Control package tests: atomic failures, duplicate names, separate cache
  identities, correct relationships, exact extension URI spelling, existing
  chart preservation, view merging, and public validation.
- [x] Expanded executable example generates three reports, selected-member
  filtering, descending sorting, chart formatting, slicer, and timeline without
  Excel. Generated workspace artifact:
  `work/phase2-rest/full-public-example.xlsx`.

## Final integration and release gates

- [x] Native validation of the expanded executable example as one combined
  fixture, including its static East-only filter alongside connected controls.
  The full and total-only reports returned `100`, while the East-only report
  retained `30`. Actual category-button selection returned Alpha `40`, clearing
  restored `100`, and January timeline selection returned `30`. Save/reopen and
  control member lists were checked. Evidence:
  `work/phase2-rest/phase2-final-asserted.xlsx.json` and the
  root's interactive Excel checks.
- [x] Native typed slicer selections: booleans True/False returned `40`/`20`;
  January 1/2/3 dates returned `10`/`20`/`30`; integer keys 1/3 returned `10`/`30`;
  the blank member returned `20`. Evidence:
  `work/phase2-rest/typed-slicer-proven-seed-verified.json`.
- [x] Native verification of all final saved-report/control lifecycle changes
  and supported chart edits; retain precise unsupported-metadata guards.
  This covers accepted-source control edits/removal, connected reports, and
  chart formatting/type edits, with native save/reopen checks.
- [x] Confirm ordinary interactive Excel open and refresh behavior. On the
  tested build, opening and ordinary Refresh/Refresh All did not populate fresh
  report layouts. Native Excel-authored references also did not automatically
  render changed model results on open. This is a documented limitation, not
  successful refresh-on-open acceptance. The tested UI action is **Defer Layout
  Update → Update**; the COM equivalent `PivotTable.Update()` populated reports
  and linked charts. See [first-use instructions](phase2-reports.md#first-use-in-desktop-excel).
- [x] Optional `tools/update_model_reports.ps1` executed successfully against
  the complete example: report counts, populated-cell counts, and chart-series
  counts survived save/reopen. Output:
  `work/phase2-rest/phase2-final-ready.xlsx`; machine-readable evidence:
  `work/phase2-rest/phase2-final-ready.json`. Regenerated from revision
  `23e2946`, the final example passed all control total assertions again.
- [x] Reject unprocessed or stale calculated-column member inventories for
  member filters, slicers, and timelines. Preserve valid processed readback;
  do not expose temporary blank caches as actual source members.
- [x] Final integrated Windows suite: **712 passed**, no failures or skips,
  461.66 seconds, runtime revision `23e2946`. Wheel/source builds and clean
  installed API/CLI smoke tests pass; all 61 runtime payload files match source.
  Archive audits exclude workspace outputs and unidentified legacy fixtures.
  Evidence: `work/phase2-rest/full-tests-final.xml` and
  `work/phase2-rest/verification-report.json`.
- [x] All nine installed-package CI jobs passed for runtime revision `23e2946`:
  Windows, Linux and macOS with Python 3.10, 3.12 and 3.14.
  [Compatibility run](https://github.com/sean-palmer-coding/excel-model-lab/actions/runs/42994636183).
- [x] Reconcile evidence paths and public documentation with supported scope.
  Version **0.4.0rc1** is a local release candidate; PyPI publication is deferred.
  The final wheel, source archive, examples and verification evidence belong in
  `checkpoints/2026-11-11-phase2-0.4.0rc1/`. The old 0.3.1rc1 artifacts remain
  separate.
- [x] Root acceptance: implemented APIs, native behavior, lifecycle operations,
  regression suite, cross-platform installed-package checks, documentation and
  distribution checks are complete. First-use layout updating and unsupported
  inputs remain explicit boundaries; automatic rendering is not claimed.

## Model-storage validation

- [x] Isolate empty-string storage failure independently of report parts.
  Eight native model-only fixtures establish that stored empty strings fail,
  while null text, bracketed text, and boolean-first columns load. The writer
  now raises `UnsupportedModelError` before storing `""`; callers can explicitly
  choose `None` for BLANK. No silent normalization occurs. DAX expressions and
  reading existing models are unaffected. Atomicity and DAX regressions pass.
  Evidence: `work/phase2-rest/model-isolation-native.json` and
  `tests/test_empty_text_validation.py`.

## Deliberate boundaries

No headless DAX/Excel calculation, fabricated cached results, measure-value
sorting, arbitrary MDX selection expressions, or general worksheet editing.
Unknown or ambiguous native metadata fails closed during regeneration.
Moving existing PivotTables and geometry edits on unsupported native anchor
shapes remain guarded. Inspecting XML alone does not certify a working control;
native acceptance must check object counts, filtering behavior, and save/reopen.
