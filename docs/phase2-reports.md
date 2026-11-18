# Phase 2: model-backed reports

The source-checkout API creates and edits native Excel PivotTables, PivotCharts,
slicers, and timelines backed by the workbook's embedded Data Model. This work
is separate from the frozen 0.3.1rc1 model-authoring checkpoint. Phase 2
acceptance is recorded in [the checklist](phase2-acceptance.md); it is not a
published-release or universal Excel-compatibility claim.

Authoring requires no Excel installation. Excel calculates and displays the
reports when their layouts are updated. The library does not evaluate DAX or
invent worksheet/chart results.

Member filters, slicers, and timelines need known source values. For calculated
columns, Excel must first calculate and save the model; reopen that saved file
before authoring these selections or controls. Unprocessed columns and calculated
caches affected by staged model changes raise a Python validation error instead
of exposing placeholder blanks or stale members.

Stored zero-length text (`""`) is currently rejected with an explicit error
because native Excel cannot load that storage produced by this writer. Choose
`None` explicitly for BLANK or supply nonempty text; the library does not
silently convert values. Existing-model reading and DAX expressions that return
`""` remain supported independently of this stored-data limitation.

## Run the complete example

```bash
python examples/model_reports.py model-reports.xlsx
```

The [example](../examples/model_reports.py) creates four Sales rows with dates,
a `Total Amount` measure, a descending-category report, a total-only report,
an East-only report, a formatted column chart, and a category slicer and date
timeline connected to the first two reports.

After the native layout update below, the full report has Gamma `40`, Beta `20`, Alpha `40`,
East `30`, West `70`, and grand total `100`. The separate East-only report totals
`30`. Selecting Alpha in the slicer changes both connected reports to `40`;
clearing it and selecting January 2026 in the timeline changes both to `30`.
Clearing both controls restores `100`. The East-only report is not connected
to these controls and retains its explicit East selection.

### First use in desktop Excel

New report cells and charts can initially appear blank. They contain report
definitions, not cached zeroes. On the tested Excel build, opening the file,
Refresh All, and ordinary PivotTable Refresh did **not** populate those layouts.
The cache's refresh-on-open flag alone did not solve this, including in an
Excel-authored reference workbook. Do not assume opening a workbook will update
previously saved report results after a headless model edit either.

To populate a report using the Excel interface:

1. Select its anchor cell, normally **A3**, on the report sheet.
2. Show the **PivotTable Fields** pane if it is hidden.
3. Check **Defer Layout Update** at the bottom of the pane, then click **Update**.
4. Repeat for the other report sheets and save the workbook.

This layout-update action populated the reports and their linked chart in native
testing. For the complete example, check grand totals `100`, `100`, and `30`
on the three report sheets before saving. Existing filters can change totals.

### Optional Windows helper

With desktop Excel installed, update all embedded-model report layouts in a
separate output workbook:

```powershell
.\tools\update_model_reports.ps1 -Source .\model-reports.xlsx -Output .\model-reports-ready.xlsx
```

The helper creates its own hidden Excel instance, disables macros and link
updates, initializes the embedded model, and calls `PivotTable.Update()` for
its reports. It does not call Refresh All or explicitly refresh upstream source
connections. It leaves the source unchanged, refuses to overwrite an existing
output, and saves/reopens a staged copy before committing the output. Its JSON
summary checks that report counts, populated-cell counts, and chart-series counts
survive reopening; it does not independently verify DAX results.

The essential COM sequence is `workbook.Model.Initialize()`, followed by
`p.Update()` for each embedded-model PivotTable. Here `p` is an Excel COM object,
not a library object. `Update()` alone populated the tested layouts; adding
`RefreshTable()` and cache refresh was unnecessary. Microsoft separately notes
that automatic PivotTable refresh is not invoked by `Workbooks.Open`.
[RefreshOnFileOpen documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache.refreshonfileopen)

Dedicated control caches are managed by their controls; do not indiscriminately
refresh every cache in the workbook. Use `ClearDateFilter()` to clear an Excel
timeline; the tested `ClearManualFilter()` call did not clear its date selection.

## Create reports

```python
from excel_model_lab import (
    WorkbookReport, PivotTableSpec, PivotChartSpec,
    PivotMemberFilter, PivotSort, SlicerSpec, TimelineSpec,
)

report = WorkbookReport("sales-model.xlsx")  # Also accepts a live WorkbookModel.
report.add_pivot_table(PivotTableSpec(
    name="SalesPivot", sheet="Sales Report",
    rows=(("Sales", "Category"),),
    filters=(("Sales", "Region"),),
    measures=("Total Amount",),
    member_filters=(PivotMemberFilter("Sales", "Region", ("East",)),),
    sorts=(PivotSort("Sales", "Category", "descending"),),
))
report.add_pivot_chart(PivotChartSpec(
    name="SalesChart", pivot_table="SalesPivot", title="East sales",
    chart_type="column", legend_position="bottom", number_format="#,##0",
    show_data_labels=True, value_axis_title="Amount",
))
report.save("sales-report.xlsx")
```

All add methods accept either a specification or its fields as keyword
arguments, and return the report for chaining. `list_reports()` lists staged
specifications. `save(path)` validates the live model again and atomically
replaces the destination. `to_bytes()` returns workbook bytes.

| PivotTableSpec fields | Behavior |
| --- | --- |
| `name`, `sheet`, `cell` | Identifier, new worksheet, and anchor (`A3`) |
| `rows`, `columns`, `filters` | Sequences of `(table, column)` pairs |
| `measures` | Explicit model measure names; numeric columns do not implicitly create measures |
| `layout` | `tabular`, `compact`, or `outline` |
| `row_grand_totals`, `column_grand_totals` | Both default to `True` |
| `member_filters` | `PivotMemberFilter(table, column, members)` on an axis field |
| `sorts` | `PivotSort(table, column, order)`; ascending or descending label order |

Selected members must exist in the current model. Supported selection values
are text, integers, booleans, timezone-naive whole-second dates/datetimes, and
`None` for blank. Unsupported scalar types, empty selections, duplicate members,
and arbitrary MDX expressions are rejected. Label sorting is distinct from
sorting by a calculated measure; measure-value sorting is not exposed.

| PivotChartSpec fields | Behavior |
| --- | --- |
| `name`, `pivot_table`, `sheet`, `cell` | Real PivotTable binding; default sheet is the pivot sheet and anchor is `G3` |
| `chart_type` | `column`, `bar`, or `line` |
| `width`, `height` | Positive integer pixels, default `640` by `360` |
| `title`, `style` | Literal title and built-in chart style |
| `show_legend`, `legend_position` | Visibility and `left`, `right`, `top`, `bottom`, or `top_right` placement |
| `category_axis_title`, `value_axis_title` | Axis titles |
| `number_format`, `show_data_labels` | Value-axis number format and value labels |
| `background_color`, `plot_background_color` | Optional six-digit RGB colors |

Charts bind to PivotTables rather than fixed cell ranges. Excel generates their
series after refresh. Existing chart edits retain cached data while changing
supported formatting; a later native refresh can regenerate that data.

## Slicers and timelines

```python
report.add_slicer(SlicerSpec(
    name="Categories", pivot_table="SalesPivot", field=("Sales", "Category"),
    cell="J3", caption="Category", columns=1,
    connected_pivots=("TotalsPivot",),
))
report.add_timeline(TimelineSpec(
    name="Dates", pivot_table="SalesPivot", field=("Sales", "OrderDate"),
    cell="J20", level="months", connected_pivots=("TotalsPivot",),
))
```

The referenced PivotTables must exist or be staged before the controls are
added. `connected_pivots` lists additional reports beyond `pivot_table`; all
must use the same model connection. Controls initially select all members.
Slicers default to `192 × 240` pixels and timelines to `480 × 140` pixels.
Both accept `sheet`, `caption`, `cell`, `width`, and `height`. Timeline levels
are `years`, `quarters`, `months`, and `days`; timelines require a date column
with at least one nonblank date. Bounds before year 9999 are supported.

These are genuine OLAP controls with dedicated caches and drawing bindings,
not pictures or decorative shapes. Slicer item lists and native report results
are populated by Excel. Arbitrary initial slicer/timeline selections are not
exposed in these specifications; use PivotTable member filters for an authored
initial report selection.

## Inspect, edit, and remove saved objects

`WorkbookReport.inspect_reports()` returns saved PivotTable/PivotChart records,
including reconstructed specifications where supported and explicit unsupported
metadata indicators. `inspect_controls()` returns saved control bindings,
source fields, connected reports, captions, and package locations. Inspection
reports package metadata; it does not evaluate the model or certify Excel
acceptance.

Use `update_pivot_table(name, **changes)`, `update_pivot_chart(name, **changes)`,
`update_slicer(name, **changes)`, and `update_timeline(name, **changes)` to stage
supported changes, then `save()` to commit them. `remove_report(name)` removes
a staged or supported saved object. Remove dependent charts and controls before
removing their PivotTable. Renaming a PivotTable updates recognized chart and
control bindings.

Saved PivotTable name, layout, and grand-total edits preserve unrelated XML.
Axis, measure, selected-member, and sort changes regenerate only recognized
self-contained layouts; shared caches, unresolved control bindings, or unknown
metadata cause an error rather than silent loss. Moving an existing PivotTable
is not supported. Removing a report leaves its worksheet and already displayed
cells in place.

Supported control edits include name, caption, slicer column count, timeline
time level, and connected reports. Geometry changes require a recognized
one-cell anchor; Excel-authored two-cell anchors are rejected rather than
rewritten approximately. A control cache shared by multiple views cannot have
its connections edited independently. Existing selection state is preserved
when changing a timeline's displayed time level.

Combined charts, unsupported axis arrangements, unknown type-specific series
formatting, and complex title markup can prevent a requested chart edit.
These guards are intentional boundaries, not promises to round-trip arbitrary
Excel report features through regeneration.

## Validation and native evidence

Names, fields, measures, axes, types, dimensions, selections, and connections
are validated before committing writes. Invalid specifications raise
`ModelValidationError`; incompatible argument shapes may raise `TypeError`.
Unrelated workbook members and model bytes are preserved by report packaging.
The API is not a general worksheet editor or a headless calculation engine.

Local desktop Excel evidence includes:

- Six PivotTables covering multiple measures, page filters, compact/outline
  layouts, total-only and column-only reports, plus linked bar/line charts;
  exact values survived save/reopen.
- Column, bar, and line charts; selected members and label sorting; chart
  formatting and supported saved-report edits (see the acceptance checklist
  for final integration status).
- Connected slicer/timeline interaction on two reports: both initially `100`,
  Alpha selection `40`, January selection `30`, clears `100`, and reopen `100`.
  Both controls remained real Excel objects with one view and two connections.
  Workspace evidence: `work/phase2-rest/public-connected-controls-final.xlsx.json`.
- Appending reports while retaining existing ones and refreshing changed model
  amounts from total `100` to `200`; earlier evidence is recorded in
  [phase2-native-evidence.json](phase2-native-evidence.json).
- The combined executable example passed native control interaction and
  save/reopen (`work/phase2-rest/full-public-example-verified-v2.xlsx.json`).
  Appending and editing reports retained working controls
  (`append-and-edit-verified.xlsx.json`); edited axes produced descending IDs
  `3, 1` with the East/North selection and total `40`
  (`edited-axes-verified.xlsx.json`, in the same workspace evidence directory).

These observations apply to the tested fixtures and local Excel installation.
The final cross-platform regression/package checks passed; interactive
refresh-on-open remains a documented limitation in the
[acceptance record](phase2-acceptance.md).

On Windows, [verify_model_reports.ps1](../tools/verify_model_reports.ps1) provides
a bounded native report-refresh/save/reopen harness. Run it only on fixtures
whose connections you intend to refresh; it uses one test workbook at a time.

[verify_model_report_controls.ps1](../tools/verify_model_report_controls.ps1)
also checks actual control objects, connections, selection effects, clearing,
and save/reopen. Its `-Example` mode asserts the combined example's totals:

```powershell
./tools/verify_model_report_controls.ps1 -Source ./model-reports.xlsx `
    -Output ./model-reports-verified.xlsx -Example
```

The output must not already exist. Other fixtures can supply `CategoryTable`,
`CategoryColumn`, `CategoryMember`, `StartDate`, `EndDate`, and expected-total
hashtables through `ExpectedInitial`, `ExpectedSlicer`, and `ExpectedTimeline`.
Total assertions assume a single measure with grand totals enabled. This tool
requires installed desktop Excel; the Python authoring API does not.
