# Excel Data Model Lab

Python imports remain `excel_model_lab` and `excelmodel`. The distribution was
renamed from `excel-model-lab`; uninstall that older distribution before switching
to avoid overlapping package files.

Current release candidate: **0.5.0rc1**.

Install with `python -m pip install excel-data-model-lab==0.5.0rc1`.

Excel Data Model Lab is a headless Python writer for Excel's embedded Power Pivot /
Data Model. The writer does not require Excel, COM, Windows, or a remote model.
Desktop Excel is used separately to verify compatibility and by the optional
`desktop` extra for live workbook interaction.

The release includes model authoring and reading, model-backed PivotTables and
charts, workbook composition and CUBE helpers, bounded DAX evaluation, external
semantic connection/query adapters, and optional desktop/agent workflows.
Release verification: **864 passed, 50 expected skips** on local hardware; see
[publication acceptance](docs/publication-acceptance.md).
See [review acceptance](docs/review-fixes-acceptance.md) for local test and native
Excel evidence, and [Phase 7 acceptance](docs/phase7-live-acceptance.md) for SSAS
query and Excel refresh verification. Cloud authentication and external-model
PivotTable generation remain outside the verified scope. Historical CI results
in milestone reports apply to those checkpoints; this release is tested locally.

See [STATUS.md](STATUS.md) for current implementation limits and
[SPEC_AND_CLI.md](SPEC_AND_CLI.md) for JSON authoring and command-line validation.

The current source adds native Data Model PivotTables, PivotCharts, slicers and
date timelines, including controls connected to multiple reports. It supports
selected-member filters, label sorting, chart formatting, inspection, and
guarded edits/removal of saved reports and controls. Excel calculates and renders
the reports; Python does not require Excel to author them. Native tests verify
filtering and save/reopen, including connected controls and the combined example.
See the [report API guide](docs/phase2-reports.md),
[runnable example](examples/model_reports.py), and
[Phase 2 acceptance record](docs/phase2-acceptance.md) for verification evidence
and precise limits. On the tested Excel build, opening a fresh report or using
ordinary Refresh/Refresh All leaves its layout blank. Select the report's anchor
cell, check **Defer Layout Update** in **PivotTable Fields**, then click **Update**.
The optional [Windows helper](tools/update_model_reports.ps1) performs the tested
`PivotTable.Update()` action for all embedded-model reports and saves a separate
output. See the guide's first-use instructions; automatic rendering on open is
not guaranteed.

## Install

```bash
python -m pip install .
```

## Python API

```python
from excelmodel import WorkbookModel

workbook = WorkbookModel.from_template("template.xlsx")
workbook.replace_tables({
    "Claims": claims_df,
    "Providers": providers_df,
})
workbook.add_relationship(
    "Claims", "ProviderID", "Providers", "ProviderID"
)
workbook.add_measure(
    "Claims",
    "Total Charges",
    "SUM('Claims'[ChargeAmount])",
    format_string="$#,##0.00",
)
workbook.save("report.xlsx")
```

Create a model without supplying a template:

```python
from excel_model_lab import WorkbookModel

model = WorkbookModel.create({"Products": products_df, "Sales": sales_df})
model.add_column("Products", "Rank", [3, 2, 1])
model.set_sort_column("Products", "Category", "Rank")
model.add_hierarchy("Products", "Catalog", ["Category", "ProductID"])
model.add_relationship("Sales", "ProductID", "Products", "ProductID")
model.save("created.xlsx")
```

The example expects three product rows. Pass `workbook="ordinary.xlsx"` to
`create` to preserve an existing workbook's worksheets while inserting a model.
`add_table` accepts an independent dataframe. `rename_table`, `remove_table`,
`rename_column`, `remove_column`, `change_column_type`, `configure_table` and
`configure_column` cover schema and display edits. Dependency guards reject
unsupported changes involving existing calculations or presentation objects.

Hierarchies support `update_hierarchy`, `configure_hierarchy` and
`remove_hierarchy`. Pass the displayed hierarchy name, even when a native
workbook uses a different internal identifier. Relationships support `set_relationship_active` and
`remove_relationship`. `set_sort_column(..., by=None)` restores default ordering.
DAX calculation-engine work is separate; existing measure-writing support remains.

`to_bytes()` is intended for web responses:

```python
payload = workbook.to_bytes()
```

Every save is independently reopened with PBIXRay by default. Requested values,
row counts, measures, existing relationships, and new relationships are checked
before the generated path is returned. Pass `validate=False` only when an
application has its own validation boundary.

## Model as code

```yaml
tables:
  Claims:
    source: data/claims.csv
  Providers:
    source: data/providers.parquet
relationships:
  - from: Claims.ProviderID
    to: Providers.ProviderID
measures:
  - table: Claims
    name: Total Charges
    expression: SUM('Claims'[ChargeAmount])
    format: "$#,##0.00"
    display_folder: Finance
```

Build it with paths resolved relative to the specification:

```bash
excelmodel build template.xlsx report.xlsx model.yaml
```

Create from a specification with no template:

```bash
excelmodel create report.xlsx model.yaml
excelmodel create report.xlsx model.yaml --workbook ordinary.xlsx
```

`build` replaces named existing tables and adds new table names. CSV definitions
can include `read_options` forwarded to Pandas (for example `dtype` or
`parse_dates`). Install `excel-data-model-lab[parquet]` for Parquet input. Table
definitions can also contain `hidden`, `description`, `columns` (per-column
`hidden`, `description`, `format`, `sort_by`) and `hierarchies` (`name`, `levels`,
optional `hidden`, `description`, `level_names`). Paths are relative to the spec.

## Supported model data

### KPIs and calculation-area visibility

Create a three-state, higher-is-better KPI using an existing target measure:

```python
model.add_kpi("Sales", "Total Amount", target_measure="Target Amount",
              lower=0.4, upper=0.8)
print(model.list_kpis())
model.update_kpi("Sales", "Total Amount", target_measure="Target Amount",
                 lower=0.5, upper=0.95)
model.remove_kpi("Sales", "Total Amount")  # Retains both user measures.
model.update_measure("Sales", "Target Amount", expression="120", format_string="0")
# Independent measures can also be removed with model.remove_measure(table, name).
```

The example uses percentage thresholds. KPIs also support fixed targets,
absolute thresholds, ascending/descending and centered modes, and multiple native
icon sets. See [the KPI guide](docs/CALCULATED_COLUMNS_AND_KPIS.md) for supported
combinations. The API writes native metadata and helper formulas; it does not
evaluate DAX. Newly visible measures receive separate
calculation-area rows; existing saved measure positions are retained.
Measure removal rejects calculation and existing PivotTable dependencies;
remove a base measure's KPI before editing that base measure. Measure names
must be unique across the model, ignoring case.

- Existing tables can be replaced individually or atomically as a group.
- Row counts can shrink or grow.
- Integer, real, Currency, Date, nullable numeric, and dictionary-encoded string
  columns are supported for the tested template encodings.
- Pure bit-packed, hybrid RLE, and multi-segment IDF data are regenerated.
- String and numeric dictionaries and numeric hash indexes are regenerated.
- DAX measures can include format strings, display folders, and descriptions.
- Active and inactive many-to-one relationships can be added.
- Untouched XLSX members and unknown ABF virtual files are preserved.
- Output remains a normal, self-contained `.xlsx` package.

Use schema inspection before selecting a template:

```bash
excelmodel schema template.xlsx
excelmodel inspect template.xlsx --semantic
```

## Compatibility boundaries

A full table replacement must provide every non-engine column in that table.
Creating new models uses a bundled synthetic seed; no caller-supplied template
is required. Existing workbook presentation parts are preserved, and unsupported
schema changes with presentation dependencies are rejected.

The current authoring backend has limits: new stored column identifiers are
restricted. Text dictionary ordering supports pinned en-US/en-GB collation;
unsupported locales and codepoints are rejected. Numeric dictionaries support
values beyond the arithmetic encoding's signed 32-bit ID span. The supported
authoring range is `-9223372036854775807` through `9223372036854775806` for integers
and `-922337203685477.5807` through `922337203685477.5806` for Currency, inclusive.
The immediately adjacent signed64 endpoints are rejected after native aggregate
failures; existing endpoint values remain readable. Isolated native DAX queries
returned exact aggregates for the supported boundaries. New datetime columns support
timezone-naive values through whole seconds. Subsecond values are rejected.
The independent reader's floating-point date artifacts are normalized only to
the whole-second grid declared by value-encoding metadata. Empty tables use native
zero-row storage; use `replace_table` to populate all their columns together.
The integrated API's create/populate/clear outputs pass desktop Power Pivot
acceptance. Unrecognized storage layouts are rejected rather than rewritten
speculatively. Combined native load, calculation, filtering, hierarchy and
reopen checks pass, including tiny dictionaries and nullable numeric lifecycles.

External data connections are preserved as package content. The API can author
Excel-native linked-table and OLE DB source bindings, but does not execute Excel's
refresh engine. Linked-table refresh has passed a native test; SQL Server two-source refresh, rebind, detach and save/reopen have passed native
checks using authorized credentials supplied through Excel data APIs. The public
writer rejects embedded passwords; it does not manage Excel session credentials. Enabling an old
workbook's external refresh can replace stored values from its original source.

### Refreshing stored data through Python

Use `refresh_sources` to read CSV, Parquet, or a parameterized SQL query into
existing tables. All sources are read and checked before any table update is
applied. Each source must supply every stored column in its target table.

```python
from excel_model_lab import WorkbookModel, CsvSource, ParquetSource, SqlSource

model = WorkbookModel("model.xlsx")
model.refresh_sources({
    "Sales": CsvSource("sales.csv"),
    "Products": ParquetSource("products.parquet"),
})
print(model.tables)
print(model.get_table("Sales"))
model.save("refreshed.xlsx")

# With a caller-owned SQLite or SQLAlchemy connection:
# model.refresh_sources({"Sales": SqlSource(connection,
#     "SELECT * FROM sales WHERE year = :year", params={"year": 2026})})
```

Parquet requires the optional `parquet` dependency. Readers accept
`read_options` for pandas options such as CSV dtypes. SQL connections remain
owned by the caller. Source paths, connection credentials, and query bindings
are not saved as workbook refresh connections. This method reads stored data;
it does not evaluate DAX.

Workbook presentation layout authoring, a DAX engine, universal historical
VertiPaq/encrypted-workbook compatibility, and `.xls`/`.xlsb`/VBA authoring are
outside this work. Unsupported encodings fail before committing the destination.
Independent decoder validation does not replace desktop Excel acceptance.

## Development

```bash
python -m pip install -e .
python -m pytest -q
python -m build
```

Low-level laboratory commands remain available under `excelmodel --help` for ABF
inventory, extraction, cloning, diffing, and controlled IDF mutation.

### Explicit relationship cardinality

`add_relationship(..., cardinality="many-to-one")` retains the existing default.
Use `one-to-many` to supply the lookup end first; the API normalizes the stored
endpoints to many-to-one. Removal/activation use these normalized endpoints.
Use `one-to-one` to require unique keys at both ends and write One/One metadata.
Explicit One/One metadata passed native save/reopen acceptance.

Invalid cardinality, duplicate keys, type mismatches and invalid active paths
raise `RelationshipValidationError`, a `ValueError` subclass. Duplicate errors
include the table, column, value and one-based data row positions. Missing
columns retain `KeyError`. Checks run at definition and again before saving
changed relationship keys, even when `save(validate=False)` is used. Activation
changes retain cardinality. Validation does not require every foreign key to
match: unmatched rows retain Excel's blank-member behavior.

Blank lookup keys remain an implementation limitation (`UnsupportedModelError`),
not a claim that Excel rejects every blank. Text keys use Windows en-US linguistic comparison when available; other locales
and equivalent distinct spellings within a dictionary remain guarded.

## Read-only inspection and streaming

The reader does not initialize the writer or require write-compatible encodings.
It delegates to the pinned PBIXRay 0.15.5 public API. It supports the formats that
PBIXRay supports (XLSX/PBIX/ABF); this does not add PBIX/ABF writing. Excel fixture
parity and resource lifecycle are tested here; broad format coverage comes from
PBIXRay rather than a new independent compatibility certification.

```python
from excel_model_lab import WorkbookModel

with WorkbookModel.open("report.xlsx", on_disk=True) as reader:
    print(reader.tables)
    print(reader.schema)
    print(reader.relationships)
    print(reader.dax_measures)
    print(reader.mashup_queries)
    for chunk in reader.iter_table("Sales", columns=["Amount"], chunk_size=10000):
        process(chunk)
```

`ModelReader` is also directly importable. All public PBIXRay metadata properties
are exposed, including `connections`, `data_mashup`, `perspectives`, `statistics`
and the `tmschema_*` properties. Availability and empty results depend on the
source format. Processed model values normalize declared/inferred types, preserving Currency as
Decimal and calculated dates, booleans and integers. Upstream format limitations remain. Keep the reader open while consuming
an iterator. Closing it releases resources; materialized DataFrames remain usable.

For editing, use `WorkbookModel.open("report.xlsx", read_only=False)` or the
existing constructor. Readers show the saved source, not another editor's staged
changes. Writer `get_table()` continues to include its staged data changes.

## Preflight validation

```python
model = WorkbookModel.open("report.xlsx", read_only=False)
report = model.validate()
for issue in report.issues:
    print(issue.severity, issue.code, issue.message)
report.raise_for_errors()
```

Current checks cover relationship key uniqueness, type compatibility, duplicate
relationships, self-links and ambiguous active paths, plus warnings for columns
outside the writer's supported encodings. Pending measure-name conflicts are also
reported before saving. This is not complete formula/dependency
validation, DAX evaluation or proof of native Excel compatibility. Existing
save-time checks remain in place. Report validity means no errors in these checks.


## Unicode and locale contract

String dictionary and relationship comparison supports pinned en-US and en-GB
collation. Windows uses NLS; the packaged portable tables provide the supported
comparison behavior on Linux, including tested Japanese context and Jamo cases.
Unsupported locales or codepoints fail explicitly. Actual macOS validation is
still pending. Equivalent distinct dictionary spellings are rejected rather than
silently rewritten; matching keys in different tables may compare equal.
Embedded NUL characters are rejected.

Native Unicode table rows and relationship filtering passed save/reopen checks.
See [STATUS.md](STATUS.md) and the native acceptance ledger for tested grouping,
dictionary layouts and platform boundaries.

## Calculated columns and date keys

```python
model.add_calculated_column("Sales", "Double Amount", "[Amount] * 2",
                            data_type="currency")
model.update_calculated_column("Sales", "Double Amount", "[Amount] * 3",
                               data_type="currency")
model.remove_calculated_column("Sales", "Double Amount")
```

Explicit types are `integer`, `decimal`, `currency`, `datetime`, `boolean` and
`text`. Definitions are stored with unprocessed caches; Excel performs DAX
calculation. The library does not invent calculated results. All six explicit
types and calculated resizing/empty-table lifecycles passed native processing
and save/reopen checks. Calculated relationships and stored-column sorting by calculated keys passed
native processing/save/reopen, including changed inputs and dependent hierarchy
checks. Final combined regression and local artifact verification passed; see
[STATUS.md](STATUS.md) for the recorded results.

`mark_date_table(table, column)` requires a nonempty stored datetime column with
unique, nonblank, contiguous whole dates. `unmark_date_table(table)` restores the
internal row key. Stored date marking passed native processing/save/reopen.
Calculated date keys remain explicitly unsupported: a processed same-table
computed date key failed native processing with a circular dependency. This is
a documented limitation, not a claim that every DAX expression is forbidden by
all Excel versions. Pending values cannot establish uniqueness or contiguity.
