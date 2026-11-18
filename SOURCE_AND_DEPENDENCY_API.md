# Native refresh sources and dependency-safe edits

Python refresh and native Excel refresh are separate operations. `CsvSource`,
`ParquetSource`, and `SqlSource` read data in Python through `refresh_sources`.
`LinkedTableSource` and `OleDbSource` write instructions that Excel can refresh
later. Binding a native source does not connect, run SQL, change worksheet cells,
or refresh model data. The current snapshot stays available, and authored
connections have refresh-on-open disabled.

## Existing worksheet tables

```python
from excel_model_lab import WorkbookModel, LinkedTableSource

model = WorkbookModel("input.xlsx")
model.bind_source("Sales", LinkedTableSource("SalesInput"))
model.save("linked.xlsx")
```

`SalesInput` must already be an Excel worksheet table in the workbook. Its
column names and order must match the model's stored columns. Calculated model
columns are excluded from this source schema. This API does not create or edit
worksheet tables. Changes made to the worksheet can replace model snapshot data
when the user deliberately refreshes in Excel.

`list_native_sources()` returns immutable binding descriptions. To detach a
source and retain a standalone snapshot, call `unbind_source("Sales")` before
making incompatible model schema changes. Unbinding removes query bindings and
library-owned, unreferenced OLE DB source definitions; unrelated connections stay
intact. A source-bound stored column cannot be silently deleted or renamed.

## OLE DB sources

Install the optional encryption dependency:

```text
pip install "excel-model-lab[native-sources]"
```

```python
from excel_model_lab import WorkbookModel, OleDbSource

model = WorkbookModel("input.xlsx")
model.bind_source("Sales", OleDbSource(
    connection_string=(
        "Provider=MSOLEDBSQL;Data Source=(localdb)\\MSSQLLocalDB;"
        "Initial Catalog=Example;Integrated Security=SSPI;"
    ),
    query="SELECT ID, Amount FROM dbo.Sales",
    name="Sales database",
))
model.save("query-bound.xlsx")
```

The query must return the stored model columns in their declared types. Schema
is taken from the snapshot; the library does not execute the query to verify it.
Excel needs the named OLE DB provider and database access on the computer doing
the refresh. Password/token connection-string properties are rejected. Integrated
security or credentials configured interactively in Excel are the supported
credential choices. Connection strings are not shown in `OleDbSource`'s repr.
Excel's native encryption is a reversible file-format requirement: its key is
stored inside the same workbook and is not a credential vault.

Native linked-table refresh and SQL Server two-source refresh, rebind, detach,
save/reopen, and exact Python readback passed on the recorded Windows Excel
installation. SQL authentication was supplied to the owned Excel session, not
persisted by the public source-definition API. Earlier credential failures in
the investigation log were superseded by those passing runs. Evidence is in
`docs/native-acceptance-2026-11-08.json` (production multi-source, rebound,
detached and source-preservation fixtures). This certifies the tested providers
and fixtures, not every database or Excel installation.

These are imported snapshots with refresh instructions, not live external
semantic models. For separate Analysis Services connections and the optional
query client, see [external semantic connections](docs/phase7-connections.md).

## Dependency inspection and restrictive deletion

```python
graph = model.dependency_graph()
column = ("tables", "Sales", "columns", "Amount")
for dependency in graph.dependents(column, transitive=True):
    print(dependency.source, dependency.reason)
```

A graph is an immutable snapshot containing `objects`, `edges`, and `unresolved`
references. Each object has a public path and a physical identity where available.
Paths include `("tables", table, "columns", column)`,
`("tables", table, "measures", measure)`,
`("tables", table, "hierarchies", hierarchy)`, and
`("perspectives", name)`. Edges describe DAX fields/tables, relationships,
hierarchies, sort columns, date-table keys, KPIs, perspective selections, native
source schemas and workbook report parts.

The DAX scanner understands quoted/escaped identifiers and skips comments and
string literals. It is not a complete DAX parser or evaluator. Ambiguous or
virtual-column references remain explicit. Existing report bindings are treated
conservatively until report rebinding is implemented.

Destructive methods reject dependencies without cascading. Catch `DependencyError`
or the broader `ModelValidationError`; `error.issues` contains an object path and
the dependency paths that prevented the operation. `DependencyError` also remains
an `UnsupportedModelError`/`ValueError` for existing exception handlers.

Table renaming preserves physical table identity and updates DAX references,
measure home-table/KPI annotations, perspective hierarchy selections, and Power
Pivot's table captions. Worksheet report dependencies remain restrictive.
Stored-column renaming can rebind unambiguous DAX references, but rejects
relationships, hierarchies, sort columns, perspective/source/report bindings and
ambiguous column/measure names. It rebuilds the column under a new storage name
and places it last in the table's column order.

## In-memory transactions

```python
with model.transaction() as edit:
    edit.unbind_source("Sales")
    edit.remove_column("Sales", "Unused")
    edit.add_measure("Sales", "Total", "SUM(Sales[Amount])")

model.save("edited.xlsx")
```

All edits on the yielded candidate commit together on normal context exit. An
exception discards the candidate. Save after leaving the context: files explicitly
saved inside the context are external effects and cannot be rolled back. This
context manager is not a concurrent-writer synchronization mechanism.

Source replacement preserves unrelated connections and unknown vendor connection extensions. It rejects a connection that is shared by another model table or referenced by a worksheet query/report part, because replacing that connection could redirect another consumer. Separate the shared connection before rebinding. Perspective update, rename and removal likewise reject unresolved workbook report bindings, including when the caller uses a different letter case.

Report dependency inspection also includes worksheet/named CUBE formulas, slicer caches and timeline caches. Formula text literals are not treated as CUBE function calls. These unresolved report bindings prevent structural removal/rename (including KPI removal); the library does not rewrite arbitrary MDX strings or cached report results. Source rebinding retains existing connection names and existing defined names. A requested connection rename is rejected when unresolved CUBE formulas may refer to it.

`rename_measure(table, name, new_name)` rewrites the measure declaration and unambiguous stored/pending DAX references. It preserves comments, quoted strings and calculation metadata; dependent calculated columns become unprocessed for Excel to calculate. It rejects KPI, perspective and report bindings, conflicting column/measure names, unresolved references and expressions that introduce virtual columns. Errors include the target and blocking object paths. This remains lexical rebinding, not a full DAX resolver. Remove/recreate the reported metadata binding explicitly when appropriate. Native acceptance of measure renaming is pending.
