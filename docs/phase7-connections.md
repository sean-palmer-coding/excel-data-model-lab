# External semantic-model connections

The new API writes external Analysis Services connection definitions without
Excel installed. It does not replace the embedded model, import tables, create
remote models, or execute a query during authoring. Existing native
`OleDbSource` and `LinkedTableSource` bindings instead refresh stored data in the
embedded model. Python `SqlSource`, `CsvSource` and `ParquetSource` refresh data
through Python and do not create workbook refresh connections.

## Author a connection

```python
from excel_model_lab.external_connections import (
    ExternalModelConnection, add_external_connection,
    inspect_external_connections,
)

connection = ExternalModelConnection(
    name="Live Sales", server=r"localhost\ASTabular", database="Sales",
    cube="Model", authentication="integrated",
)
add_external_connection("dashboard.xlsx", "connected.xlsx", connection)
print(inspect_external_connections("connected.xlsx"))
```

The API adds an OLE DB OLAP connection that Excel CUBE formulas can reference
by its name. `CUBEVALUE("Live Sales","[Measures].[Total Amount]")` requests a
measure from the remote model when Excel connects. This API currently authors
the connection only; existing `WorkbookReport` specs still target the embedded
model and cannot be silently redirected to the external model.

Names are unique case-insensitively. `ThisWorkbookDataModel` is reserved.
Existing connections and unrelated ZIP entries are retained; output is staged
and committed atomically. Digitally signed packages are rejected. The API does
not replace or remove existing connections, because those operations require
dependency analysis of formulas, reports and other consumers.

No password or token field exists in the workbook definition. Integrated
security targets Windows-authenticated SSAS. For cloud endpoints use
`authentication="interactive"`, which delegates login to Excel and its provider;
this is not a guarantee of unattended cloud authentication. Excel needs MSOLAP,
network access and permissions to the actual database. Refresh-on-open is off.
Connection strings quote their values to prevent property injection.

## Optional live query adapter

```python
from excel_model_lab.external_connections import AdomdQueryClient

client = AdomdQueryClient(connection, assembly_path=r"C:\libs\Microsoft.AnalysisServices.AdomdClient.dll")
result = client.query('EVALUATE ROW("Total", [Total Amount])', timeout=30, max_rows=1000)
print(result.columns, result.rows)
frame = result.to_dataframe()
```

Install `pythonnet` and Microsoft's ADOMD.NET client separately. Importing the
core writer or authoring connections never loads .NET. The adapter accepts DAX
queries beginning with `EVALUATE` or `DEFINE`, bounds materialized rows, and
raises instead of silently truncating. It accepts exactly one result set;
multiple `EVALUATE` result sets raise an explicit error instead of dropping
later results. Provider errors retain their diagnostic
cause without copying provider text into the public error message. Reader,
command and connection resources are disposed even when a query fails.

For runtime tokens or another authentication strategy, provide
`connection_factory=lambda: configured_adomd_connection`. Each call must return
a fresh connection owned by the adapter; the caller can configure its in-memory
`AccessToken`. No runtime credentials are written to the workbook. Result cells
preserve provider values, with database null mapped to `None`; this adapter does
not promise conversion of every .NET-specific type. `to_dataframe()` uses object
columns so nullable large integers, exact decimals and `None` are not silently
converted to floating point. Callers can explicitly select pandas dtypes after
checking their data's range and precision requirements.

## Verified local SSAS connection workflow

Automated tests cover package preservation, relationship/content-type wiring,
alternate connection-part names, names and credential policy, atomic rejection,
row bounds, null results, provider failures and disposal. Controlled provider
tests also cover connection-string assignment failure before ownership passes
to the query method, and exact nullable integer/decimal DataFrame conversion.

Native Excel metadata acceptance passed on the current Windows installation:
`work/phase7/external-metadata.xlsx` opens with the existing embedded model and
the new `Live Sales` connection recognized as COM type 1 (OLE DB), OLAP true,
command type 1, and cube command `Model`. Evidence is
`work/phase7/native-metadata-result.json`. XML connection type 5 and Excel COM
connection type 1 use different enumerations. No remote query or refresh was
attempted in this metadata-only check.

The previous installation blocker was resolved after the user installed a local
SQL Server Analysis Services Tabular instance. On November 17, 2026, the adapter
and Excel were exercised against `localhost\ExcelModelLab`, SSAS **17.0.25.218**,
using Windows integrated authentication and Microsoft's ADOMD.NET **19.118.3**
client (`net472`). This is an Analysis Services endpoint, not the relational SQL
Server engine.

A dedicated synthetic database, `ExcelModelLab_Phase7_Test_20261117`, contains
Sales amounts 10, 20, 30 and labels Alpha, Beta, Gamma. Native acceptance recorded:

- A live DAX measure query returned total **60**, and projected label/amount rows
  matched the independently expected three rows.
- BLANK mapped to `None`; an empty result returned no rows with column metadata.
- Row-limit overflow, additional result sets and invalid DAX raised explicit
  errors. A subsequent query succeeded after those error paths.
- A nullable large-integer query preserved **9007199254740993** exactly in
  `ExternalQueryResult.rows`, followed by `None`.
- An authored external CUBE workbook displayed **60**. After changing the server
  data and reprocessing, the server returned **120**, while the workbook still
  displayed its saved **60** until deliberate refresh. Excel then displayed
  **120**, saved, and reopened with **120**. The workbook did not receive a new
  embedded model to obtain that changed result.

The final repeatable run is recorded in
[portable native evidence](phase7-live-native-evidence.json). The executable
procedure is `scripts/verify_external_semantic_endpoint.py`; it creates a new
synthetic database and output directory and refuses to overwrite existing ones.

To repeat this workflow on an authorized tabular endpoint:

1. Deploy a small Sales model with Amount values 10, 20, 30 and measure
   `Total Amount := SUM(Sales[Amount])`.
2. Generate `examples/external_semantic_connection.py` with the real server and
   database. Confirm the workbook opens without repair and the Live Sales
   connection remains external OLAP with no embedded credentials.
3. Authenticate, calculate its CUBE formulas, assert Total Amount = 60, save and
   reopen. Change server data and reprocess; deliberate refresh must retrieve
   the changed result without copying a new model into the workbook.
4. Run `AdomdQueryClient.query` against the same endpoint and compare the total,
   blanks, labels and row limits with independently expected results.

Power BI/Azure Analysis Services authentication, external PivotTable generation,
general DirectQuery configuration, and remote model deployment remain outside
this acceptance. The verified workflow is connection authoring, live DAX querying
and Excel CUBE calculation/refresh against the tested local SSAS Tabular endpoint;
it is not a certification of every provider, deployment or Excel version.

## Primary references

- [Microsoft connection properties](https://learn.microsoft.com/en-us/analysis-services/instances/connection-string-properties-analysis-services)
- [SpreadsheetML connection definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.connection)
- [Database properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.databaseproperties)
- [OLAP properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.olapproperties)
- [Analysis Services installation](https://learn.microsoft.com/en-us/analysis-services/instances/install-windows/install-analysis-services)

The final repeatable run also verified exact nullable large integers through
`to_dataframe()`. See [current live acceptance](phase7-live-acceptance.md) and
[its native evidence](phase7-live-native-evidence.json).
