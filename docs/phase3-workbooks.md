# Complete analytical workbooks

Phase 3 composes XlsxWriter worksheet content with the embedded-model and report
APIs. Authoring is headless. Excel calculates worksheet and CUBE formulas and
renders report layouts; this API does not invent cached numeric results.

Install the optional worksheet writer from the repository root:

```sh
python -m pip install '.[workbooks]'
```

## Compose a new workbook

```python
import pandas as pd
from excel_model_lab import compose_workbook, CubeCell, PivotTableSpec

tables = {'Sales': pd.DataFrame({'Category': ['Alpha', 'Beta'], 'Amount': [10, 20]})}

def model_setup(model):
    model.add_measure('Sales', 'Total Amount', 'SUM(Sales[Amount])')

def worksheets(writer, cubes):
    sheet = writer.add_worksheet('Dashboard')
    sheet.write('A1', 'Total sales')
    cubes.write(sheet, 'B1', cubes.value(cubes.measure('Total Amount')))
    cubes.write(sheet, 'B3', cubes.member('Sales', 'Category', 'Alpha'))
    cubes.write(sheet, 'B4', cubes.value(cubes.measure('Total Amount'), CubeCell('B3')))

def reports(report):
    report.add_pivot_table(PivotTableSpec('Summary', 'Report',
        rows=(('Sales', 'Category'),), measures=('Total Amount',)))

compose_workbook('dashboard.xlsx', tables, worksheets,
                 configure_model=model_setup, configure_reports=reports)
```

The model callback runs first, so worksheet formulas can validate measure and
field references. The worksheet callback receives an ordinary XlsxWriter
`Workbook` and a model-bound `CubeFormulas` helper. It can create formats, names,
tables, worksheet formulas and ordinary charts using XlsxWriter's API. Do not
close the writer from the callback. The composition function closes it before
attaching the model and adding reports.

The destination is replaced only after every stage succeeds. A callback or
validation failure leaves an existing destination untouched. Staging uses a
private directory beside the output. The output directory must already exist.
Neither a live Excel session nor an open workbook is modified by this API.

`strings_to_formulas` and `strings_to_urls` default to false: source text such as
`=SUM(A1)` remains text unless explicitly written as a formula. Advanced callers
can pass `xlsxwriter_options`; XlsxWriter's normal feature restrictions still
apply to those options. Callbacks are trusted Python application code.

## Attach a model to an already finalized workbook

```python
from excel_model_lab import attach_model

attach_model('worksheets.xlsx', 'dashboard.xlsx', tables,
             configure_model=model_setup, configure_reports=reports)
```

This accepts a closed workbook from XlsxWriter or another compatible writer.
It preserves unrelated package parts, existing worksheets, tables, charts,
formats, names and connections. New model connection IDs are allocated without
colliding with existing IDs. An existing embedded model or conflicting
`ThisWorkbookDataModel` connection is rejected; use `WorkbookModel` to edit an
existing model instead. XlsxWriter itself does not read or modify existing
workbooks. See [its FAQ](https://xlsxwriter.readthedocs.io/faq.html).

## Typed CUBE formulas

`CubeFormulas(model)` targets the embedded `ThisWorkbookDataModel` connection.

| Method | Excel function | Validation |
| --- | --- | --- |
| `member(table, column, value, caption=None)` | CUBEMEMBER | Field exists and stored member exists |
| `measure(name, caption=None)` | CUBEMEMBER | Explicit measure exists |
| `value(*members)` | CUBEVALUE | Member/set arguments; no known duplicate hierarchy |
| `set(members, caption='')` | CUBESET | Explicit unique members from one hierarchy |
| `ranked_member(set, rank, caption=None)` | CUBERANKEDMEMBER | Set argument and positive rank |
| `set_count(set)` | CUBESETCOUNT | Set argument |

Methods return immutable `CubeFormula` values. Use `str(formula)` with another
worksheet writer, or `cubes.write(worksheet, 'B2', formula, cell_format)` with
XlsxWriter. The latter writes an explicit `#N/A` unavailable result cache;
Excel replaces it during calculation. Readers that do not calculate will see
`#N/A`, not a fabricated zero. XlsxWriter's direct `write_formula` normally
caches zero unless a result is supplied; see its
[formula documentation](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

Use `CubeCell('B3')` to refer to a cell containing a CUBE member, or
`CubeCell('B7', kind='set', sheet='Dashboard')` for a CUBE set. The helper validates
A1 bounds and quotes worksheet names. Excel verifies the referenced cell's
actual type and contents. A caption string is not interchangeable with a member.

MDX identifiers escape closing brackets, and Excel string literals separately
escape double quotes. Arbitrary MDX strings and Excel expressions are not
accepted as member/set arguments. Inline text arguments above 255 characters
are rejected rather than producing Excel's CUBEMEMBER `#VALUE!`; use a custom
worksheet formula with a text cell if your application needs longer MDX.
See [Microsoft's CUBEMEMBER documentation](https://support.microsoft.com/en-us/excel/functions/cubemember-function).

Supported stored member keys are text, integer, boolean, whole-second naive
date/datetime, and explicit blank (`None`). Unprocessed or stale calculated
columns are rejected until Excel has calculated and saved them. Rank bounds
against dynamically changed sets and calculated formula results
remain Excel's responsibility. This phase does not add arbitrary set expressions,
KPI-member formula helpers, or a headless CUBE/MDX evaluator.
`member_property` raises an explicit validation error: Microsoft states that
CUBEMEMBERPROPERTY does not work with Excel Data Models edited in Power Pivot.
Use the caption returned by CUBEMEMBER instead. See
[Microsoft's documentation](https://support.microsoft.com/en-us/excel/functions/cubememberproperty-function).

## Complete acceptance example

```sh
python examples/analytical_dashboard.py dashboard.xlsx
```

The example includes formatted inputs, a named range, conditional formatting,
two ordinary charts, model measures, a PivotTable/PivotChart, a slicer, a timeline,
and all five supported CUBE function families. It writes `dashboard.expected.json`
with independently specified cell results and object counts for native tests.
Expected Dashboard results include total `100`, Alpha `40`, set count `2`, second
member `Beta`, and Alpha-plus-Beta `60`.

Excel must calculate the formulas and update the report layout. The existing
[Phase 2 first-use instructions](phase2-reports.md#first-use-in-desktop-excel)
still apply; this API does not guarantee automatic layout rendering on open.
## Native acceptance evidence

Local native acceptance passed on November 11, 2026. All five supported CUBE
function families returned the expected values; Excel retained three charts,
the model-backed report total `100`, literal source text beginning with `=`, and
the same results after save/reopen. Evidence:
`work/phase3-dashboard-native-v7.xlsx.json` in the development workspace.

The composed workbook's controls also passed actual native selection tests:
Alpha returned `40`, January returned `30`, and clearing restored `100`.
Evidence: `work/phase3-controls-native.xlsx.json`.

A 100,000-row, 100-group model passed native totals `4,999,950,000` and
`49,950,000`, followed by save/reopen. The recorded authoring measurement was
15.29 seconds, 47.36 MB peak Python allocations, and a 780,310-byte workbook.
These are measurements of this fixture and machine, not universal performance
guarantees or process-wide memory limits. Evidence:
`work/phase3-scale-100000.benchmark.json` and its native verification record.

The supported local feature gates are complete; final integrated suite,
installed-package checks and CI status are maintained in the delivery
plan/acceptance records. Native acceptance is not inferred from XML tests.
