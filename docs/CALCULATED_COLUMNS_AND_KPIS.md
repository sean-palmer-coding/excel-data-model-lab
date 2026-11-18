# Calculated columns and KPIs

The API stores DAX definitions and explicit result types. It does not evaluate
DAX. New or changed calculated columns contain BLANK cached values and report
`Unprocessed` until Excel processes them. Opening a workbook alone may not
process these columns: use Power Pivot's Calculate Now action, then save.

`examples/calculated_columns_and_kpis.py` creates a complete example. For its
three rows, Excel should calculate:

| Column | `data_type` | Expected values |
|---|---|---|
| Tripled | `integer` | 60, 90, 120 |
| Half | `decimal` | 10, 15, 20 |
| Quarter | `currency` | 5, 7.5, 10 |
| Next Date | `datetime` | January 1, 2, 3 of 2026 |
| Large | `boolean` | FALSE, TRUE, TRUE |
| Caption | `text` | Alpha!, Beta!, Gamma! |

```python
model.add_calculated_column('Sales', 'Double', '[Amount]*2', data_type='integer')
model.update_calculated_column('Sales', 'Double', '[Amount]*3', data_type='integer')
definitions = model.list_calculated_columns()
model.remove_calculated_column('Sales', 'Double')
```

The result type must be explicit. Definitions expose their expression and
processing state. `ModelReader.excel_metadata.calculated_columns` provides a
read-only metadata snapshot without initializing the writer. `get_table` and
`iter_table` read saved caches; they never execute DAX. Currency reads preserve
`Decimal` values, nullable integers retain exact values, and whole/chunked reads
apply the same native date normalization.

Editing stored inputs invalidates affected calculated caches conservatively.
Unrelated calculations proven to use only their own table's stored row values
retain their caches. Resizing a calculated table accepts a dataframe containing
only its stored columns and rebuilds BLANK calculated rows, including empty
tables and subsequent repopulation. Relationships involving calculated keys retain
their bindings and become unprocessed when their inputs change. Excel rebuilds
the indexes during calculation. Calculated date keys remain guarded because the
tested native configuration produces a circular dependency. Dependency guards
also reject removing a column that is still referenced.

Calculated-column hierarchies and sorting with calculated columns have passed
native calculation, query, save, and reopen checks. The API retains
bindings and writes explicitly unprocessed indexes for Excel to build during
calculation. Hierarchies retain captions and presentation metadata but lose
obsolete traversal caches. `reader.excel_metadata.hierarchies` exposes
`processing_state` and `structure_type`. These operations do not verify DAX
result uniqueness or a one-to-one label/sort-key mapping before Excel calculates.

Use the ordinary `add_relationship` API for calculated endpoints. Relationship
types, topology, and available stored lookup keys are checked immediately.
Unprocessed calculated keys cannot be checked for uniqueness until Excel
calculates them; `validate()` reports `unprocessed_relationship_keys` as a
warning. Invalid stored lookup keys still raise a Python validation error.
Native checks cover new calculated relationships and changed stored inputs,
including correct filtering after saving and reopening.

Invalid calculated definitions raise `CalculatedColumnValidationError`, which
is both `ModelValidationError` and `ValueError`. `validate()` reports intended
unprocessed columns as `unprocessed_calculated_column` warnings, not failures.

## KPI configurations

A KPI attaches to an existing measure. Choose exactly one target: an existing
measure (`target_measure`, optionally `target_table`) or finite `target_value`.
The legacy `lower=0.4, upper=0.8` interface remains available. Explicit
`thresholds` are increasing numeric boundaries. `threshold_type='percentage'`
uses fractions of the target; `'absolute'` uses raw values. The defaults are
percentage for measure targets and absolute for fixed targets.

```python
model.add_kpi('Sales', 'Total Amount', target_value=100,
              mode='descending', thresholds=(40, 80), icon='Road Signs')
model.update_kpi('Sales', 'Total Amount', target_measure='Target Amount',
                 mode='centered', thresholds=(0.5, 0.8, 1.2, 1.5), icon='Shapes')
```

| Mode | Preferred values | Normal icon boundaries | Gauge boundaries |
|---|---|---:|---:|
| `ascending` | Higher | 2 | 4 |
| `descending` | Lower | 2 | 4 |
| `centered` | Middle range | 4 | 8 |
| `centered_reversed` | Outer ranges | 4 | 8 |

Normal icons are `Traffic Light - Single`, `Road Signs`, `Traffic Light` and
`Shapes`. The five-state icon is `Gauge - Ascending`. Centered percentage
thresholds may exceed 1. Blank measures produce blank status. Invalid targets,
unsupported icons/layouts and helper dependencies raise structured
`ModelValidationError` subclasses with validation issues. KPI helper measures
are implementation details; avoid referencing them unless you intend to retain
those dependencies when updating or deleting the KPI.

Native acceptance is tracked separately from Python round-trip tests in the
project's compatibility results. A readable file or valid metadata is not proof
that every configuration has been accepted and calculated by Excel.
