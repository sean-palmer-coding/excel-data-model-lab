# Date tables and perspectives

These APIs edit the embedded model. They do not refresh sources, calculate DAX,
or change worksheet PivotTable filters.

## Date tables

```python
import pandas as pd
from excel_model_lab import WorkbookModel, ModelReader

model = WorkbookModel.create({
    'Calendar': pd.DataFrame({'Date': pd.date_range('2026-01-01', periods=365)})
})
model.mark_date_table('Calendar', 'Date')
model.validate().raise_for_errors()
path = model.save('calendar.xlsx')
with ModelReader(path) as reader:
    print(reader.excel_metadata.date_tables)

model.unmark_date_table('Calendar')
```

A date key must be a stored datetime column with at least one row, no blanks,
unique dates, no time components, and a contiguous daily range. Row ordering
does not matter. The API changes key metadata and preserves stored values.
Marking a different valid column moves the designation; unmarking restores the
internal row key. Referenced date keys cannot be removed until the designation
is removed. Editing a marked key is checked again before saving, so invalid
pending values cannot be published accidentally.

Validation reports use `invalid_date_table` issues. Preflight checks establish
the supported writer contract; they are not a replacement for native Excel
acceptance or DAX correctness checks.

Direct invalid-key errors are `DateTableValidationError` (`ValueError` compatible);
missing objects raise `DateTableLookupError` (`KeyError` compatible). Both expose
structured issues through `ModelValidationError`.

Direct perspective errors expose structured issues through
`PerspectiveValidationError`, `PerspectiveLookupError` and `PerspectiveTypeError`,
preserving `ValueError`, `KeyError` and `TypeError` compatibility respectively.

## Perspectives

```python
model.create_perspective('Reporting',
    columns=[('Calendar', 'Date')], description='Report fields')
model.rename_perspective('Reporting', 'Calendar View')
model.update_perspective('Calendar View', columns=[('Calendar', 'Date')])
print(model.list_perspectives())
model.remove_perspective('Calendar View')
```

Members are explicit `(table caption, object caption)` pairs. `columns`,
`measures`, and `hierarchies` accept these pairs. Duplicate selections are
deduplicated; unknown objects and empty selections are rejected. Updates
replace the complete selection, so omitted member categories become empty.
Omitting `description` replaces it with the empty string.

Perspective names are case-insensitively unique and cannot be `Model`, contain
filename-reserved characters or semicolons, or have leading/trailing whitespace.
Renaming preserves internal identity and storage paths. The previous caption
can subsequently be reused for a separate perspective with a new identity.
Membership does not change object visibility in the default model. References
from perspectives prevent removing selected fields until membership is updated.
Report bindings may also prevent renaming or removing a perspective.

The read-only `reader.excel_metadata` snapshot exposes perspectives,
`perspective_columns`, `perspective_measures`, and native hierarchy annotations.
Native Excel stores hierarchy selections in a separate cube annotation.
Multiple hierarchies from the same table use comma-separated captions, observed
in the installed native serializer and reader. Captions containing commas are
rejected because the format has no escaping. Models with hierarchy annotations
also reject perspective captions containing apostrophes: the native reader
interpolates each perspective caption into a single-quoted XPath literal.
Apostrophes remain allowed when the model has no hierarchy annotation.
`perspective_hierarchies` resolves selected IDs, while the raw annotation remains
available separately. Multiple-hierarchy writing is experimental pending native
workbook acceptance, despite the observed serialization rule.

Format provenance: read-only reflection of the installed
`Microsoft.AnalysisServices.Excel.BackEnd.dll`, file version `2015.130.1613.1`,
SHA256 `14d38261ea31dec6ec99a7bef4b3bbeb38133d757ea2c69b4aba6258ac0314e3`.
The reader and writer both expose the constant
`EMPTY_HIERARCHY_NAMES_DELIMITER` as comma. Reader method behavior confirms a
character split without escaping. This implementation independently encodes
observed format facts; no Microsoft code or binaries are copied or redistributed.

Updates retain perspective identity, translations, top-level extension fields,
wrapper metadata, and extensions attached to retained members and their
containers. Removed members and their metadata are removed together. Native acceptance results are tracked in
the compatibility reports rather than inferred from XML round-trip tests.
