# CLI and model specifications

`excelmodel` and `xmodel-lab` are equivalent installed commands. JSON and YAML specifications support the same fields. Paths in table `source` entries are relative to the specification file.

```
excelmodel create output.xlsx model.json
excelmodel build existing.xlsx output.xlsx edits.json
excelmodel dependencies output.xlsx
excelmodel validate output.xlsx
```

`dependencies` emits objects, edges and unresolved lexical references as JSON. It does not evaluate DAX, infer every virtual-table reference, or certify Excel compatibility. `validate` emits static preflight issues and exits 1 when errors are present; success does not prove Excel will load or calculate the workbook. Neither command refreshes connections or modifies the workbook.

## Authoring fields

- `tables`: CSV or Parquet sources; per-table `hidden`, `description`, `columns` formatting/visibility/sort settings, `hierarchies`, `date_key`, and `source_binding`.
- `relationships`: `from`/`to` strings, `active`, `cardinality`. For names containing dots, use an `add_relationship` operation with separate table/column arguments.
- `measures`: table, name, expression, format, display_folder, description.
- `calculated_columns`: table, name, expression, data_type. DAX values remain unprocessed until Excel calculates them.
- `kpis`: the public `add_kpi` keyword arguments, including fixed or measure targets, modes, icons and thresholds.
- `perspectives`: name and columns/measures/hierarchies selections as `[table, object]` pairs.
- `operations`: ordered edits using public keyword argument names and an `op` name. This is an explicit allowlist, not arbitrary Python execution.

A table definition may omit `source` when editing a template's existing table. When creating a new model every table requires a source. CSV `read_options` are forwarded to pandas, for example `{"parse_dates":["Date"]}` before marking a date key. Parquet requires the optional parquet dependency.

Authoring order is table data, table/column presentation and hierarchies, relationships, measures, calculated columns, KPIs, date keys/source bindings, perspectives, then ordered operations. Use operations when a different dependency order is required. The destination is written only after the entire build succeeds. Unknown top-level fields and operation names are rejected.

```json
{
  "tables": {"Sales": {"source": "sales.csv"}},
  "measures": [{"table":"Sales","name":"Total","expression":"SUM(Sales[Amount])"}],
  "calculated_columns": [{"table":"Sales","name":"Double","expression":"Sales[Amount]*2","data_type":"integer"}],
  "kpis": [{"table":"Sales","measure":"Total","target_value":100,"thresholds":[40,80]}],
  "perspectives": [{"name":"Finance","columns":[["Sales","Amount"]],"measures":[["Sales","Total"]]}],
  "operations": [{"op":"rename_table","table":"Sales","name":"Sales Data"}]
}
```

Supported operations cover add/remove/rename/type-change columns; remove/rename/configure tables; add/update/remove measures, calculated columns, KPIs and hierarchies; hierarchy/column presentation; sorting; add/remove/activate relationships; date-table mark/unmark; perspective create/update/rename/remove; and source bind/unbind. Their argument names match the corresponding Python methods. A source operation uses `{"op":"bind_source","table":"Sales","source":{"kind":"linked_table","table_name":"SalesInput"}}`. Native source kinds are `linked_table`, `ole_db`, and `snapshot`; see [source binding documentation](SOURCE_AND_DEPENDENCY_API.md) for connection settings and limitations. Credentials are rejected and no connection is opened by the specification runner.

CLI `add-relationship` also accepts `--cardinality many-to-one`, `one-to-many` or `one-to-one`; Python's compatibility checks still apply. All edit operations retain the Python API's restrictive dependency/report guards. Existing report bindings are not automatically rewritten.

The ordered-operation allowlist also includes `rename_measure` with `table`, `name`, and `new_name`. It follows the restrictive dependency rules described in SOURCE_AND_DEPENDENCY_API.md.
