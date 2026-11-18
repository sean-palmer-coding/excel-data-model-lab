from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .model import WorkbookModel


# Deliberately finite: a specification cannot invoke arbitrary Python methods,
# save intermediate outputs, refresh a source, or execute a shell command.
EDIT_OPERATIONS = frozenset({
    'add_column', 'remove_column', 'rename_column', 'change_column_type',
    'remove_table', 'rename_table', 'configure_table', 'configure_column',
    'add_measure', 'update_measure', 'rename_measure', 'remove_measure',
    'add_calculated_column', 'update_calculated_column', 'remove_calculated_column',
    'add_kpi', 'update_kpi', 'remove_kpi',
    'add_hierarchy', 'update_hierarchy', 'remove_hierarchy', 'configure_hierarchy',
    'set_sort_column', 'add_relationship', 'remove_relationship', 'set_relationship_active',
    'mark_date_table', 'unmark_date_table', 'create_perspective', 'update_perspective',
    'rename_perspective', 'remove_perspective', 'unbind_source',
})


def _source_binding(model, table, definition):
    from .native_sources import LinkedTableSource, OleDbSource
    settings = dict(definition)
    kind = settings.pop('kind')
    if kind == 'linked_table':
        source = LinkedTableSource(**settings)
    elif kind == 'ole_db':
        source = OleDbSource(**settings)
    elif kind == 'snapshot' and not settings:
        return model.unbind_source(table)
    else:
        raise ValueError('Unsupported specification source binding kind: ' + str(kind))
    return model.bind_source(table, source)


def load_spec(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        value = json.loads(text)
        if not isinstance(value, dict):
            raise ValueError("Model specification must be a mapping")
        return value
    if path.suffix.lower() in {".yaml", ".yml"}:
        import yaml

        value = yaml.safe_load(text)
        if not isinstance(value, dict):
            raise ValueError("Model specification must be a mapping")
        return value
    raise ValueError("Model specification must be JSON or YAML")


def _field_reference(model, value):
    """Resolve ``Table.Column`` (or ``[table, column]``) against known tables.

    Table names may contain '.', so prefer the longest known table prefix.
    """
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return str(value[0]), str(value[1])
    if not isinstance(value, str) or "." not in value:
        raise ValueError(f"Relationship endpoint must be 'Table.Column': {value!r}")
    tables = {column.table for column in model.columns}
    matches = [t for t in tables if value.startswith(t + ".") and len(value) > len(t) + 1]
    if matches:
        table = max(matches, key=len)
        return table, value[len(table) + 1:]
    return tuple(value.split(".", 1))


def build_from_spec(
    template: str | Path | None,
    destination: str | Path,
    spec: dict[str, Any],
    *,
    base_directory: str | Path | None = None,
    workbook: str | Path | None = None,
) -> Path:
    import pandas as pd

    base = Path(base_directory) if base_directory is not None else Path.cwd()
    unknown = set(spec) - {'tables', 'relationships', 'measures', 'calculated_columns', 'kpis', 'perspectives', 'operations'}
    if unknown:
        raise ValueError('Unknown specification fields: ' + ', '.join(sorted(unknown)))
    if template is not None and workbook is not None:
        raise ValueError("Use either a model template or an ordinary workbook, not both")
    frames = {}
    for table, definition in spec.get("tables", {}).items():
        if isinstance(definition, str):
            source = definition
        else:
            source = definition.get("source")
            if source is None:
                if template is None:
                    raise ValueError('A new table needs a source: ' + table)
                continue
        source_path = base / source
        suffix = source_path.suffix.lower()
        if suffix == ".csv":
            options = definition.get("read_options", {}) if isinstance(definition, dict) else {}
            frame = pd.read_csv(source_path, **options)
        elif suffix == ".parquet":
            frame = pd.read_parquet(source_path)
        else:
            raise ValueError(f"Unsupported table source: {source_path}")
        frames[table] = frame
    if template is None:
        model = WorkbookModel.create(frames, workbook=workbook)
    else:
        model = WorkbookModel.from_template(template)
        known = {column.table for column in model.columns}
        for table, frame in frames.items():
            if table in known:
                model.replace_table(table, frame)
            else:
                model.add_table(table, frame)
    for table, definition in spec.get("tables", {}).items():
        if not isinstance(definition, dict):
            continue
        model.configure_table(table, hidden=definition.get("hidden"), description=definition.get("description"))
        for column, settings in definition.get("columns", {}).items():
            # Sort-only entries may target calculated columns, which
            # configure_column does not accept.
            if any(key in settings for key in ("hidden", "description", "format")):
                model.configure_column(table, column, hidden=settings.get("hidden"),
                    description=settings.get("description"), format_string=settings.get("format"))
            if "sort_by" in settings:
                model.set_sort_column(table, column, settings["sort_by"])
        for hierarchy in definition.get("hierarchies", []):
            model.add_hierarchy(table, hierarchy["name"], hierarchy["levels"])
            model.configure_hierarchy(table, hierarchy["name"], hidden=hierarchy.get("hidden"),
                description=hierarchy.get("description"), level_names=hierarchy.get("level_names"))
    for item in spec.get("relationships", []):
        from_table, from_column = _field_reference(model, item["from"])
        to_table, to_column = _field_reference(model, item["to"])
        model.add_relationship(
            from_table, from_column, to_table, to_column,
            active=item.get("active", True),
            cardinality=item.get("cardinality", "many-to-one"),
        )
    for item in spec.get("measures", []):
        model.add_measure(
            item["table"], item["name"], item["expression"],
            format_string=item.get("format", ""),
            display_folder=item.get("display_folder", ""),
            description=item.get("description", ""),
        )
    for item in spec.get('calculated_columns', []):
        model.add_calculated_column(**item)
    for item in spec.get('kpis', []):
        model.add_kpi(**item)
    for table, definition in spec.get('tables', {}).items():
        if isinstance(definition, dict):
            if definition.get('date_key') is not None:
                model.mark_date_table(table, definition['date_key'])
            if 'source_binding' in definition:
                _source_binding(model, table, definition['source_binding'])
    for item in spec.get('perspectives', []):
        model.create_perspective(**item)
    with model.transaction() as edit:
        for operation in spec.get('operations', []):
            settings = dict(operation)
            name = settings.pop('op')
            if name == 'bind_source':
                _source_binding(edit, settings.pop('table'), settings.pop('source'))
                if settings:
                    raise ValueError('Unknown bind_source operation arguments')
            elif name in EDIT_OPERATIONS:
                getattr(edit, name)(**settings)
            else:
                raise ValueError('Unsupported specification operation: ' + str(name))
    return model.save(destination)
