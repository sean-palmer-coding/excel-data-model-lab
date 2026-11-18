"""Non-mutating preflight checks for the supported Excel writer contract."""
from dataclasses import dataclass
from .abf import UnsupportedModelError
from .errors import ModelValidationError
from .relationships import RelationshipValidationError, validate_relationship_keys, validate_relationship_collation


@dataclass(frozen=True)
class ValidationIssue:
    code: str
    message: str
    severity: str = 'error'
    object_path: tuple[str, ...] = ()
    dependencies: tuple[tuple[str, ...], ...] = ()


@dataclass(frozen=True)
class ValidationReport:
    issues: tuple[ValidationIssue, ...]

    @property
    def valid(self):
        return not any(issue.severity == 'error' for issue in self.issues)

    def raise_for_errors(self):
        errors = tuple(issue for issue in self.issues if issue.severity == 'error')
        if errors:
            relationship_codes = {'duplicate_relationship', 'self_relationship',
                'ambiguous_path', 'unsupported_relationship', 'invalid_relationship'}
            error_type = (RelationshipValidationError if all(
                issue.code in relationship_codes for issue in errors) else ModelValidationError)
            raise error_type('\n'.join(issue.message for issue in errors), issues=errors)


def validate_model(model):
    """Check relationship keys/types/topology and report writer limitations.

    This is a preflight, not proof that Excel will load the eventual output.
    Stored data is decoded but neither the model nor its destination is modified.
    """
    from .relationship_storage import stored_relationships, relationship_key_type, validate_cached_relationship_keys
    issues = []
    from .excel_metadata import ExcelMetadata
    metadata = ExcelMetadata.from_files((f.logical_path, model._image.read_logical_file(f))
        for f in model._image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
    for row in metadata.calculated_columns.itertuples(index=False):
        if row.processing_state.casefold() == 'unprocessed':
            issues.append(ValidationIssue('unprocessed_calculated_column',
                f'{row.table}.{row.name}: Excel must calculate this column before cached values are available',
                'warning', object_path=('tables', row.table, 'columns', row.name)))
    for row in metadata.date_tables.itertuples(index=False):
        key = row.date_key_columns[0] if len(row.date_key_columns) == 1 else ''
        matching = metadata.columns[(metadata.columns.table_id == row.id) & (metadata.columns.id == key)]
        if len(matching) == 1:
            key = matching.iloc[0]['name']
        try:
            model._validate_date_table_key(row.name, key)
        except (ValueError, KeyError, UnsupportedModelError) as exc:
            issues.append(ValidationIssue('invalid_date_table', str(exc),
                object_path=('tables', row.name, 'columns', key)))
    # Reuse the actual script writer's checks without assembling or mutating the
    # model. This includes pending measures, which readers cannot see yet.
    from .measures import add_measures_to_script
    if model._pending_measures:
        try:
            script = next(f for f in model._image.files
                          if f.logical_path and f.logical_path.endswith('.scr.xml'))
            add_measures_to_script(model._image.read_logical_file(script), model._pending_measures)
        except (ValueError, UnsupportedModelError) as exc:
            issues.append(ValidationIssue('invalid_measure', str(exc)))
    relationships = [r for _, r in stored_relationships(model)] + model._pending_relationships
    tables = {}
    seen = set()
    graph = {}
    for r in relationships:
        edge = (r.from_table, r.from_column, r.to_table, r.to_column)
        if edge in seen:
            issues.append(ValidationIssue('duplicate_relationship', f'Duplicate relationship: {edge}', object_path=('relationships', *edge)))
        seen.add(edge)
        if r.from_table == r.to_table:
            issues.append(ValidationIssue('self_relationship', f'Self relationship: {edge}', object_path=('relationships', *edge)))
        if r.active:
            frontier, visited = [r.from_table], set()
            while frontier:
                node = frontier.pop()
                if node == r.to_table:
                    issues.append(ValidationIssue('ambiguous_path', f'Ambiguous active path: {edge}', object_path=('relationships', *edge)))
                    break
                if node not in visited:
                    visited.add(node)
                    frontier.extend(graph.get(node, set()) - visited)
            graph.setdefault(r.from_table, set()).add(r.to_table)
            graph.setdefault(r.to_table, set()).add(r.from_table)
        try:
            validate_relationship_collation(model, r)
            if relationship_key_type(model, r.from_table, r.from_column) != relationship_key_type(model, r.to_table, r.to_column):
                raise RelationshipValidationError(f'Relationship columns have different types: {edge}')
            for table in (r.from_table, r.to_table):
                if table not in tables:
                    tables[table] = model.get_table(table)
            if validate_cached_relationship_keys(model, r):
                issues.append(ValidationIssue('unprocessed_relationship_keys',
                    'Excel must calculate and validate relationship keys before querying this model',
                    'warning', object_path=('relationships', *edge)))
        except UnsupportedModelError as exc:
            issues.append(ValidationIssue('unsupported_relationship', str(exc), object_path=('relationships', *edge)))
        except (RelationshipValidationError, KeyError) as exc:
            issues.append(ValidationIssue('invalid_relationship', str(exc), object_path=('relationships', *edge)))
    for column in model.columns:
        if not column.supported:
            issues.append(ValidationIssue('read_only_column',
                f'{column.table}.{column.name}: {column.reason}', 'warning',
                object_path=('tables', column.table, 'columns', column.name)))
    return ValidationReport(tuple(issues))
