"""Excel date-table key metadata; DAX time intelligence remains Excel's job."""
from xml.dom import minidom

from .abf import AbfImage, UnsupportedModelError
from .schema import _child, _elements, _text
from .errors import ModelValidationError


class DateTableValidationError(ModelValidationError, ValueError):
    """Invalid date-table key or unsupported designation."""


class DateTableLookupError(ModelValidationError, KeyError):
    """Missing date-table object, preserving KeyError compatibility."""


def _error(table, column, message, *, lookup=False):
    from .validation import ValidationIssue
    error = DateTableLookupError if lookup else DateTableValidationError
    raise error(message, issues=(ValidationIssue('invalid_date_table', message,
        object_path=('tables', str(table), 'columns', str(column))),))


def validate_date_table_key(model, table, column):
    """Validate staged date values without changing the model."""
    import pandas as pd
    from .relationship_storage import relationship_key_type
    if table not in model._dimension_ids:
        _error(table, column, f'Unknown table: {table}', lookup=True)
    from .calculated_columns import list_calculated_columns
    if any(c.table_id == model._dimension_ids[table] and c.name == column
           for c in list_calculated_columns(model._image)):
        _error(table, column, 'Calculated date-table keys are not supported; use a stored datetime column. '
               'Promoting a calculated key can introduce a native circular dependency, and pending values '
               'cannot be checked for uniqueness or contiguous dates.')
    try:
        data_type = relationship_key_type(model, table, column)
        values = pd.Series(model.get_table(table)[column])
    except (KeyError, StopIteration):
        _error(table, column, f'Unknown date key: {table}[{column}]', lookup=True)
    if data_type != 'Date':
        _error(table, column, 'A date table key must have the datetime data type')
    if values.empty or values.isna().any():
        _error(table, column, 'A date table key must be nonempty and contain no blanks')
    dates = pd.to_datetime(values)
    if dates.duplicated().any():
        _error(table, column, 'A date table key must contain unique dates')
    if not dates.eq(dates.dt.normalize()).all():
        _error(table, column, 'A date table key must contain whole days without time components')
    ordered = dates.sort_values().reset_index(drop=True)
    if len(ordered) > 1 and not ordered.diff().iloc[1:].eq(pd.Timedelta(days=1)).all():
        _error(table, column, 'A date table key must contain contiguous dates')


def _set(node, name, value):
    child = _child(node, name)
    while child.firstChild:
        child.removeChild(child.firstChild)
    child.appendChild(node.ownerDocument.createTextNode(str(value)))


def configure_date_table(image, table_id, column_id=None):
    """Mark a stored column as the date key, or restore the internal row key.

    The caller validates the date values before changing metadata. Native files
    identify a date key by dimension Type=Time, Usage=Key, and ColumnFlags bit 4.
    Existing data streams and their version references remain consistent.
    """
    replacements = {}
    dimension_file = next((f for f in image.files if f.logical_path and
                          f.logical_path.replace('\\', '/').rsplit('/', 1)[-1].startswith(table_id + '.') and
                          f.logical_path.endswith('.dim.xml')), None)
    if dimension_file is None:
        _error(table_id, column_id, f'Unknown table: {table_id}', lookup=True)
    doc = minidom.parseString(image.read_logical_file(dimension_file))
    dimension = _elements(doc, 'Dimension')[0]
    attributes = list(_child(dimension, 'Attributes').childNodes)
    attributes = [n for n in attributes if n.nodeType == n.ELEMENT_NODE]
    row = next(a for a in attributes if _text(_child(a, 'Type')) == 'RowNumber')
    row_id = _text(_child(row, 'ID'))
    old_key = next(a for a in attributes if _text(_child(a, 'Usage')) == 'Key')
    new_key = row if column_id is None else next(
        (a for a in attributes if _text(_child(a, 'ID')) == column_id), None)
    if new_key is None:
        _error(table_id, column_id, f'Unknown date key: {column_id}', lookup=True)
    if new_key is not row and _elements(new_key, 'Expression'):
        _error(table_id, column_id, 'Date keys must be stored columns')
    if new_key is old_key:
        return image
    old_id = _text(_child(old_key, 'ID'))
    new_id = _text(_child(new_key, 'ID'))
    # Reverse only the row-key edge. Other user/engine attribute relationships
    # are preserved instead of rebuilt from guessed dependency rules.
    old_relations = _child(old_key, 'AttributeRelationships')
    new_relations = _child(new_key, 'AttributeRelationships')
    if old_key is not row:
        edge = next(n for n in _elements(old_relations, 'AttributeRelationship')
                    if _text(_child(n, 'AttributeID')) == row_id)
        old_relations.removeChild(edge)
        _set(edge, 'AttributeID', old_id)
        _set(edge, 'Name', _text(_child(old_key, 'Name')))
        # The old date no longer supplies row identity. Retain its uniqueness
        # only when an existing relationship still requires a one-side key.
        required_unique = False
        for candidate in image.files:
            if not candidate.logical_path or not candidate.logical_path.endswith('.dim.xml'):
                continue
            related = minidom.parseString(image.read_logical_file(candidate))
            for relationship in _elements(related, 'Relationship'):
                for end in (_elements(relationship, 'FromRelationshipEnd') +
                            _elements(relationship, 'ToRelationshipEnd')):
                    if (_text(_child(end, 'DimensionID')) == table_id and
                            any(_text(a) == old_id for a in _elements(end, 'AttributeID')) and
                            any(_text(m) == 'One' for m in _elements(end, 'Multiplicity'))):
                        required_unique = True
        _set(edge, 'Cardinality', 'One' if required_unique else 'Many')
        _child(row, 'AttributeRelationships').appendChild(edge)
    if new_key is not row:
        relations = _child(row, 'AttributeRelationships')
        edge = next(n for n in _elements(relations, 'AttributeRelationship')
                    if _text(_child(n, 'AttributeID')) == new_id)
        relations.removeChild(edge)
        _set(edge, 'AttributeID', row_id)
        _set(edge, 'Name', _text(_child(row, 'Name')))
        # Native date keys identify each row uniquely. Preserving the old
        # many-side edge loads alone but fails once another relationship is added.
        _set(edge, 'Cardinality', 'One')
        new_relations.appendChild(edge)
    _set(old_key, 'Usage', 'Regular')
    _set(new_key, 'Usage', 'Key')
    _set(dimension, 'Type', 'Regular' if new_key is row else 'Time')
    replacements[dimension_file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    # Processing resolves the dimension key through the owning measure group's
    # granularity. Leaving RowNumber here makes a date key appear to require an
    # indirect self-relationship when Excel next recalculates the model.
    for f in image.files:
        if not f.logical_path or not f.logical_path.endswith('.det.xml'):
            continue
        detail = minidom.parseString(image.read_logical_file(f))
        groups = _elements(detail, 'MeasureGroup')
        if not groups or _text(_child(groups[0], 'ID')) != table_id:
            continue
        for binding in _elements(_child(groups[0], 'Dimensions'), 'Dimension'):
            if (_text(_child(binding, 'CubeDimensionID')) != table_id or
                    not binding.getAttribute('xsi:type').endswith('DegenerateMeasureGroupDimension')):
                continue
            for attribute in _elements(_child(binding, 'Attributes'), 'Attribute'):
                identifier = _text(_child(attribute, 'AttributeID'))
                if identifier in (old_id, new_id):
                    _set(attribute, 'Type', 'Granularity' if identifier == new_id else 'Regular')
        replacements[f.storage_path] = detail.documentElement.toxml(encoding='utf-8')
    for f in image.files:
        if not f.logical_path or not f.logical_path.replace('\\', '/').rsplit('/', 1)[-1].startswith(table_id + '.') or not f.logical_path.endswith('.tbl.xml'):
            continue
        table_doc = minidom.parseString(image.read_logical_file(f))
        for raw in _elements(table_doc, 'XMObject'):
            if raw.getAttribute('class') != 'XMRawColumn':
                continue
            name = raw.getAttribute('name')
            if name not in (old_id, new_id):
                continue
            props = _child(raw, 'Properties')
            flags = int(_text(_child(props, 'ColumnFlags')))
            _set(props, 'ColumnFlags', flags | 4 if name == new_id else flags & ~4)
        replacements[f.storage_path] = table_doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))
