"""Calculated-column definitions with explicitly unprocessed BLANK storage.

This module stores DAX; it does not evaluate it. Excel must calculate the column
before its cached values can be consumed. All changes return a new ABF image.
"""
from dataclasses import dataclass
import struct
import re
from xml.dom import minidom
from .abf import AbfImage, UnsupportedModelError
from .schema import _elements, _child, _text, clone_column
from .errors import ModelValidationError


class CalculatedColumnValidationError(ModelValidationError, ValueError):
    """Invalid calculated definition, retaining ValueError compatibility."""


def _error(table, name, message, code='invalid_calculated_column'):
    from .validation import ValidationIssue
    raise CalculatedColumnValidationError(message, issues=(ValidationIssue(
        code, message, object_path=('tables', str(table), 'columns', str(name))),))

_TYPES = {'integer': ('BigInt', 20, 'Long', 0),
          'decimal': ('Double', 5, 'Real', 1),
          'currency': ('Currency', 6, 'Long', 0),
          'datetime': ('Date', 7, 'Real', 1),
          'boolean': ('Boolean', 11, 'Long', 0),
          'text': ('WChar', 130, 'String', 2)}


@dataclass(frozen=True)
class CalculatedColumn:
    table_id: str
    name: str
    expression: str
    data_type: str
    processing_state: str
    id: str | None = None


def _dimensions(image):
    for f in image.files:
        if f.logical_path and f.logical_path.endswith('.dim.xml'):
            doc = minidom.parseString(image.read_logical_file(f))
            for d in _elements(doc, 'Dimension'):
                if getattr(d.parentNode, 'localName', None) == 'ObjectDefinition':
                    yield f, doc, d


def list_calculated_columns(image):
    result = []
    for _, _, dim in _dimensions(image):
        for attr in _elements(_child(dim, 'Attributes'), 'Attribute'):
            expressions = _elements(attr, 'Expression')
            if expressions:
                inferred = [n for n in attr.childNodes if getattr(n, 'localName', None) == 'InferredDatatype']
                state = [n for n in attr.childNodes if getattr(n, 'localName', None) == 'ProcessingState']
                result.append(CalculatedColumn(_text(_child(dim, 'ID')), _text(_child(attr, 'Name')),
                    _text(expressions[0]), _text(inferred[0]) if inferred else '', _text(state[0]) if state else '',
                    _text(_child(attr, 'ID'))))
    return tuple(result)


def _set(node, name, value):
    matches = [n for n in node.childNodes if getattr(n, 'localName', None) == name]
    target = matches[0] if matches else node.appendChild(node.ownerDocument.createElement(name))
    while target.firstChild:
        target.removeChild(target.firstChild)
    target.appendChild(node.ownerDocument.createTextNode(str(value)))
    return target


def _validate(name, expression, data_type, table_id=''):
    if not isinstance(name, str) or not name.strip() or name != name.strip() or any(ord(c) < 32 or c in '\\/$' for c in name):
        _error(table_id, name, 'Calculated column name must be nonempty and trimmed without controls, path separators or $')
    if not isinstance(expression, str) or not expression.strip():
        _error(table_id, name, 'A nonempty DAX expression is required')
    if not isinstance(data_type, str) or data_type not in _TYPES:
        _error(table_id, name, 'data_type must be one of: ' + ', '.join(_TYPES))
    expression = expression.strip().removeprefix('=').strip()
    if not expression:
        _error(table_id, name, 'A nonempty DAX expression is required')
    return expression


def add_calculated_column(image, table_id, name, expression, *, data_type):
    expression = _validate(name, expression, data_type, table_id)
    dim = next((d for _, _, d in _dimensions(image) if _text(_child(d, 'ID')) == table_id), None)
    if dim is None:
        raise KeyError(table_id)
    attrs = _elements(_child(dim, 'Attributes'), 'Attribute')
    if any(_text(_child(a, field)).casefold() == name.casefold() for a in attrs for field in ('ID', 'Name')):
        _error(table_id, name, 'Column already exists: ' + name, 'duplicate_column')
    source = next((a for a in attrs if _text(_child(a, 'Usage')) == 'Regular' and not _elements(a, 'Expression')), None)
    if source is None:
        raise UnsupportedModelError('Adding a calculated column requires a stored column in the table')
    cloned = clone_column(image, table_id, _text(_child(source, 'ID')), name)
    # clone_column preserves the requested caption but may allocate a legal
    # physical ID for punctuation-bearing names. All storage invalidation must
    # use that resolved identity, never the user-facing caption.
    file, document, dimension = next((f, doc, d) for f, doc, d in _dimensions(cloned)
                                     if _text(_child(d, 'ID')) == table_id)
    attribute = next(a for a in _elements(_child(dimension, 'Attributes'), 'Attribute')
                     if _text(_child(a, 'Name')) == name)
    identity = _text(_child(attribute, 'ID'))
    return _invalidate(cloned, table_id, identity, expression, data_type)


def update_calculated_column(image, table_id, name, expression, *, data_type):
    expression = _validate(name, expression, data_type, table_id)
    column = next((c for c in list_calculated_columns(image) if c.table_id == table_id and c.name == name), None)
    if column is None:
        raise KeyError((table_id, name))
    return _invalidate(image, table_id, column.id, expression, data_type)


def remove_calculated_column(image, table_id, name):
    column = next((c for c in list_calculated_columns(image) if c.table_id == table_id and c.name == name), None)
    if column is None:
        raise KeyError((table_id, name))
    from .schema import remove_column
    return remove_column(image, table_id, column.id)



def invalidate_calculated_columns(image, *, changed_tables=None):
    """Clear every calculated cache after input changes, retaining definitions.

    Row counts must already match the saved storage. Invalidation is global
    except for proven local-row scalar expressions in unchanged tables. Merely
    lacking an explicit cross-table reference does not establish independence:
    measures, context transitions and unrecognized functions invalidate globally.
    ``changed_tables`` contains physical table IDs, not display captions.
    """
    mapping = {values[0]: key for key, values in _TYPES.items()}
    columns = list_calculated_columns(image)
    if changed_tables is not None:
        changed_tables = frozenset(changed_tables)
        columns = tuple(c for c in columns if c.table_id in changed_tables
                        or not _local_row_expression(image, c))
    for column in columns:
        if column.data_type not in mapping:
            raise UnsupportedModelError('Cannot invalidate calculated type: ' + column.data_type)
    for column in columns:
        image = _invalidate(image, column.table_id, column.id, column.expression, mapping[column.data_type])
    return image




def resize_calculated_columns(image, table_id, row_count):
    """Rebuild unprocessed calculated caches after stored table rows are resized.

    Call after updating stored column and table row metadata. This operation
    retains all DAX definitions and requires Excel to calculate the new rows.
    """
    import operator
    row_count = operator.index(row_count)
    if row_count < 0:
        raise ValueError('Calculated row count cannot be negative')
    calculated_ids = {c.id for c in list_calculated_columns(image) if c.table_id == table_id}
    table_file = next(f for f in image.files if f.logical_path and
        f.logical_path.split('\\')[-1].startswith(table_id + '.') and f.logical_path.endswith('.tbl.xml'))
    doc = minidom.parseString(image.read_logical_file(table_file))
    counts = None
    for raw in _elements(doc, 'XMObject'):
        if raw.getAttribute('class') != 'XMRawColumn' or raw.getAttribute('name') in calculated_ids | {'RowNumber', '__XL_RowNumber'}:
            continue
        segments = next(c for c in _elements(raw, 'Collection') if _text(_child(c, 'Name')) == 'Segments')
        candidate = [int(_text(_child(_child(s, 'Properties'), 'Records'))) for s in segments.childNodes
                     if getattr(s, 'localName', None) == 'XMObject']
        if sum(candidate) == row_count:
            counts = candidate
            break
    if counts is None:
        raise UnsupportedModelError('Stored row metadata must be resized before calculated caches')
    mapping = {values[0]: key for key, values in _TYPES.items()}
    for column in list_calculated_columns(image):
        if column.table_id != table_id:
            continue
        if column.data_type not in mapping:
            raise UnsupportedModelError('Cannot resize calculated type: ' + column.data_type)
        image = _invalidate(image, table_id, column.id, column.expression,
                            mapping[column.data_type], record_counts=counts)
    return image


def _local_row_expression(image, column):
    """Prove a deliberately narrow scalar subset independent of other tables."""
    from .formula_statements import formula_references, protected_spans
    dimensions = list(_dimensions(image))
    dim = next(d for _, _, d in dimensions if _text(_child(d, 'ID')) == column.table_id)
    table_name = _text(_child(dim, 'Name'))
    stored = {_text(_child(a, 'Name')).casefold() for a in _elements(_child(dim, 'Attributes'), 'Attribute')
              if not _elements(a, 'Expression') and _text(_child(a, 'Type')) != 'RowNumber'}
    table_names = [_text(_child(d, 'Name')) for _, _, d in dimensions]
    try:
        refs = formula_references(column.expression, table_names=table_names)
        if any(r.kind != 'field' or r.name.casefold() not in stored or
               (r.table is not None and r.table.casefold() != table_name.casefold()) for r in refs):
            return False
        masked = list(column.expression)
        for start, end in protected_spans(column.expression):
            masked[start:end] = ' ' * (end - start)
        for ref in refs:
            masked[ref.start:ref.end] = ' ' * (ref.end - ref.start)
        # This whitelist is intentionally small, deterministic and row scalar.
        # Unknown functions (especially CALCULATE, RELATED, aggregates, NOW and
        # RAND) never qualify merely because their explicit references are local.
        words = re.findall(r'[A-Za-z_][A-Za-z_0-9]*', ''.join(masked))
        allowed = {'FORMAT', 'LEN', 'LOWER', 'UPPER', 'TRIM', 'LEFT', 'RIGHT', 'MID',
                   'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'DATE',
                   'ABS', 'INT', 'ROUND', 'MOD', 'IF', 'ISBLANK', 'BLANK', 'TRUE', 'FALSE'}
        return all(word.upper() in allowed for word in words)
    except (ValueError, UnsupportedModelError):
        return False


def _invalidate(image, table_id, name, expression, data_type, *, record_counts=None):
    semantic, dbtype, stored, xmtype = _TYPES[data_type]
    _check_index_dependencies(image, table_id, name)
    dim_file, dim_doc, dimension = next((f, doc, d) for f, doc, d in _dimensions(image)
                                      if _text(_child(d, 'ID')) == table_id)
    attribute = next(a for a in _elements(_child(dimension, 'Attributes'), 'Attribute')
                     if _text(_child(a, 'ID')) == name)
    inferred = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'InferredDatatype']
    for old in inferred:
        attribute.removeChild(old)
    inferred_node = attribute.ownerDocument.createElementNS(
        'http://schemas.microsoft.com/analysisservices/2010/engine/200', 'ddl200:InferredDatatype')
    inferred_node.appendChild(attribute.ownerDocument.createTextNode(semantic))
    attribute.insertBefore(inferred_node, _child(attribute, 'KeyColumns'))
    for prop, value in [('ProcessingState', 'Unprocessed'),
                        ('AttributeHierarchyProcessingState', 'Unprocessed')]:
        _set(attribute, prop, value)
    # A stored label sorted by this calculated key retains its values, but its
    # semantic attribute must be unprocessed so native recalculation rebuilds
    # the order. Hierarchy state alone leaves the stored attribute unusable.
    sort_dependents = {name}
    for dependent in _elements(_child(dimension, 'Attributes'), 'Attribute'):
        if any(_text(n) == name for n in dependent.childNodes
               if getattr(n, 'localName', None) == 'OrderByAttributeID'):
            sort_dependents.add(_text(_child(dependent, 'ID')))
            _set(dependent, 'ProcessingState', 'Unprocessed')
            _set(dependent, 'AttributeHierarchyProcessingState', 'Unprocessed')
    for key in _elements(attribute, 'KeyColumn') + _elements(attribute, 'NameColumn'):
        _bind(key, expression, 'WChar' if key.localName == 'NameColumn' else semantic)
    replacements = {}
    raw_file = next(f for f in image.files if f.logical_path and
        f.logical_path.replace('/', '\\').split('\\')[-1].startswith(table_id + '.') and f.logical_path.endswith('.tbl.xml'))
    raw_doc = minidom.parseString(image.read_logical_file(raw_file))
    raw = next(n for n in _elements(raw_doc, 'XMObject')
               if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == name)
    segments = next(c for c in _elements(raw, 'Collection') if _text(_child(c, 'Name')) == 'Segments')
    counts = [int(_text(_child(_child(s, 'Properties'), 'Records'))) for s in segments.childNodes
              if getattr(s, 'localName', None) == 'XMObject']
    if record_counts is not None:
        counts = list(record_counts)
    if not counts or any(n < 0 or n > 0xffffffff for n in counts):
        raise UnsupportedModelError('Invalid calculated column segment record counts')
    from .schema import column_file_matcher
    owns = column_file_matcher(table_id, name, [_text(_child(a, 'ID')) for a in
                               _elements(_child(dimension, 'Attributes'), 'Attribute')])
    # The partition stream only; intrinsic-hierarchy H$ streams are also .idf.
    idf_files = [f for f in image.files if f.logical_path and
        owns(f.logical_path.split('\\')[-1]) and f.logical_path.endswith('.idf')
        and f'.{table_id}.{name}.' in f.logical_path.split('\\')[-1]]
    if len(idf_files) != 1:
        raise UnsupportedModelError('Calculated columns require one partition stream')
    idf = idf_files[0]
    idf_name = idf.logical_path.split('\\')[-1]
    locale = _text(_child(_child(raw, 'Properties'), 'Locale'))
    replacement = _raw_column(raw_doc, name, counts, idf_name, locale, dbtype, stored, xmtype)
    # Native H39 retains the sort binding while the calculated intrinsic
    # hierarchy is unprocessed; Excel builds its order during calculation.
    sort_keys = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'OrderByAttributeID']
    if sort_keys and _text(sort_keys[0]):
        _set(_child(replacement, 'Properties'), 'OrderByColumn', _text(sort_keys[0]))
    raw.parentNode.replaceChild(replacement, raw)
    _set(_child(raw_doc.documentElement, 'Properties'), 'Settings', 4097)
    replacements[raw_file.storage_path] = raw_doc.documentElement.toxml(encoding='utf-8')
    replacements[idf.storage_path] = b''.join(_blank_segment(n) for n in counts)
    removed = {f.storage_path for f in image.files if f.logical_path and f.storage_path != idf.storage_path and
        owns(f.logical_path.split('\\')[-1])}
    # Native H31B persists hierarchies over calculated fields without U$ lookup
    # storage until Excel calculates. Preserve levels/captions/extensions while
    # discarding only obsolete traversal caches and marking their state explicit.
    affected_hierarchies = [h for h in _elements(_child(dimension, 'Hierarchies'), 'Hierarchy')
                           if any(_text(r) in sort_dependents for r in _elements(h, 'SourceAttributeID'))]
    for hierarchy in affected_hierarchies:
        hierarchy_id = _text(_child(hierarchy, 'ID'))
        levels = [_text(r) for r in _elements(hierarchy, 'SourceAttributeID')]
        _set(hierarchy, 'ProcessingState', 'Unprocessed')
        _set(hierarchy, 'StructureType', 'Unknown')
        raw_hierarchy = next((n for n in _elements(raw_doc, 'XMObject')
            if n.getAttribute('class') == 'XMUserHierarchy' and n.getAttribute('name') == hierarchy_id), None)
        if raw_hierarchy is None:
            raise UnsupportedModelError('Missing raw user hierarchy: ' + hierarchy_id)
        properties = _child(raw_hierarchy, 'Properties')
        _set(properties, 'IsProcessed', 'false')
        _set(properties, 'TableStore', ''.join(f'${level}$-1' for level in levels) + '$')
        _set(properties, 'TableName', '')
        stem = f'U${table_id}${hierarchy_id}.'
        removed.update(f.storage_path for f in image.files if f.logical_path and
            f.logical_path.replace('\\', '/').rsplit('/', 1)[-1].startswith((stem, '2.' + stem)))
    replacements[raw_file.storage_path] = raw_doc.documentElement.toxml(encoding='utf-8')
    removed_names = {f.logical_path.split('\\')[-1] for f in image.files if f.storage_path in removed}
    _set(dimension, 'DataFileList', ';'.join(n for n in _text(_child(dimension, 'DataFileList')).split(';') if n not in removed_names))
    replacements[dim_file.storage_path] = dim_doc.documentElement.toxml(encoding='utf-8')
    for f in image.files:
        if not f.logical_path or not f.logical_path.endswith('.det.xml'):
            continue
        doc = minidom.parseString(image.read_logical_file(f))
        changed = False
        for d in _elements(doc, 'Dimension'):
            ids = [n for n in d.childNodes if getattr(n, 'localName', None) == 'CubeDimensionID']
            if not ids or _text(ids[0]) != table_id:
                continue
            for attr in _elements(d, 'Attribute'):
                if _text(_child(attr, 'AttributeID')) == name:
                    for key in _elements(attr, 'KeyColumn'):
                        _bind(key, expression, semantic)
                    changed = True
        if changed:
            replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    # Native Excel accepts retained relationship lookup storage when its
    # semantic relationship is explicitly unprocessed. ProcessRecalc rebuilds
    # it from the newly calculated keys before the model can be queried.
    for file, _, _ in _dimensions(image):
        doc = minidom.parseString(replacements.get(file.storage_path, image.read_logical_file(file)))
        changed = False
        for relationship in _elements(doc, 'Relationship'):
            ends = (_elements(relationship, 'FromRelationshipEnd') +
                    _elements(relationship, 'ToRelationshipEnd'))
            if any(_text(_child(end, 'DimensionID')) == table_id and
                   any(_text(node) == name for node in _elements(end, 'AttributeID')) for end in ends):
                _set(relationship, 'ProcessingState', 'Unprocessed')
                changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements, removals=removed))



def _check_index_dependencies(image, table_id, name):
    # Relationships, sorts and user hierarchies are deferred to native processing.
    # Date keys remain guarded: native testing found a key/calculation cycle.
    for file, _, dimension in _dimensions(image):
        own_table = _text(_child(dimension, 'ID')) == table_id
        if own_table:
            for attr in _elements(_child(dimension, 'Attributes'), 'Attribute'):
                if _text(_child(attr, 'ID')) == name:
                    if _text(_child(attr, 'Usage')) == 'Key':
                        raise UnsupportedModelError('Calculated date-table keys cannot be invalidated')


def _bind(key, expression, data_type):
    for prop, value in [('DataType', data_type), ('DataSize', 0), ('NullProcessing', 'Automatic')]:
        _set(key, prop, value)
    source = next((n for n in key.childNodes if getattr(n, 'localName', None) == 'Source'), None)
    if source is None:
        source = key.appendChild(key.ownerDocument.createElement('Source'))
    while source.firstChild:
        source.removeChild(source.firstChild)
    source.setAttribute('xsi:type', 'ddl200_200:ExpressionBinding')
    _set(source, 'Expression', expression)


def _blank_segment(records):
    # Native zero-row columns retain allocated slots but contain no BLANK run.
    # A zero-length run with data ID 2 prevents Excel loading an empty model.
    if records == 0:
        return struct.pack('<Q', 16) + bytes(128) + struct.pack('<QQ', 1, 0)
    return struct.pack('<QII', 16, 2, records) + bytes(120) + struct.pack('<QQ', 1, 0)


def _raw_column(doc, name, counts, idf, locale, dbtype, stored, xmtype):
    def element(tag, parent=None, text=None, **attrs):
        e = doc.createElement(tag)
        for k, v in attrs.items():
            e.setAttribute(k, str(v))
        if text is not None:
            e.appendChild(doc.createTextNode(str(text)))
        if parent is not None:
            parent.appendChild(e)
        return e
    def obj(cls, parent=None, name=None, version=0):
        e = element('XMObject', parent, **{'class': cls, 'ProviderVersion': version})
        if name is not None:
            e.setAttribute('name', name)
        return e
    def props(node, values):
        p = element('Properties', node)
        for key, (value, dtype) in values.items():
            e = element(key, p, value)
            if dtype:
                e.setAttribute('xsi:type', 'xsd:' + dtype)
        return p
    def member(parent, label, cls):
        members = next((n for n in parent.childNodes if getattr(n, 'localName', None) == 'Members'), None)
        if members is None:
            members = element('Members', parent)
        m = element('Member', members)
        element('Name', m, label)
        return obj(cls, m)
    def stats(node, records, *, sub=False, total=False):
        values = {'DistinctStates': (0, 'int' if total else 'long'), 'MinDataID': (2, 'int'),
            'MaxDataID': (2, 'int'), 'OriginalMinSegmentDataID': (2, 'int'), 'RLESortOrder': (-1, 'long'),
            'RowCount': (records, 'long'), 'HasNulls': ('false' if sub or not records else 'true', 'boolean'),
            'RLERuns': (0 if sub or total or not records else 1, 'long'), 'OthersRLERuns': (0, 'long')}
        if total:
            values.update({'Usage': (3, 'int'), 'DBType': (dbtype, 'short'), 'XMType': (xmtype, 'int'),
                'CompressionType': (0, 'int'), 'CompressionParam': (0, 'long'), 'EncodingHint': (0, 'int'),
                'AggCounter': (0, 'long'), 'WhereCounter': (0, 'long'), 'OrderByCounter': (0, 'long')})
        props(node, values)
    raw = obj('XMRawColumn', name=name, version=1)
    props(raw, {'Settings': (3073, 'long'), 'ColumnFlags': (8, 'long'), 'Collation': (None, None),
        'OrderByColumn': (None, None), 'Locale': (locale, 'long'), 'BinaryCharacters': (0, 'unsignedInt')})
    props(member(raw, 'IntrinsicHierarchy', 'XMHierarchy'), {'SortOrder': (0, 'int'), 'IsProcessed': ('false', 'boolean'),
        'TypeMaterialization': (-1, 'int'), 'ColumnPosition2DataID': (-1, 'long'),
        'ColumnDataID2Position': (-1, 'long'), 'DistinctDataIDs': (-1, 'long'), 'TableStore': (None, None)})
    stats(member(raw, 'ColumnStats', 'XMColumnStats'), sum(counts), total=True)
    collection = element('Collection', element('Collections', raw))
    element('Name', collection, 'Segments')
    for records in counts:
        segment = obj('XMColumnSegment', collection)
        props(segment, {'Records': (records, 'long'), 'Mask': (0, 'long')})
        sub = member(segment, 'SubSegment', 'XMColumnSegment')
        props(sub, {'Records': (0, 'long'), 'Mask': (0, 'long')})
        props(member(sub, 'CompressionInfo', 'XMRENoSplitCompressionInfo<1>'), {'Min': (2, 'int')})
        stats(member(sub, 'ColumnSegmentStats', 'XMColumnSegmentStats'), 0, sub=True)
        compression = member(segment, 'CompressionInfo', 'XMHybridRLECompressionInfo<class XMRENoSplitCompressionInfo<1>>')
        props(member(compression, 'RLECompression', 'XMRLECompressionInfo'), {
            'BookmarkBits': (max(1, records).bit_length() - 1, 'long'), 'StorageAllocSize': (32, 'long'),
            'StorageUsedSize': (4 if records else 0, 'long'), 'SegmentNeedsResizing': ('false', 'boolean')})
        props(member(compression, 'SubCompression', 'XMRENoSplitCompressionInfo<1>'), {'Min': (2, 'int')})
        stats(member(segment, 'ColumnSegmentStats', 'XMColumnSegmentStats'), records)
    data = element('DataObjects', raw)
    dictionary = obj('XMHashDataDictionary<XM_' + stored + '>', element('DataObject', data))
    properties = {'DataVersion': (0, 'int'), 'LastId': (2, 'int'), 'Nullable': ('true', 'boolean'), 'Unique': ('false', 'boolean')}
    if stored == 'String':
        properties['DictionaryFlags'] = (3, 'long')
    elif stored == 'Long':
        properties['OperatingOn32'] = ('true', 'boolean')
    props(dictionary, properties)
    props(obj('XMRawColumnPartitionDataObject', element('DataObject', data), name=idf),
          {'DataVersion': (0, 'int'), 'Partition': (0, 'int'), 'SegmentCount': (len(counts), 'int')})
    return raw
