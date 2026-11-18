"""Relationship semantic bindings and dense foreign-ID lookup storage."""
from .collation import text_key
from pathlib import Path
from xml.dom import minidom
import struct
import uuid
from dataclasses import replace
import pandas as pd

from .abf import AbfImage, UnsupportedModelError, VirtualFileAddition
from .schema import _elements, _child, _text
from .relationships import add_relationships_to_dimension
from .vertipaq import select_bit_width


def stored_relationships(model):
    from .relationships import Relationship
    captions = {value: key for key, value in model._dimension_ids.items()}
    result = []
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):
            continue
        doc = minidom.parseString(model._image.read_logical_file(file))
        for node in _elements(doc, 'Relationship'):
            ends = [_child(node, name) for name in ('FromRelationshipEnd', 'ToRelationshipEnd')]
            arguments = []
            for end in ends:
                table = captions[_text(_child(end, 'DimensionID'))]
                identity = _text(_elements(end, 'AttributeID')[0])
                column = next(name for (owner, name), value in model._column_ids.items()
                              if owner == table and value == identity)
                arguments.extend([table, column])
            result.append((_text(_child(node, 'ID')), Relationship(*arguments, active=_text(_child(node, 'Visible')) == 'true', cardinality='one-to-one' if _text(_child(ends[0], 'Multiplicity')) == 'One' else 'many-to-one')))
    return result


def relationship_key_type(model, table, column):
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths[table]))
    attribute = next(n for n in _elements(doc, 'Attribute')
                     if (_text(_child(n, 'Name')) or _text(_child(n, 'ID'))) == column)
    datatype = _text(_child(_elements(attribute, 'KeyColumn')[0], 'DataType'))
    if datatype == 'Empty' and _elements(attribute, 'Expression'):
        datatype = _text(_child(attribute, 'InferredDatatype'))
    return datatype


def validate_cached_relationship_keys(model, relationship):
    """Validate available keys; return whether Excel must validate a pending key."""
    from .relationships import validate_relationship_keys
    pending = {(table, c.name) for table, identity in model._dimension_ids.items()
               for c in model.list_calculated_columns()
               if c.table_id == identity and c.processing_state == 'Unprocessed'}
    ends = [(relationship.to_table, relationship.to_column)]
    if relationship.cardinality == 'one-to-one':
        ends.append((relationship.from_table, relationship.from_column))
    for table, column in ends:
        if (table, column) in pending:
            continue
        check = replace(relationship, to_table=table, to_column=column, cardinality='many-to-one')
        validate_relationship_keys(check, [], model.get_table(table)[column].tolist())
    return bool(pending.intersection({(relationship.from_table, relationship.from_column),
                                      (relationship.to_table, relationship.to_column)}))


def synchronize_unknown_members(model):
    """Maintain the legacy engine's virtual blank-member presence metadata.

    A dense relationship index can return the unknown row even when its lookup
    table says it has no RI violations. Ordinary attribute queries tolerate
    that inconsistency, but user hierarchies then fail with member-not-found.
    A nonzero RIViolationCount enables the virtual member; it must propagate
    through lookup chains and disappear once the underlying violation is gone.
    Inactive relationships retain their storage and unknown-member contract.
    """
    stored = stored_relationships(model)
    relationships = [r for _, r in stored]
    violations = {}
    required = set()
    if relationships:
        tables = {t for r in relationships for t in (r.from_table, r.to_table)}
        with model._open_reader() as reader:
            frames = {t: reader.get_table(t) for t in tables}
        def canonical(value):
            return text_key(value) if isinstance(value, str) else value
        for relationship_id, relationship in stored:
            primary = {canonical(v) for v in frames[relationship.to_table][relationship.to_column]
                       if not pd.isna(v)}
            foreign = frames[relationship.from_table][relationship.from_column]
            violations[relationship_id] = any(pd.isna(v) or canonical(v) not in primary for v in foreign)
            if violations[relationship_id]:
                required.add(relationship.to_table)
        # A virtual blank in an intermediate lookup has a blank foreign key.
        # Fixed-point propagation also terminates for inactive cycles.
        while True:
            propagated = required | {r.to_table for r in relationships if r.from_table in required}
            if propagated == required:
                break
            required = propagated
    captions = {physical: caption for caption, physical in model._dimension_ids.items()}
    replacements = {}
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.endswith('.tbl.xml'):
            continue
        doc = minidom.parseString(model._image.read_logical_file(file))
        table = captions.get(doc.documentElement.getAttribute('name'))
        if table is None:
            continue
        node = _child(_child(doc.documentElement, 'Properties'), 'RIViolationCount')
        current = int(_text(node))
        # Preserve an existing native positive count when presence is unchanged.
        value = max(1, current) if table in required else 0
        changed = value != current
        if value != current:
            node.firstChild.data = str(value)
        # Native relationship creation sets bit zero for a direct RI violation.
        # Without this flag, later Power Pivot edits recalculate the table's
        # count as zero and the hierarchy failure returns.
        for relationship in _elements(doc, 'XMObject'):
            if relationship.getAttribute('class') != 'XMRelationship':
                continue
            relationship_id = relationship.getAttribute('name')
            if relationship_id not in violations:
                continue
            dense = next((n for n in _elements(relationship, 'XMObject')
                          if n.getAttribute('class') == 'XMRelationshipIndexDenseDIDs'), None)
            if dense is None:
                continue
            flag = _child(_child(dense, 'Properties'), 'Flags')
            previous = int(_text(flag))
            updated = (previous & ~1) | int(violations[relationship_id])
            if updated != previous:
                flag.firstChild.data = str(updated)
                changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(model._image.assemble(replacements)) if replacements else model._image


def synchronize_indirect_bindings(model):
    """Bind every reachable lookup dimension through its immediate predecessor.

    Excel materializes direct edges as Regular and transitive edges as Indirect.
    Indirect bindings have no RelationshipID: the intermediate dimension owns
    that edge. Recompute them after graph edits so removals cannot leave orphans.
    """
    image = model._image
    documents, edges = {}, {}
    for table, physical in model._dimension_ids.items():
        file = next(f for f in image.files if f.logical_path and
                    f.logical_path.split('\\')[-1].startswith(physical + '.') and
                    f.logical_path.endswith('.det.xml'))
        doc = minidom.parseString(image.read_logical_file(file))
        dimensions = _elements(doc, 'Dimensions')[0]
        documents[physical] = (file, doc, dimensions)
        edges[physical] = []
        for node in list(dimensions.childNodes):
            if node.nodeType != node.ELEMENT_NODE or node.getAttribute('xsi:type') != 'ReferenceMeasureGroupDimension':
                continue
            if _text(_child(node, 'Materialization')) == 'Indirect':
                dimensions.removeChild(node)
            else:
                edges[physical].append((_text(_child(node, 'CubeDimensionID')), node))
    # A newly added lookup column must be present in every reference dimension,
    # not just in its own table. Preserve each edge's granularity key.
    for bindings in edges.values():
        for target, binding in bindings:
            source_dimension = next(n for n in _elements(documents[target][1], 'Dimension')
                                    if _text(_child(n, 'CubeDimensionID')) == target)
            previous = _child(binding, 'Attributes')
            granularity = {_text(_child(n, 'AttributeID')) for n in _elements(previous, 'Attribute')
                           if _text(_child(n, 'Type')) == 'Granularity'}
            attributes = binding.ownerDocument.importNode(_child(source_dimension, 'Attributes'), True)
            for attribute in _elements(attributes, 'Attribute'):
                is_key = _text(_child(attribute, 'AttributeID')) in granularity
                _child(attribute, 'Type').firstChild.data = 'Granularity' if is_key else 'Regular'
                if is_key:
                    _elements(attribute, 'NullProcessing')[0].firstChild.data = 'Error'
            binding.replaceChild(attributes, previous)
    replacements = {}
    for physical, (file, doc, dimensions) in documents.items():
        seen = {physical}
        queue = [(physical, 0)]
        for current, depth in queue:
            for target, binding in edges.get(current, []):
                if target in seen:
                    raise UnsupportedModelError('Active relationships contain a cycle or ambiguous lookup path')
                seen.add(target)
                queue.append((target, depth + 1))
                if depth == 0:
                    continue
                indirect = doc.importNode(binding, True)
                _child(indirect, 'Materialization').firstChild.data = 'Indirect'
                for node in list(_elements(indirect, 'RelationshipID')):
                    node.parentNode.removeChild(node)
                dimensions.appendChild(indirect)
        data = doc.documentElement.toxml(encoding='utf-8')
        if data != image.read_logical_file(file):
            replacements[file.storage_path] = data
    return AbfImage(image.assemble(replacements)) if replacements else image


def remove_relationship_storage(model, relationship_id):
    image = model._image
    removed = next((r for rid, r in stored_relationships(model) if rid == relationship_id), None)
    remaining = [r for rid, r in stored_relationships(model) if rid != relationship_id]
    released = set()
    if removed is not None and not any((r.to_table, r.to_column) ==
            (removed.to_table, removed.to_column) for r in remaining):
        released.add((model._dimension_ids[removed.to_table], model._column_ids[(removed.to_table, removed.to_column)]))
    replacements, removals = {}, []
    for file in image.files:
        if not file.logical_path:
            continue
        if f'${relationship_id}.' in file.logical_path:
            removals.append(file.storage_path)
            continue
        if not file.logical_path.endswith(('.dim.xml', '.det.xml', '.tbl.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        for node in list(_elements(doc, 'Relationship')):
            if _text(_child(node, 'ID')) == relationship_id:
                node.parentNode.removeChild(node)
                changed = True
        for node in list(_elements(doc, 'XMObject')):
            if node.getAttribute('class') == 'XMRelationship' and node.getAttribute('name') == relationship_id:
                node.parentNode.removeChild(node)
                changed = True
        for node in list(_elements(doc, 'RelationshipID')):
            if _text(node) == relationship_id:
                dimension = node.parentNode
                dimension.parentNode.removeChild(dimension)
                changed = True
        for node in _elements(doc, 'DataFileList'):
            names = _text(node).split(';')
            retained = [name for name in names if f'${relationship_id}.' not in name]
            if names != retained:
                node.firstChild.data = ';'.join(retained)
                changed = True
        # The final relationship no longer requires a unique, nonblank lookup
        # key. Retaining its contract would reject otherwise valid future edits.
        filename = file.logical_path.split('\\')[-1]
        for physical, column in released:
            if not filename.startswith(physical + '.'):
                continue
            if filename.endswith('.dim.xml'):
                for attribute in _elements(doc, 'Attribute'):
                    identifiers = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'ID']
                    if identifiers and _text(identifiers[0]) == column:
                        _elements(attribute, 'NullProcessing')[0].firstChild.data = 'Preserve'
                        changed = True
                for binding in _elements(doc, 'AttributeRelationship'):
                    if _text(_child(binding, 'AttributeID')) == column:
                        _child(binding, 'Cardinality').firstChild.data = 'Many'
                        changed = True
            elif filename.endswith('.tbl.xml'):
                for raw in _elements(doc, 'XMObject'):
                    if raw.getAttribute('class') == 'XMRawColumn' and raw.getAttribute('name') == column:
                        flags = _child(_child(raw, 'Properties'), 'ColumnFlags')
                        flags.firstChild.data = str(int(_text(flags)) & ~3)
                        changed = True
        if removed is not None and filename.endswith('.tbl.xml'):
            affected = [(removed.from_table, removed.from_column), (removed.to_table, removed.to_column)]
            for table, column in affected:
                if not filename.startswith(model._dimension_ids[table] + '.'):
                    continue
                usage = 0 if any((r.to_table, r.to_column) == (table, column) for r in remaining) else (
                    1 if any((r.from_table, r.from_column) == (table, column) for r in remaining) else 3)
                for raw in _elements(doc, 'XMObject'):
                    if raw.getAttribute('class') == 'XMRawColumn' and raw.getAttribute('name') == model._column_ids[(table, column)]:
                        stats = next(n for n in _elements(raw, 'XMObject') if n.getAttribute('class') == 'XMColumnStats')
                        _child(_child(stats, 'Properties'), 'Usage').firstChild.data = str(usage)
                        changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements, removals=removals))


def add_relationship_storage(model, relationship, *, existing_id=None, _deferred_calculated=None):
    if _deferred_calculated is None:
        calculated = {(table, c.name) for table, identity in model._dimension_ids.items()
                      for c in model.list_calculated_columns() if c.table_id == identity}
        _deferred_calculated = bool(calculated.intersection({
            (relationship.from_table, relationship.from_column),
            (relationship.to_table, relationship.to_column)}))
    if _deferred_calculated:
        from .calculated_columns import invalidate_calculated_columns
        model = type(model)(model.source, _image=invalidate_calculated_columns(model._image))
    image = model._image
    foreign = model._dimension_ids[relationship.from_table]
    primary = model._dimension_ids[relationship.to_table]
    foreign_column = model._column_ids[(relationship.from_table, relationship.from_column)]
    primary_column = model._column_ids[(relationship.to_table, relationship.to_column)]
    if relationship_key_type(model, relationship.from_table, relationship.from_column) != relationship_key_type(model, relationship.to_table, relationship.to_column):
        raise ValueError('Relationship columns must have the same data type')
    with model._open_reader() as reader:
        fk = reader.get_table(relationship.from_table)[relationship.from_column].tolist()
        pk = reader.get_table(relationship.to_table)[relationship.to_column].tolist()
    if foreign == primary:
        raise ValueError('Self relationships are not supported')
    def canonical(value):
        return text_key(value) if isinstance(value, str) else value
    from .relationships import validate_relationship_keys, validate_relationship_collation
    validate_relationship_collation(model, relationship)
    if not _deferred_calculated:
        validate_relationship_keys(relationship, fk, pk)
    encoding = model._encodings.get((relationship.from_table, relationship.from_column))
    ids = [2] * len(fk) if encoding is None and _deferred_calculated else encoding.encode(fk)
    if isinstance(ids, tuple):
        ids = ids[0]
    count = max(ids, default=2) + 1
    if count > 4_000_000:
        raise UnsupportedModelError('Relationship requires a sparse index; dense ID range exceeds 4 million')
    positions = {canonical(value): index for index, value in enumerate(pk)}
    index_values = [-1] * count
    for value, data_id in zip(fk, ids):
        index_values[data_id] = -1 if pd.isna(value) else positions.get(canonical(value), -1)
    def file(kind, table):
        return next(f for f in image.files if f.logical_path and f.logical_path.split('\\')[-1].startswith(table + '.') and f.logical_path.endswith('.' + kind + '.xml'))
    if existing_id is not None:
        stem = f'R${foreign}${existing_id}'
        matching = [f for f in image.files if f.logical_path and stem in f.logical_path.split('\\')[-1]]
        index_file = next((f for f in matching if f.logical_path.endswith('.idf')), None)
        meta_file = next((f for f in matching if f.logical_path.endswith('.tbl.xml')), None)
        if index_file is None or meta_file is None:
            raise UnsupportedModelError('Cannot resolve relationship lookup storage')
        width = select_bit_width(len(pk))
        slots = 64 // width
        words = [sum((v + 1) << (width * slot) for slot, v in enumerate(index_values[start:start + slots]))
                 for start in range(0, count, slots)]
        if count % slots == 0:
            words.append(0)
        replacements = {index_file.storage_path: struct.pack('<Q', len(words)) + b''.join(struct.pack('<Q', w) for w in words)}
        for target in [meta_file, file('tbl', foreign)]:
            doc = minidom.parseString(image.read_logical_file(target))
            if target == meta_file:
                scope = doc
                codecs = [n for n in _elements(doc, 'XMObject') if n.getAttribute('class').startswith('XMRENoSplitCompressionInfo')]
                if len(codecs) != 1:
                    raise UnsupportedModelError('Unsupported relationship index compression')
                codecs[0].setAttribute('class', f'XMRENoSplitCompressionInfo<{width}>')
            else:
                scope = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRelationship' and n.getAttribute('name') == existing_id)
                if not any(n.getAttribute('class') == 'XMRelationshipIndexDenseDIDs' for n in _elements(scope, 'XMObject')):
                    raise UnsupportedModelError('Unsupported relationship index layout')
            for node in _elements(scope, 'Records'):
                node.firstChild.data = str(count)
            replacements[target.storage_path] = doc.documentElement.toxml(encoding='utf-8')
        return AbfImage(image.assemble(replacements))
    dim = file('dim', foreign)
    dim_data = add_relationships_to_dimension(image.read_logical_file(dim), from_dimension_id=foreign,
        relationships=[(replace(relationship, from_column=foreign_column, to_column=primary_column), primary)])
    dim_doc = minidom.parseString(dim_data)
    relation = _elements(dim_doc, 'Relationship')[-1]
    if _deferred_calculated:
        _child(relation, 'ProcessingState').firstChild.data = 'Unprocessed'
    rid = _text(_child(relation, 'ID'))
    stem = f'R${foreign}${rid}'
    index_filename, meta_filename = '0.' + stem + '.INDEX.0.idf', stem + '.0.tbl.xml'
    _elements(dim_doc, 'DataFileList')[0].firstChild.data += ';' + index_filename + ';' + meta_filename
    replacements = {dim.storage_path: dim_doc.documentElement.toxml(encoding='utf-8')}
    primary_dim = file('dim', primary)
    primary_doc = minidom.parseString(image.read_logical_file(primary_dim))
    for attribute in _elements(primary_doc, 'Attribute'):
        identifiers = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'ID']
        if identifiers and _text(identifiers[0]) == primary_column:
            _elements(attribute, 'NullProcessing')[0].firstChild.data = 'Error'
    for binding in _elements(primary_doc, 'AttributeRelationship'):
        if _text(_child(binding, 'AttributeID')) == primary_column:
            _child(binding, 'Cardinality').firstChild.data = 'One'
    replacements[primary_dim.storage_path] = primary_doc.documentElement.toxml(encoding='utf-8')
    namespace = 'http://schemas.microsoft.com/analysisservices/imbi'
    def append(parent, name, value=None, *, attributes=None):
        doc = parent.ownerDocument
        node = doc.createElement(name)
        for key, content in (attributes or {}).items():
            node.setAttribute(key, content)
        if value is not None:
            node.appendChild(doc.createTextNode(str(value)))
        parent.appendChild(node)
        return node
    for table, column, usage in [(foreign, foreign_column, 1), (primary, primary_column, 0)]:
        target = file('tbl', table)
        doc = minidom.parseString(image.read_logical_file(target))
        raw = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == column)
        stats = next(n for n in _elements(raw, 'XMObject') if n.getAttribute('class') == 'XMColumnStats')
        _child(_child(stats, 'Properties'), 'Usage').firstChild.data = str(usage)
        if table == primary:
            flags = _child(_child(raw, 'Properties'), 'ColumnFlags')
            flags.firstChild.data = str(int(_text(flags)) | 3)
        else:
            collection = next(n for n in _elements(doc, 'Collection') if _text(_child(n, 'Name')) == 'Relationships')
            obj = append(collection, 'XMObject', attributes={'class':'XMRelationship', 'name':rid, 'ProviderVersion':'0'})
            props = append(obj, 'Properties')
            for name, value in [('PrimaryTable', primary), ('PrimaryColumn', primary_column), ('ForeignColumn', foreign_column)]:
                append(props, name, value)
            data = append(append(obj, 'DataObjects'), 'DataObject')
            dense = append(data, 'XMObject', attributes={'class':'XMRelationshipIndexDenseDIDs','ProviderVersion':'0'})
            props = append(dense, 'Properties')
            append(props, 'Records', count, attributes={'xsi:type':'xsd:long'})
            append(props, 'TableName', stem)
            append(props, 'Flags', 0, attributes={'xsi:type':'xsd:long'})
        replacements[target.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    # Bind the lookup table into the fact table's measure group.
    det, primary_det = file('det', foreign), file('det', primary)
    doc = minidom.parseString(image.read_logical_file(det))
    source = minidom.parseString(image.read_logical_file(primary_det))
    dimensions = _elements(doc, 'Dimensions')[0]
    reference = append(dimensions, 'Dimension', attributes={'xsi:type':'ReferenceMeasureGroupDimension'})
    source_dim = next(n for n in _elements(source, 'Dimension') if _text(_child(n, 'CubeDimensionID')) == primary)
    attrs = doc.importNode(_child(source_dim, 'Attributes'), True)
    reference.appendChild(attrs)
    for attr in [n for n in attrs.childNodes if n.nodeType == n.ELEMENT_NODE]:
        name = _text(_child(attr, 'AttributeID'))
        _child(attr, 'Type').firstChild.data = 'Granularity' if name == primary_column else 'Regular'
        if name == primary_column:
            _elements(attr, 'NullProcessing')[0].firstChild.data = 'Error'
    for key, value in [('IntermediateCubeDimensionID',foreign), ('IntermediateGranularityAttributeID',foreign_column),
                        ('Materialization','Regular'), ('ddl200:ProcessingState','Processed'),
                        ('ddl300:RelationshipID',rid), ('CubeDimensionID',primary)]:
        append(reference, key, value)
    if not relationship.active:
        dimensions.removeChild(reference)
    replacements[det.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    template = (Path(__file__).parent / 'data/relationship-index.xml').read_text(encoding='utf-8')
    doc = minidom.parseString(template.replace('RELATIONSHIP_INDEX', stem))
    for node in _elements(doc, 'Records'):
        node.firstChild.data = str(count)
    width = select_bit_width(len(pk))
    codec = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class').startswith('XMRENoSplitCompressionInfo'))
    codec.setAttribute('class', f'XMRENoSplitCompressionInfo<{width}>')
    slots = 64 // width
    words = [sum((v + 1) << (width * slot) for slot, v in enumerate(index_values[start:start + slots]))
             for start in range(0, count, slots)]
    if count % slots == 0:
        words.append(0)
    binary = struct.pack('<Q', len(words)) + b''.join(struct.pack('<Q', word) for word in words)
    anchor = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.idf')
                  and f'.{foreign}.{foreign_column}.' in f.logical_path.split('\\')[-1])
    parent = anchor.logical_path.rsplit('\\', 1)[0]
    trailer = image.read_file(anchor)[-4:] if image.error_code else None
    additions = [VirtualFileAddition(uuid.uuid4().hex.upper()[:20], parent + '\\' + name, anchor.storage_path, data, trailer)
        for name, data in [(index_filename, binary), (meta_filename, doc.documentElement.toxml(encoding='utf-8'))]]
    return AbfImage(image.assemble(replacements, additions=additions))
