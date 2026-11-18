"""User hierarchy metadata and breadth-first traversal indexes."""
from pathlib import Path
from xml.dom import minidom
from xml.sax.saxutils import escape
import re
import struct
import uuid

from .abf import AbfImage, VirtualFileAddition
from .schema import _elements, _child, _text, storage_identity


_HIERARCHY_STREAMS = ('MULTI_LEVEL_ID', 'PARENT_POS', 'FIRST_CHILD_POS', 'CHILD_COUNT')


def _owned_file_pattern(stem):
    """Match only this stem's native storage names.

    Names may contain '.', so a prefix test would also claim the files of a
    sibling such as ``A.B`` when ``A`` is targeted.
    """
    stem = re.escape(stem)
    return re.compile(rf'(?:{stem}\.\d+\.tbl\.xml|2\.{stem}\.(?:{"|".join(_HIERARCHY_STREAMS)})\.\d+\.idf)')


def _attribute_meta_pattern(stem):
    return re.compile(re.escape(stem) + r'\.\d+\.tbl\.xml')


def resolve_hierarchy(model, table, name):
    """Resolve a public hierarchy caption to its stable storage identity."""
    doc = minidom.parseString(model._image.read_logical_file(model._dimension_paths[table]))
    members = [(_text(_child(h, 'ID')), _text(_child(h, 'Name')) or _text(_child(h, 'ID')))
               for h in _elements(_elements(doc, 'Hierarchies')[0], 'Hierarchy')]
    for identity, caption in members:
        if caption == name:
            return identity, caption
    for identity, caption in members:
        if identity == name:
            return identity, caption
    raise KeyError(f'Hierarchy {table}.{name} does not exist')


def stored_hierarchies(model):
    result = []
    for table, file in model._dimension_paths.items():
        doc = minidom.parseString(model._image.read_logical_file(file))
        for hierarchy in _elements(_elements(doc, 'Hierarchies')[0], 'Hierarchy'):
            result.append((table, _text(_child(hierarchy, 'Name')) or _text(_child(hierarchy, 'ID')),
                           [next(caption for (owner, caption), identity in model._column_ids.items()
                                 if owner == table and identity == _text(n))
                            for n in _elements(hierarchy, 'SourceAttributeID')]))
    return result


def traversal_indexes(level_ids, orders=None):
    """Flatten distinct paths by depth, keeping siblings adjacent."""
    if not level_ids or len({len(ids) for ids in level_ids}) != 1:
        raise ValueError('Hierarchy levels must have equal row counts')
    orders = orders or [{} for _ in level_ids]
    def order(path):
        return tuple(orders[i].get(value, value) for i, value in enumerate(path))
    paths = sorted(set(zip(*level_ids)), key=order)
    nodes, previous, offsets = [], {(): -1}, []
    for depth in range(len(level_ids)):
        offsets.append(len(nodes))
        current = {}
        for path in sorted({p[:depth + 1] for p in paths}, key=order):
            position = len(nodes)
            current[path] = position
            parent = previous[path[:-1]]
            nodes.append([path[-1], parent, -1, -1])
            if parent >= 0:
                if nodes[parent][2] == -1:
                    nodes[parent][2] = position
                    nodes[parent][3] = 0
                nodes[parent][3] += 1
        previous = current
    return offsets, list(zip(*nodes)) if nodes else [[], [], [], []]


def intrinsic_order(model, table, column):
    """Read Excel's member positions, including custom sort-column ordering."""
    from .abf import UnsupportedModelError
    image = model._image
    stem = f'H${model._dimension_ids[table]}${model._column_ids[(table, column)]}'
    meta = next(f for f in image.files if f.logical_path and
                _attribute_meta_pattern(stem).fullmatch(f.logical_path.split('\\')[-1]))
    doc = minidom.parseString(image.read_logical_file(meta))
    raw = next(n for n in _elements(doc, 'XMObject')
               if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == 'POS_TO_ID')
    codecs = [n for n in _elements(raw, 'XMObject')
              if n.getAttribute('class').startswith('XMRENoSplitCompressionInfo')]
    if any(n.getAttribute('class') != 'XMRENoSplitCompressionInfo<32>' for n in codecs):
        raise UnsupportedModelError('Unsupported hierarchy member-position encoding')
    segments = [n for n in _elements(raw, 'XMObject') if n.getAttribute('class') == 'XMColumnSegment']
    filename = next(n.getAttribute('name') for n in _elements(raw, 'XMObject')
                    if n.getAttribute('class') == 'XMRawColumnPartitionDataObject')
    file = next(f for f in image.files if f.logical_path and
                f.logical_path.rsplit('\\', 1)[0] == meta.logical_path.rsplit('\\', 1)[0] and
                f.logical_path.split('\\')[-1] == filename)
    data, offset, result = image.read_logical_file(file), 0, {}
    for segment in segments:
        records = int(_text(_child(_child(segment, 'Properties'), 'Records')))
        words = struct.unpack_from('<Q', data, offset)[0]
        offset += 8
        if records > words * 2 or offset + words * 8 > len(data):
            raise UnsupportedModelError('Invalid hierarchy member-position stream')
        for value in struct.unpack_from('<' + 'I' * records, data, offset):
            if value >= 2 and value not in result:
                result[value] = len(result)
        offset += words * 8
    return result


def _append(parent, tag, value=None, **attributes):
    doc = parent.ownerDocument
    node = doc.createElement(tag)
    for key, content in attributes.items():
        node.setAttribute(key, content)
    if value is not None:
        node.appendChild(doc.createTextNode(str(value)))
    parent.appendChild(node)
    return node


def remove_hierarchy_storage(model, table, name):
    name, _ = resolve_hierarchy(model, table, name)
    image = model._image
    physical = model._dimension_ids[table]
    stem = f'U${physical}${name}'
    owned = _owned_file_pattern(stem)
    replacements, removals = {}, []
    for file in image.files:
        if not file.logical_path:
            continue
        filename = file.logical_path.split('\\')[-1]
        if owned.fullmatch(filename):
            removals.append(file.storage_path)
            continue
        if not filename.endswith(('.dim.xml', '.cub.xml', '.tbl.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        if filename.startswith(physical + '.'):
            for node in list(_elements(doc, 'Hierarchy')):
                if _text(_child(node, 'ID')) == name:
                    node.parentNode.removeChild(node)
                    changed = True
            for node in list(_elements(doc, 'XMObject')):
                if node.getAttribute('class') == 'XMUserHierarchy' and node.getAttribute('name') == name:
                    node.parentNode.removeChild(node)
                    changed = True
            for node in _elements(doc, 'DataFileList'):
                names = _text(node).split(';')
                retained = [n for n in names if not owned.fullmatch(n)]
                if names != retained:
                    node.firstChild.data = ';'.join(retained)
                    changed = True
        if filename.endswith('.cub.xml'):
            for dimension in _elements(doc, 'Dimension'):
                if any(getattr(n, 'localName', None) == 'DimensionID' and _text(n) == physical for n in dimension.childNodes):
                    for node in list(_elements(dimension, 'Hierarchy')):
                        if _text(_child(node, 'HierarchyID')) == name:
                            node.parentNode.removeChild(node)
                            changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements, removals=removals))


def add_hierarchy_storage(model, table, name, levels, *, _identity=None):
    existing = minidom.parseString(model._image.read_logical_file(model._dimension_paths[table]))
    if any(name.casefold() in {_text(_child(h, 'ID')).casefold(), _text(_child(h, 'Name')).casefold()}
           for h in _elements(_elements(existing, 'Hierarchies')[0], 'Hierarchy')):
        raise ValueError(f'Hierarchy {table}.{name} already exists')
    image = model._image
    physical = model._dimension_ids[table]
    identity = storage_identity(name, 'Hierarchy') if _identity is None else _identity
    stem = f'U${physical}${identity}'
    calculated = {c.name for c in model.list_calculated_columns() if c.table_id == physical}
    unprocessed = bool(calculated.intersection(levels))
    ids = []
    if unprocessed:
        offsets, arrays, count = [-1] * len(levels), (), 0
    else:
        with model._open_reader() as reader:
            frame = reader.get_table(table)
        for column in levels:
            encoded = model._encodings[(table, column)].encode(frame[column].tolist())
            ids.append(encoded[0] if isinstance(encoded, tuple) else encoded)
        offsets, arrays = traversal_indexes(ids, [intrinsic_order(model, table, c) for c in levels])
        count = len(arrays[0])
    filenames = [] if unprocessed else [f'2.{stem}.{column}.0.idf' for column in _HIERARCHY_STREAMS]
    meta_name = stem + '.2.tbl.xml'
    dim = next(f for f in image.files if f.storage_path == model._dimension_paths[table])
    doc = minidom.parseString(image.read_logical_file(dim))
    attributes = [a for a in _child(_elements(doc, 'Dimension')[0], 'Attributes').childNodes
                  if getattr(a, 'localName', None) == 'Attribute']
    identities = {_text(_child(a, 'Name')): _text(_child(a, 'ID')) for a in attributes}
    hierarchy = _append(_elements(doc, 'Hierarchies')[0], 'Hierarchy')
    for key, value in [('Name', name), ('ID', identity), ('Description', ''),
                       ('ddl300:ProcessingState', 'Unprocessed' if unprocessed else 'Processed'),
                       ('ddl300:StructureType', 'Unknown' if unprocessed else 'Natural'),
                       ('ddl300:RaggedHoleState', 'None'), ('DisplayFolder', ''), ('AllMemberName', ''),
                       ('MemberNamesUnique', 'false'), ('AllowDuplicateNames', 'true'),
                       ('ddl2:MemberKeysUnique', 'NotUnique'), ('AllMemberTranslations', None), ('Translations', None)]:
        _append(hierarchy, key, value)
    container = _append(hierarchy, 'Levels')
    for column in levels:
        level = _append(container, 'Level')
        for key, value in [('Name', column), ('ID', identities.get(column, column)), ('Description', ''),
                           ('SourceAttributeID', identities.get(column, column)), ('HideMemberIf', 'Never'), ('Translations', None)]:
            _append(level, key, value)
    if not unprocessed:
        _elements(doc, 'DataFileList')[0].firstChild.data += ';' + ';'.join(filenames + [meta_name])
    replacements = {dim.storage_path: doc.documentElement.toxml(encoding='utf-8')}
    raw_file = next(f for f in image.files if f.logical_path and
                    f.logical_path.split('\\')[-1].startswith(physical + '.') and f.logical_path.endswith('.tbl.xml'))
    doc = minidom.parseString(image.read_logical_file(raw_file))
    collection = next(n for n in _elements(doc, 'Collection') if _text(_child(n, 'Name')) == 'UserHierarchies')
    obj = _append(collection, 'XMObject', **{'class': 'XMUserHierarchy', 'name': identity, 'ProviderVersion': '0'})
    props = _append(obj, 'Properties')
    _append(props, 'IsProcessed', 'false' if unprocessed else 'true', **{'xsi:type': 'xsd:boolean'})
    _append(props, 'TableStore', ''.join(f'${identities.get(column, column)}${offset}' for column, offset in zip(levels, offsets)) + '$')
    _append(props, 'TableName', '' if unprocessed else stem)
    replacements[raw_file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    cube = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(image.read_logical_file(cube))
    dimension = next(n for n in _elements(doc, 'Dimension')
                     if any(getattr(c, 'localName', None) == 'DimensionID' and _text(c) == physical for c in n.childNodes))
    hierarchy = _append(_child(dimension, 'Hierarchies'), 'Hierarchy')
    for key, value in [('HierarchyID', identity), ('Enabled', 'true'), ('Visible', 'true'), ('OptimizedState', 'FullyOptimized')]:
        _append(hierarchy, key, value)
    replacements[cube.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if unprocessed:
        return AbfImage(image.assemble(replacements))
    template = (Path(__file__).parent / 'data/hierarchy-index.xml').read_text(encoding='utf-8')
    doc = minidom.parseString(template.replace('HIERARCHY_INDEX', escape(stem, {'"': '&quot;'})))
    for node in _elements(doc, 'Records'):
        node.firstChild.data = str(count)
    contents = [(meta_name, doc.documentElement.toxml(encoding='utf-8'))]
    for index, values in enumerate(arrays):
        encoded = [v + (index != 0) for v in values]
        # Native NoSplit<32> includes an extra zero word after the packed values.
        encoded += [0] * (2 - len(encoded) % 2)
        contents.append((filenames[index], struct.pack('<Q', len(encoded) // 2) + struct.pack('<' + 'I' * len(encoded), *encoded)))
    anchor = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.idf') and
                  f'.{physical}.' in f.logical_path.split('\\')[-1])
    parent = anchor.logical_path.rsplit('\\', 1)[0]
    trailer = image.read_file(anchor)[-4:] if image.error_code else None
    additions = [VirtualFileAddition(uuid.uuid4().hex.upper()[:20], parent + '\\' + filename,
                                    anchor.storage_path, data, trailer) for filename, data in contents]
    return AbfImage(image.assemble(replacements, additions=additions))


def rebuild_hierarchy_storage(model, table, name, levels):
    """Rebuild traversal data while preserving level captions and presentation metadata."""
    identity, caption = resolve_hierarchy(model, table, name)
    original = model._image
    image = remove_hierarchy_storage(model, table, name)
    current = type(model)(model.source, _image=image)
    # Rebuilding changes traversal data, never the identity referenced by the
    # preserved dimension/cube metadata (including legacy or native IDs).
    image = add_hierarchy_storage(current, table, caption, levels, _identity=identity)
    preserved = {f.storage_path: original.read_logical_file(f) for f in original.files
                 if f.logical_path and f.logical_path.endswith(('.dim.xml', '.cub.xml'))}
    return AbfImage(image.assemble(preserved))


def configure_hierarchy_storage(model, table, name, *, description=None,
                                hidden=None, level_names=None):
    """Change user-facing metadata without changing hierarchy identity or indexes."""
    name, _ = resolve_hierarchy(model, table, name)
    if description is not None and not isinstance(description, str):
        raise TypeError('description must be a string')
    if hidden is not None and not isinstance(hidden, bool):
        raise TypeError('hidden must be a boolean')
    image = model._image
    path = model._dimension_paths[table]
    doc = minidom.parseString(image.read_logical_file(path))
    hierarchy = next((h for h in _elements(doc, 'Hierarchy')
                      if _text(_child(h, 'ID')) == name), None)
    if hierarchy is None:
        raise KeyError(f'Hierarchy {table}.{name} does not exist')
    def set_text(node, value):
        while node.firstChild:
            node.removeChild(node.firstChild)
        node.appendChild(node.ownerDocument.createTextNode(value))
    if description is not None:
        set_text(_child(hierarchy, 'Description'), description)
    if level_names is not None:
        levels = _elements(hierarchy, 'Level')
        captions = {identity: caption for (owner, caption), identity in model._column_ids.items() if owner == table}
        known = {captions[_text(_child(level, 'SourceAttributeID'))] for level in levels}
        if not isinstance(level_names, dict) or set(level_names) - known:
            raise ValueError('level_names must map existing source columns to captions')
        names = [level_names.get(captions[_text(_child(level, 'SourceAttributeID'))],
                                 _text(_child(level, 'Name'))) for level in levels]
        if any(not isinstance(n, str) or not n.strip() for n in names):
            raise ValueError('Level captions must be nonempty strings')
        if len({n.casefold() for n in names}) != len(names):
            raise ValueError('Level captions must be distinct')
        for level, caption in zip(levels, names):
            set_text(_child(level, 'Name'), caption)
    replacements = {path: doc.documentElement.toxml(encoding='utf-8')}
    if hidden is not None:
        cube = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
        doc = minidom.parseString(image.read_logical_file(cube))
        dimension = next(n for n in _elements(doc, 'Dimension')
                         if any(getattr(c, 'localName', None) == 'DimensionID' and
                                _text(c) == model._dimension_ids[table] for c in n.childNodes))
        hierarchy = next(h for h in _elements(dimension, 'Hierarchy')
                         if _text(_child(h, 'HierarchyID')) == name)
        set_text(_child(hierarchy, 'Visible'), str(not hidden).lower())
        replacements[cube.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))
