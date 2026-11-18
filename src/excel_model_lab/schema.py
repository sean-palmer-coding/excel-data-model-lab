"""Experimental schema primitives; desktop acceptance is recorded separately."""
from __future__ import annotations

import re
import uuid
from xml.dom import minidom

from .abf import AbfImage, VirtualFileAddition, VirtualFileGroupAddition, UnsupportedModelError


def storage_identity(name, kind):
    """Keep legal native IDs; separate punctuation-bearing captions from IDs."""
    if not isinstance(name, str) or not name.strip():
        raise ValueError('Names must be nonempty strings')
    forbidden = set(".,;'`:/\\*|?\"&%$!+=[]{}<>")
    if any(c in forbidden or ord(c) < 32 for c in name):
        return '_' + kind + '_' + uuid.uuid4().hex
    return name


def _elements(node, local):
    return [n for n in node.getElementsByTagName('*') if n.localName == local]


def _text(node):
    return ''.join(c.data for c in node.childNodes if c.nodeType == c.TEXT_NODE)


def _child(node, local):
    return next(n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == local)


def _next_table_ordinal(image):
    # Native workbooks can start at one and retain gaps after table removal.
    # Reusing len(tables) can duplicate a live ordinal and break native editing.
    ordinals = []
    for file in image.files:
        if file.logical_path and file.logical_path.endswith(('.dim.xml', '.det.xml')):
            doc = minidom.parseString(image.read_logical_file(file))
            ordinals.extend(int(_text(n)) for n in _elements(doc, 'Ordinal'))
    return max(ordinals, default=-1) + 1


def configure_table(image: AbfImage, table_id: str, *, hidden=None, description=None) -> AbfImage:
    if hidden is not None and not isinstance(hidden, bool):
        raise TypeError('hidden must be a boolean')
    if description is not None and not isinstance(description, str):
        raise TypeError('description must be a string')
    replacements = {}
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.dim.xml', '.det.xml', '.cub.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        changed = False
        for node in _elements(doc, 'Dimension') + _elements(doc, 'MeasureGroup'):
            ids = [n for n in node.childNodes if getattr(n, 'localName', None) == 'ID']
            if not ids or _text(ids[0]) != table_id:
                continue
            property_name, value = ('Visible', 'false' if hidden else 'true') if file.logical_path.endswith('.cub.xml') else ('Description', description)
            if (file.logical_path.endswith('.cub.xml') and hidden is None) or value is None:
                continue
            target = _child(node, property_name)
            while target.firstChild:
                target.removeChild(target.firstChild)
            target.appendChild(doc.createTextNode(value))
            changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))


def change_numeric_type(image: AbfImage, table_id: str, column: str,
                        data_type: str, magnitude: int) -> AbfImage:
    """Change a value-encoded column's semantic and physical numeric type."""
    # Public APIs take captions; native models can retain a different stable ID.
    for file in image.files:
        if not file.logical_path:
            continue
        filename = file.logical_path.replace('/', '\\').rsplit('\\', 1)[-1]
        if not filename.startswith(table_id + '.') or not filename.endswith('.dim.xml'):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        for attribute in _elements(doc, 'Attribute'):
            names = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'Name']
            if names and _text(names[0]) == column:
                column = _text(_child(attribute, 'ID'))
                break
        break
    types = {'integer': ('BigInt', '20', '0', 'XM_Long'),
             'decimal': ('Double', '5', '1', 'XM_Real'),
             'datetime': ('Date', '7', '1', 'XM_Real'),
             'currency': ('Currency', '6', '0', 'XM_Long'),
             'boolean': ('Boolean', '11', '0', 'XM_Long')}
    semantic, dbtype, xmtype, dictionary = types[data_type]
    replacements = {}
    found = set()
    for f in image.files:
        if not f.logical_path:
            continue
        filename = f.logical_path.replace('/', '\\').rsplit('\\', 1)[-1]
        if filename.endswith('.det.xml'):
            doc = minidom.parseString(image.read_logical_file(f))
            changed = False
            for dimension in _elements(doc, 'Dimension'):
                if _text(_child(dimension, 'CubeDimensionID')) != table_id:
                    continue
                for attribute in _elements(dimension, 'Attribute'):
                    if _text(_child(attribute, 'AttributeID')) == column:
                        for key in _elements(attribute, 'KeyColumn'):
                            _child(key, 'DataType').firstChild.data = semantic
                            changed = True
            if changed:
                replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
            continue
        if not filename.startswith(table_id + '.') or not filename.endswith(('.dim.xml', '.tbl.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(f))
        if filename.endswith('.dim.xml'):
            for attribute in _elements(doc, 'Attribute'):
                ids = [n for n in attribute.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == 'ID']
                if not ids or _text(ids[0]) != column:
                    continue
                keys = _elements(attribute, 'KeyColumn')
                if len(keys) != 1:
                    raise UnsupportedModelError('Type changes require a single column key')
                _child(keys[0], 'DataType').firstChild.data = semantic
                found.add('dimension')
        else:
            for raw in _elements(doc, 'XMObject'):
                if raw.getAttribute('class') != 'XMRawColumn' or raw.getAttribute('name') != column:
                    continue
                dictionaries = [n for n in _elements(raw, 'XMObject')
                                if n.getAttribute('class').startswith('XMValueDataDictionary<')]
                if len(dictionaries) != 1:
                    raise UnsupportedModelError('Type changes require value encoding')
                dictionaries[0].setAttribute('class', f'XMValueDataDictionary<{dictionary}>')
                for tag, value in [('DBType', dbtype), ('XMType', xmtype), ('Magnitude', str(magnitude))]:
                    for node in _elements(raw, tag):
                        node.firstChild.data = value
                found.add('storage')
        replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if found != {'dimension', 'storage'}:
        raise UnsupportedModelError('Cannot resolve numeric column metadata')
    return AbfImage(image.assemble(replacements))


def configure_column(image: AbfImage, table_id: str, column: str, *,
                     hidden: bool | None = None, description: str | None = None,
                     format_string: str | None = None) -> AbfImage:
    """Edit column presentation metadata without changing stored values."""
    if hidden is not None and not isinstance(hidden, bool):
        raise TypeError('hidden must be a boolean')
    if any(v is not None and not isinstance(v, str) for v in (description, format_string)):
        raise TypeError('Description and format string must be strings')
    replacements = {}
    found = False
    for f in image.files:
        if not f.logical_path or not f.logical_path.endswith(('.dim.xml', '.cub.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(f))
        changed = False
        for dimension in _elements(doc, 'Dimension'):
            ids = [_text(n) for n in dimension.childNodes if n.nodeType == n.ELEMENT_NODE
                   and n.localName in ('ID', 'DimensionID')]
            if table_id not in ids:
                continue
            for attribute in _elements(dimension, 'Attribute'):
                ids = [_text(n) for n in attribute.childNodes if n.nodeType == n.ELEMENT_NODE
                       and n.localName in ('ID', 'AttributeID')]
                if column not in ids:
                    continue
                def assign(local, value):
                    nodes = [n for n in attribute.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == local]
                    if nodes:
                        node = nodes[0]
                    elif local == 'FormatString':
                        node = doc.createElementNS('http://schemas.microsoft.com/analysisservices/2011/engine/300/300',
                                                   'ddl300_300:FormatString')
                        attribute.appendChild(node)
                    else:
                        raise UnsupportedModelError(f'Missing column property: {local}')
                    while node.firstChild:
                        node.removeChild(node.firstChild)
                    node.appendChild(doc.createTextNode(value))
                if hidden is not None:
                    assign('AttributeHierarchyVisible', 'false' if hidden else 'true')
                    changed = True
                if f.logical_path.endswith('.dim.xml'):
                    found = True
                    if description is not None:
                        assign('Description', description)
                        changed = True
                    if format_string is not None:
                        assign('FormatString', format_string)
                        changed = True
                        # Excel's UI format annotation may override the custom
                        # string. Remove stale annotations when replacing it.
                        for annotation in list(_elements(attribute, 'Annotation')):
                            if _text(_child(annotation, 'Name')) == 'Format':
                                annotation.parentNode.removeChild(annotation)
                        decimal = re.fullmatch(r'0\.(0+)', format_string)
                        if decimal:
                            annotations = _child(attribute, 'Annotations')
                            annotation = doc.createElementNS(attribute.namespaceURI, 'Annotation')
                            label = doc.createElementNS(attribute.namespaceURI, 'Name')
                            label.appendChild(doc.createTextNode('Format'))
                            value = doc.createElementNS(attribute.namespaceURI, 'Value')
                            display = doc.createElement('Format')
                            display.setAttribute('xmlns', '')
                            display.setAttribute('Format', 'NumberDecimal')
                            display.setAttribute('Accuracy', str(len(decimal.group(1))))
                            value.appendChild(display)
                            annotation.appendChild(label)
                            annotation.appendChild(value)
                            annotations.appendChild(annotation)
        if changed:
            replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if not found:
        raise KeyError(column)
    return AbfImage(image.assemble(replacements))


def rename_table(image: AbfImage, table_id: str, name: str) -> AbfImage:
    """Change a table's caption while retaining its storage identity."""
    if not isinstance(name, str) or not name.strip() or any(ord(c) < 32 for c in name):
        raise ValueError('Table name must be nonempty and contain no control characters')
    replacements = {}
    for f in image.files:
        if not f.logical_path or not f.logical_path.endswith(('.dim.xml', '.cub.xml', '.det.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(f))
        changed = False
        for node in _elements(doc, 'Dimension') + _elements(doc, 'MeasureGroup'):
            identifiers = [_text(n) for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
                           and n.localName in ('ID', 'DimensionID')]
            if table_id not in identifiers:
                continue
            names = [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == 'Name']
            for value in names:
                while value.firstChild:
                    value.removeChild(value.firstChild)
                value.appendChild(doc.createTextNode(name))
                changed = True
        if changed:
            replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    if not replacements:
        raise KeyError(table_id)
    return AbfImage(image.assemble(replacements))


def remove_table(image: AbfImage, table_id: str) -> AbfImage:
    """Remove a table and its storage groups without cascading dependencies."""
    groups = image.backup_log_root.findall('FileGroups/FileGroup')
    selected = [g for g in groups if g.findtext('ID') == table_id
                and g.findtext('Class') in ('100006', '100016', '100021')]
    if len(selected) != 3:
        raise UnsupportedModelError('Table removal requires one dimension, measure group and partition')
    if sum(g.findtext('Class') == '100006' for g in groups) <= 1:
        raise ValueError('The model must retain at least one table')
    removed = {n.findtext('StoragePath') for g in selected for n in g.findall('FileList/BackupFile')}
    det_names = {n.findtext('Path').split('\\')[-1] for g in selected for n in g.findall('FileList/BackupFile')
                 if n.findtext('Path').endswith('.det.xml')}
    replacements = {}
    for f in image.files:
        if not f.logical_path or not f.logical_path.endswith('.xml'):
            continue
        doc = minidom.parseString(image.read_logical_file(f))
        for end in _elements(doc, 'FromRelationshipEnd') + _elements(doc, 'ToRelationshipEnd'):
            if _text(_child(end, 'DimensionID')) == table_id:
                raise UnsupportedModelError('Table is used by a relationship')
        if f.storage_path in removed:
            continue
        if f.logical_path.endswith('.cub.xml'):
            for dimension in list(_elements(doc, 'Dimension')):
                if _text(_child(dimension, 'DimensionID')) == table_id:
                    dimension.parentNode.removeChild(dimension)
            files = _elements(doc, 'MeasureGroupFileList')[0]
            files.firstChild.data = ';'.join(n for n in _text(files).split(';') if n not in det_names)
            replacements[f.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements, removals=removed))


def clone_table(image: AbfImage, source: str, name: str) -> AbfImage:
    """Clone an independent table's physical storage and schema as a new table."""
    if any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_-]*', value) for value in (source, name)):
        raise UnsupportedModelError('Table cloning currently requires simple identifiers')
    all_groups = image.backup_log_root.findall('FileGroups/FileGroup')
    if any(g.findtext('Class') == '100006' and g.findtext('ID', '').casefold() == name.casefold()
           for g in all_groups):
        raise ValueError(f'Table already exists: {name}')
    groups = [g for g in all_groups if g.findtext('ID') == source
              and g.findtext('Class') in ('100006', '100016', '100021')]
    if len(groups) != 3:
        raise UnsupportedModelError('Table cloning requires one dimension, measure group and partition')
    # Table IDs occur in path tokens and scoped engine references. Reject
    # relationships and calculations until their cross-table bindings can be
    # cloned deliberately.
    dimension_group = next(g for g in groups if g.findtext('Class') == '100006')
    dim_file = next(n for n in dimension_group.findall('FileList/BackupFile')
                    if n.findtext('Path').endswith('.dim.xml'))
    dim_doc = minidom.parseString(image.read_logical_file(dim_file.findtext('StoragePath')))
    if _elements(dim_doc, 'Relationship') or _elements(dim_doc, 'Expression'):
        raise UnsupportedModelError('Cannot clone a table with relationships or calculations')
    if any(_text(n) == source for a in _elements(dim_doc, 'Attribute') for n in a.childNodes
           if n.nodeType == n.ELEMENT_NODE and n.localName in ('ID', 'Name')):
        raise UnsupportedModelError('Table and column identifiers must differ for table cloning')
    instance_ids = {g.findtext('ObjectID'): str(uuid.uuid4()).upper() for g in groups}
    def renamed(value):
        # Spaces are boundaries of captions such as 'Sales Amount', never of
        # the path tokens and engine references that carry the table ID.
        value = re.sub(r'(?<![A-Za-z0-9_ -])' + re.escape(source) + r'(?![A-Za-z0-9_ -])', name, value)
        for old, new in instance_ids.items():
            value = value.replace(old, new)
        return value
    additions, new_groups = [], []
    table_ordinal = _next_table_ordinal(image)
    for group in groups:
        entries = group.findall('FileList/BackupFile')
        anchor = entries[0].findtext('StoragePath')
        group_key = group.findtext('Class') + ':' + name
        new_groups.append(VirtualFileGroupAddition(
            group_key, anchor, name, name, renamed(group.findtext('PersistLocationPath')),
            instance_ids[group.findtext('ObjectID')]))
        for entry in entries:
            storage = entry.findtext('StoragePath')
            path = entry.findtext('Path')
            data = image.read_logical_file(storage)
            if path.endswith('.xml'):
                data = renamed(data.decode('utf-8')).encode('utf-8')
                if path.endswith(('.dim.xml', '.det.xml')):
                    doc = minidom.parseString(data)
                    ordinal = _elements(doc, 'Ordinal')[0]
                    ordinal.firstChild.data = str(table_ordinal)
                    data = doc.documentElement.toxml(encoding='utf-8')
            additions.append(VirtualFileAddition(
                uuid.uuid4().hex.upper()[:20], renamed(path), storage, data,
                image.read_file(storage)[-4:] if image.error_code else None, group_key))
    cubes = [f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml')]
    if len(cubes) != 1:
        raise UnsupportedModelError('Expected one cube')
    cube = cubes[0]
    doc = minidom.parseString(image.read_logical_file(cube))
    dimensions = [d for d in _elements(doc, 'Dimension')
                  if _text(_child(d, 'DimensionID')) == source]
    if len(dimensions) != 1:
        raise UnsupportedModelError('Cannot resolve cube table reference')
    cloned = minidom.parseString(renamed(dimensions[0].toxml())).documentElement
    dimensions[0].parentNode.appendChild(doc.importNode(cloned, True))
    file_list = _elements(doc, 'MeasureGroupFileList')[0]
    det_name = next(entry.findtext('Path').split('\\')[-1] for g in groups
                    for entry in g.findall('FileList/BackupFile') if entry.findtext('Path').endswith('.det.xml'))
    file_list.firstChild.data += ';' + renamed(det_name)
    return AbfImage(image.assemble({cube.storage_path: doc.documentElement.toxml(encoding='utf-8')},
                                   additions=additions, groups=new_groups))


def import_table(image: AbfImage, donor: AbfImage, table_id: str) -> AbfImage:
    """Import an independent table from the bundled authoring backend."""
    target_groups = image.backup_log_root.findall('FileGroups/FileGroup')
    donor_groups = donor.backup_log_root.findall('FileGroups/FileGroup')
    groups = [g for g in donor_groups if g.findtext('ID') == table_id
              and g.findtext('Class') in ('100006', '100016', '100021')]
    if len(groups) != 3 or any(g.findtext('ID') == table_id for g in target_groups):
        raise UnsupportedModelError('Imported table must have independent, unique storage')
    source_root = next(g.findtext('PersistLocationPath') for g in donor_groups if g.findtext('Class') == '100002')
    target_root = next(g.findtext('PersistLocationPath') for g in target_groups if g.findtext('Class') == '100002')
    target_database = next(g.findtext('ID') for g in target_groups if g.findtext('Class') == '100002')
    def relocated(path):
        if not path.startswith(source_root + '\\'):
            raise UnsupportedModelError('Imported table path is outside its database')
        return target_root + path[len(source_root):]
    additions, new_groups = [], []
    table_ordinal = _next_table_ordinal(image)
    for group in groups:
        prototype = next(g for g in target_groups if g.findtext('Class') == group.findtext('Class'))
        anchor = prototype.findtext('FileList/BackupFile/StoragePath')
        group_key = group.findtext('Class') + ':' + table_id
        new_groups.append(VirtualFileGroupAddition(group_key, anchor, table_id, table_id,
            relocated(group.findtext('PersistLocationPath')), group.findtext('ObjectID'),
            object_version=group.findtext('ObjectVersion')))
        for entry in group.findall('FileList/BackupFile'):
            storage, path = entry.findtext('StoragePath'), entry.findtext('Path')
            data = donor.read_logical_file(storage)
            if path.endswith(('.dim.xml', '.det.xml', '.prt.xml')):
                doc = minidom.parseString(data)
                if path.endswith(('.dim.xml', '.det.xml')):
                    _elements(doc, 'Ordinal')[0].firstChild.data = str(table_ordinal)
                for node in _elements(doc, 'DatabaseID'):
                    node.firstChild.data = target_database
                data = doc.documentElement.toxml(encoding='utf-8')
            trailer = donor.read_file(storage)[-4:] if donor.error_code else None
            if image.error_code and trailer is None:
                raise UnsupportedModelError('Imported storage requires native file trailers')
            additions.append(VirtualFileAddition(uuid.uuid4().hex.upper()[:20], relocated(path),
                anchor, data, trailer if image.error_code else None, group_key))
    cube = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    source_cube = next(f for f in donor.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(image.read_logical_file(cube))
    source_doc = minidom.parseString(donor.read_logical_file(source_cube))
    dimension = next(d for d in _elements(source_doc, 'Dimension') if _text(_child(d, 'DimensionID')) == table_id)
    _elements(doc, 'Dimensions')[0].appendChild(doc.importNode(dimension, True))
    det_name = next(entry.findtext('Path').split('\\')[-1] for g in groups
        for entry in g.findall('FileList/BackupFile') if entry.findtext('Path').endswith('.det.xml'))
    _elements(doc, 'MeasureGroupFileList')[0].firstChild.data += ';' + det_name
    return AbfImage(image.assemble({cube.storage_path: doc.documentElement.toxml(encoding='utf-8')},
        additions=additions, groups=new_groups))


def remove_column(image: AbfImage, table_id: str, name: str) -> AbfImage:
    """Remove an independent stored column, rejecting model dependencies."""
    from .formula_statements import formula_references, iter_measure_definitions
    files = {f.storage_path: f for f in image.files if f.logical_path}
    docs = {key: minidom.parseString(image.read_logical_file(f))
            for key, f in files.items() if f.logical_path.endswith('.xml')}
    matches = [(key, d) for key, doc in docs.items() for d in _elements(doc, 'Dimension')
               if d.parentNode.localName == 'ObjectDefinition'
               and _text(_child(d, 'ID')) == table_id]
    if len(matches) != 1:
        raise UnsupportedModelError('Cannot resolve table dimension')
    dim_key, dimension = matches[0]
    table_name = _text(_child(dimension, 'Name'))
    table_names = [_text(_child(d, 'Name')) for doc in docs.values()
                   for d in _elements(doc, 'Dimension')
                   if d.parentNode.localName == 'ObjectDefinition']
    attrs = _child(dimension, 'Attributes')
    attributes = [n for n in attrs.childNodes if n.nodeType == n.ELEMENT_NODE]
    target = next((n for n in attributes if _text(_child(n, 'ID')) == name), None)
    if target is None:
        raise KeyError(name)
    if _text(_child(target, 'Usage')) != 'Regular':
        raise UnsupportedModelError('Cannot remove an engine column')
    # DAX refers to the caption; storage and bindings use the attribute ID.
    caption = _text(_child(target, 'Name')) or name
    if sum(_text(_child(n, 'Usage')) == 'Regular' for n in attributes) <= 1:
        raise ValueError('A table must retain at least one stored column')
    for doc in docs.values():
        for end in _elements(doc, 'FromRelationshipEnd') + _elements(doc, 'ToRelationshipEnd'):
            if _text(_child(end, 'DimensionID')) == table_id and any(
                _text(n) == name for n in _elements(end, 'AttributeID')
            ):
                raise UnsupportedModelError('Column is used by a relationship')
        for expression in _elements(doc, 'Expression') + _elements(doc, 'Text'):
            text = _text(expression)
            definitions = list(iter_measure_definitions(text)) if expression.localName == 'Text' else []
            expressions = [(definition.group('table').replace("''", "'"), definition.group('expression'))
                           for definition in definitions] if definitions else [(None, text)]
            for home, formula in expressions:
                for ref in formula_references(formula, table_names=table_names):
                    if ref.kind == 'field' and ref.name.casefold() == caption.casefold() and (
                            ref.table is not None and ref.table.casefold() == table_name.casefold()
                            or ref.table is None and (home is None or home.casefold() == table_name.casefold())):
                        raise UnsupportedModelError('Column may be used by a calculation')
    for reference in _elements(dimension, 'OrderByAttributeID') + _elements(dimension, 'SourceAttributeID'):
        if _text(reference) == name:
            raise UnsupportedModelError('Column is used by sorting or a hierarchy')
    owns = column_file_matcher(table_id, name, [_text(_child(n, 'ID')) for n in attributes])
    removed = {key for key, f in files.items() if owns(f.logical_path.split('\\')[-1])}
    if not removed:
        raise UnsupportedModelError('Cannot resolve column data files')
    filenames = {files[key].logical_path.split('\\')[-1] for key in removed}
    data_list = _child(dimension, 'DataFileList')
    data_list.firstChild.data = ';'.join(n for n in _text(data_list).split(';') if n not in filenames)
    changed = {dim_key}
    ordinal = attributes.index(target)
    attrs.removeChild(target)
    for relation in list(_elements(dimension, 'AttributeRelationship')):
        if _text(_child(relation, 'AttributeID')) == name:
            relation.parentNode.removeChild(relation)
    dim_directory = files[dim_key].logical_path.rsplit('\\', 1)[0] + '\\' + table_id + '.0.dim\\'
    for key, doc in docs.items():
        if key in removed:
            continue
        if files[key].logical_path.startswith(dim_directory) and files[key].logical_path.split('\\')[-1].startswith('info.'):
            properties = _elements(doc, 'Property')
            if len(properties) != len(attributes):
                raise UnsupportedModelError('Unexpected dimension property count')
            properties[ordinal].parentNode.removeChild(properties[ordinal])
            changed.add(key)
        for d in _elements(doc, 'Dimension'):
            identifiers = [_text(n) for n in d.childNodes if n.nodeType == n.ELEMENT_NODE
                           and n.localName in ('DimensionID', 'CubeDimensionID')]
            if table_id in identifiers:
                for a in list(_elements(d, 'Attribute')):
                    if _text(_child(a, 'AttributeID')) == name:
                        a.parentNode.removeChild(a)
                        changed.add(key)
        root = doc.documentElement
        if root.getAttribute('class') == 'XMSimpleTable' and root.getAttribute('name') == table_id:
            for raw in list(_elements(root, 'XMObject')):
                if raw.getAttribute('class') == 'XMRawColumn' and raw.getAttribute('name') == name:
                    raw.parentNode.removeChild(raw)
                    changed.add(key)
    return AbfImage(image.assemble({key: docs[key].documentElement.toxml(encoding='utf-8')
                                   for key in changed}, removals=removed))


def _column_file_pattern(table_id, column_id):
    t, c = re.escape(table_id), re.escape(column_id)
    return re.compile(rf'\d+\.{t}\.{c}\.(?:\d+\.idf|dictionary)'
                      rf'|\d+\.H\${t}\${c}\.(?:(?:ID_TO_POS|POS_TO_ID)\.\d+\.idf|hidx)'
                      rf'|H\${t}\${c}\.\d+\.tbl\.xml')


def _column_file_loose(filename, table_id, column_id):
    return f'.{table_id}.{column_id}.' in filename or f'H${table_id}${column_id}.' in filename


def column_file_matcher(table_id, column_id, sibling_ids=()):
    """Return a predicate selecting only ``column_id``'s physical streams.

    Column IDs may contain '.', so ``No`` must not claim ``No.``'s files.
    Known native names are matched exactly; an unrecognized name is claimed
    only when no sibling column explains it equally or more specifically.
    """
    own = _column_file_pattern(table_id, column_id)
    siblings = [s for s in dict.fromkeys(sibling_ids) if s != column_id]
    sibling_patterns = [_column_file_pattern(table_id, s) for s in siblings]

    def owns(filename):
        if any(p.fullmatch(filename) for p in sibling_patterns):
            return False
        if own.fullmatch(filename):
            return True
        return _column_file_loose(filename, table_id, column_id) and not any(
            len(s) > len(column_id) and _column_file_loose(filename, table_id, s) for s in siblings)
    return owns


def clone_column(image: AbfImage, table_id: str, source: str, name: str) -> AbfImage:
    """Clone a stored column and its lookup structures within the same table.

    Values and type are copied; this primitive does not synthesize an encoding.
    Names are preserved in archive paths; path separators and the intrinsic
    hierarchy delimiter are rejected until their escaping dialect is supported.
    """
    for value in (table_id, source, name):
        if (not isinstance(value, str) or not value.strip()
                or any(ord(c) < 32 or c in '\\/$' for c in value)):
            raise UnsupportedModelError('Schema names must be nonempty and cannot contain controls, path separators or $')
    if source.casefold() == name.casefold():
        raise ValueError('New column must have a distinct name')
    files = {f.logical_path.replace('\\', '/').split('/')[-1]: f
             for f in image.files if f.logical_path}
    def unique(pattern):
        found = [f for key, f in files.items() if re.fullmatch(pattern, key)]
        if len(found) != 1:
            raise UnsupportedModelError(f'Cannot uniquely resolve schema file: {pattern}')
        return found[0]
    dim = unique(re.escape(table_id) + r'\.\d+\.dim\.xml')
    tbl = unique(re.escape(table_id) + r'\.\d+\.tbl\.xml')
    det = unique(re.escape(table_id) + r'\.\d+\.det\.xml')
    cube = unique(r'Model\.\d+\.cub\.xml')
    docs = {f.storage_path: minidom.parseString(image.read_logical_file(f)) for f in (dim,tbl,det,cube)}
    dimension = _elements(docs[dim.storage_path], 'Dimension')[0]
    attrs = _child(dimension, 'Attributes')
    attributes = [n for n in attrs.childNodes if n.nodeType == n.ELEMENT_NODE]
    if any(name.casefold() in {_text(_child(n,'ID')).casefold(),
                              _text(_child(n,'Name')).casefold()} for n in attributes):
        raise ValueError(f'Column already exists: {name}')
    caption = name
    name = storage_identity(name, 'Column')
    original = next((n for n in attributes if _text(_child(n,'ID')) == source), None)
    if original is None:
        raise KeyError(source)
    if _text(_child(original, 'Usage')) != 'Regular':
        raise UnsupportedModelError('Cannot clone an engine column')
    copied = {}
    owns = column_file_matcher(table_id, source, [_text(_child(n, 'ID')) for n in attributes])
    for filename, f in files.items():
        if owns(filename):
            copied[filename] = filename.replace(f'.{table_id}.{source}.',f'.{table_id}.{name}.').replace(
                f'H${table_id}${source}.', f'H${table_id}${name}.')
    if not copied:
        raise UnsupportedModelError('No physical column files found')
    def rename(node):
        if node.nodeType == node.ELEMENT_NODE:
            for i in range(node.attributes.length):
                attr = node.attributes.item(i)
                # Only the raw column object's name is a standalone physical
                # identity. Class names, types, versions and nested object
                # names such as Dictionary are schema constants, even when a
                # user happens to give their column the same name.
                if (attr.localName == 'name' and node.localName == 'XMObject'
                        and node.getAttribute('class') == 'XMRawColumn'
                        and attr.value == source):
                    attr.value = name
                else:
                    attr.value = attr.value.replace(f'[Hierarchy for column {source}]',
                                                    f'[Hierarchy for column {name}]')
                    for old,new in copied.items(): attr.value = attr.value.replace(old,new)
                    attr.value = attr.value.replace(f'H${table_id}${source}',f'H${table_id}${name}')
        elif node.nodeType == node.TEXT_NODE:
            parent = node.parentNode
            identity = (parent.localName in ('AttributeID', 'SourceAttributeID', 'OrderByAttributeID')
                        or (parent.localName == 'ID' and parent.parentNode.localName == 'Attribute'))
            # Captions are assigned explicitly below. Do not rewrite enum or
            # scalar values: cloning a column named Key must retain OrderBy=Key,
            # and a column named Regular must retain Type/Usage=Regular.
            if node.data == source and identity: node.data = name
            else:
                for old,new in copied.items(): node.data = node.data.replace(old,new)
                node.data = node.data.replace(f'H${table_id}${source}',f'H${table_id}${name}')
        for child in node.childNodes: rename(child)
        return node
    cloned = rename(original.cloneNode(True))
    # ``source`` is the physical ID; a natively renamed column's caption
    # differs from it and would otherwise be copied verbatim.
    caption_node = next((n for n in cloned.childNodes
                         if n.nodeType == n.ELEMENT_NODE and n.localName == 'Name'), None)
    if caption_node is not None:
        while caption_node.firstChild:
            caption_node.removeChild(caption_node.firstChild)
        caption_node.appendChild(caption_node.ownerDocument.createTextNode(caption))
    # ShortColumnId is a unique UI annotation, not the physical column name.
    used_short_ids = {_text(_child(a,'Value')) for a in _elements(dimension,'Annotation')
                      if _text(_child(a,'Name')) == 'ShortColumnId'}
    index=1
    while True:
        number=index
        short_id=''
        while number:
            number,remainder=divmod(number-1,26)
            short_id=chr(65+remainder)+short_id
        if short_id not in used_short_ids: break
        index+=1
    for annotation in _elements(cloned,'Annotation'):
        if _text(_child(annotation,'Name')) == 'ShortColumnId':
            val=_child(annotation,'Value')
            while val.firstChild: val.removeChild(val.firstChild)
            val.appendChild(val.ownerDocument.createTextNode(short_id))
    attrs.appendChild(cloned)
    for attribute in attributes:
        for relation in list(_elements(attribute,'AttributeRelationship')):
            if _text(_child(relation,'AttributeID')) == source:
                cloned_relation = rename(relation.cloneNode(True))
                # Relationship display names can differ from AttributeID (Excel
                # emits ID_ for a column named ID). They must remain unique.
                relation_name = _child(cloned_relation, 'Name')
                while relation_name.firstChild:
                    relation_name.removeChild(relation_name.firstChild)
                # These are member-property names, not column captions. An
                # ordinary column such as ID is reserved in this namespace.
                # Use an internal suffix and avoid existing sibling names.
                property_name = name + '_'
                used_names = {_text(_child(n, 'Name')).casefold()
                              for n in _elements(relation.parentNode, 'AttributeRelationship')}
                while property_name.casefold() in used_names:
                    property_name += '_'
                relation_name.appendChild(relation_name.ownerDocument.createTextNode(property_name))
                relation.parentNode.appendChild(cloned_relation)
    for f in (cube,det):
        doc=docs[f.storage_path]
        # Only the table's own dimension, not unrelated dimensions in a cube.
        for d in _elements(doc,'Dimension'):
            ids=[_text(n) for n in d.childNodes if n.nodeType == n.ELEMENT_NODE
                 and n.localName in ('DimensionID','CubeDimensionID')]
            if table_id not in ids: continue
            for attr in list(_elements(d,'Attribute')):
                if _text(_child(attr,'AttributeID')) == source:
                    attr.parentNode.appendChild(rename(attr.cloneNode(True)))
    raw = next(n for n in _elements(docs[tbl.storage_path],'XMObject')
               if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == source)
    raw.parentNode.appendChild(rename(raw.cloneNode(True)))
    dfl=_child(dimension,'DataFileList')
    dfl.firstChild.data += ';' + ';'.join(copied.values())
    # The extension metadata has one property for each dimension attribute.
    dim_dir=dim.logical_path.rsplit('\\',1)[0]+'\\'+table_id+'.0.dim\\'
    info=[f for f in image.files if f.logical_path and f.logical_path.startswith(dim_dir)
          and re.fullmatch(r'info\.\d+\.xml',f.logical_path.split('\\')[-1])]
    if len(info)!=1: raise UnsupportedModelError('Cannot resolve dimension property metadata')
    info=info[0]
    info_doc=minidom.parseString(image.read_logical_file(info))
    properties=_elements(info_doc,'Property')
    if len(properties)!=len(attributes): raise UnsupportedModelError('Unexpected dimension properties')
    properties[0].parentNode.appendChild(properties[attributes.index(original)].cloneNode(True))
    docs[info.storage_path]=info_doc
    owners={n.findtext('StoragePath'):n.findtext('Path') for n in image.backup_log_root.findall(
        'FileGroups/FileGroup/FileList/BackupFile')}
    additions=[]
    for old,new in copied.items():
        f=files[old]
        data=image.read_logical_file(f)
        if old.endswith('.xml'):
            doc=minidom.parseString(data);rename(doc);data=doc.documentElement.toxml(encoding='utf-8')
        path=owners[f.storage_path].rsplit('\\',1)[0]+'\\'+new
        additions.append(VirtualFileAddition(uuid.uuid4().hex.upper()[:20],path,f.storage_path,data,
                                              image.read_file(f)[-4:] if image.error_code else None))
    return AbfImage(image.assemble({key:doc.documentElement.toxml(encoding='utf-8') for key,doc in docs.items()},
                                   additions=additions))
