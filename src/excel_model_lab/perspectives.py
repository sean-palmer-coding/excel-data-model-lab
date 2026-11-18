"""Named subsets of Excel model objects (not worksheet PivotTable filters)."""
from datetime import datetime, timezone
from copy import deepcopy
import re
import uuid
from xml.etree import ElementTree as ET
from xml.dom import minidom

from .abf import AbfImage, VirtualFileAddition, VirtualFileGroupAddition
from .measures import ENGINE, list_script_measures
from .schema import _child, _elements, _text
from .errors import ModelValidationError


class PerspectiveValidationError(ModelValidationError, ValueError):
    """Invalid perspective selection or name."""


class PerspectiveLookupError(ModelValidationError, KeyError):
    """Missing perspective/member, preserving KeyError compatibility."""


class PerspectiveTypeError(ModelValidationError, TypeError):
    """Incorrect argument type, preserving TypeError compatibility."""


def _error(name, message, *, kind='value', code='invalid_perspective'):
    from .validation import ValidationIssue
    error = {'value': PerspectiveValidationError, 'lookup': PerspectiveLookupError,
             'type': PerspectiveTypeError}[kind]
    raise error(message, issues=(ValidationIssue(code, message,
        object_path=('perspectives', str(name))),))


def _tag(name):
    return '{' + ENGINE + '}' + name


def _xml(image):
    for file in image.files:
        if file.logical_path and file.logical_path.endswith('.persp.xml'):
            yield file, ET.fromstring(image.read_logical_file(file))


def list_perspectives(image):
    return tuple(root.findtext('.//' + _tag('Perspective') + '/' + _tag('Name'))
                 for _, root in _xml(image))


def _validate_name(name):
    if not isinstance(name, str) or not name or name != name.strip() or name.casefold() == 'model' or re.search(r'[;\\/:*?"<>|\x00-\x1f]', name):
        _error(name, 'Perspective name must be nonempty, trimmed and contain no filename-reserved characters')


def _preserve_retained_members(old, new, member_tag, identity_tag):
    """Keep metadata for retained members, without restoring removed members."""
    if old is None:
        return
    new.attrib.update(old.attrib)
    member_tag, identity_tag = _tag(member_tag), _tag(identity_tag)
    previous = {item.findtext(identity_tag): item for item in old if item.tag == member_tag}
    for member in list(new):
        if member.tag != member_tag:
            continue
        original = previous.get(member.findtext(identity_tag))
        if original is not None:
            index = list(new).index(member)
            new.remove(member)
            new.insert(index, deepcopy(original))
    for child in old:
        if child.tag != member_tag:
            new.append(deepcopy(child))


def _preserve_dimension_members(old, new):
    if old is None:
        return
    new.attrib.update(old.attrib)
    previous = {item.findtext(_tag('CubeDimensionID')): item
                for item in old if item.tag == _tag('Dimension')}
    for dimension in new.findall(_tag('Dimension')):
        original = previous.get(dimension.findtext(_tag('CubeDimensionID')))
        if original is None:
            continue
        dimension.attrib.update(original.attrib)
        _preserve_retained_members(original.find(_tag('Attributes')),
                                  dimension.find(_tag('Attributes')), 'Attribute', 'AttributeID')
        edited = {_tag('CubeDimensionID'), _tag('Attributes'), _tag('Hierarchies')}
        for child in original:
            if child.tag not in edited:
                dimension.append(deepcopy(child))
    for child in old:
        if child.tag != _tag('Dimension'):
            new.append(deepcopy(child))


def set_perspective(image, name, *, columns=(), measures=(), hierarchies=(), description='', replace=False):
    """Create/replace a subset. Members are pairs of display table/object names.

    Updates preserve the perspective identity and filename. The original model
    is returned untouched on validation failure. Perspective membership never
    changes an object's visibility in the default model.
    """
    _validate_name(name)
    if not isinstance(description, str):
        _error(name, 'description must be a string', kind='type')
    existing = [(f, r) for f, r in _xml(image)
                if r.findtext('.//' + _tag('Perspective') + '/' + _tag('Name')).casefold() == name.casefold()]
    if bool(existing) != bool(replace):
        _error(name, 'Perspective already exists' if existing else 'Perspective does not exist')
    tables = {}
    for f in image.files:
        if f.logical_path and f.logical_path.endswith('.dim.xml'):
            root = ET.fromstring(image.read_logical_file(f))
            d = root.find('.//' + _tag('Dimension'))
            if d is None:
                continue
            tables[d.findtext(_tag('Name'))] = (d.findtext(_tag('ID')), d)
    available_measures = {}
    for f in image.files:
        if f.logical_path and f.logical_path.endswith('.scr.xml'):
            available_measures.update({(m.table, m.name): m for m in list_script_measures(image.read_logical_file(f))})
    members = {}
    for kind, values in (('columns', columns), ('measures', measures), ('hierarchies', hierarchies)):
        resolved = []
        try:
            values = iter(values)
        except TypeError:
            _error(name, f'{kind} must be an iterable of (table, name) pairs', kind='type')
        for pair in values:
            if not isinstance(pair, (tuple, list)) or len(pair) != 2:
                _error(name, f'{kind} members must be (table, name) pairs', kind='type')
            table, member = pair
            if not isinstance(table, str) or not isinstance(member, str):
                _error(name, f'{kind} table and member names must be strings', kind='type')
            if table not in tables:
                _error(name, f'Unknown table: {table}', kind='lookup')
            table_id, dimension = tables[table]
            if kind == 'measures':
                if (table, member) not in available_measures:
                    _error(name, f'Unknown measure: {table}[{member}]', kind='lookup')
                member_id = member
            else:
                group, tag = ('Attributes', 'Attribute') if kind == 'columns' else ('Hierarchies', 'Hierarchy')
                found = [a for a in dimension.findall(_tag(group) + '/' + _tag(tag))
                         if a.findtext(_tag('Name')) == member and a.findtext(_tag('Type')) != 'RowNumber']
                if len(found) != 1:
                    _error(name, f'Unknown {kind} member: {table}[{member}]', kind='lookup')
                member_id = found[0].findtext(_tag('ID'))
            entry = (table, table_id, member, member_id)
            if entry not in resolved:
                resolved.append(entry)
        members[kind] = resolved
    if not any(members.values()):
        _error(name, 'Perspective must include at least one object')
    if any(',' in member[2] for member in members['hierarchies']):
        _error(name, 'Hierarchy captions containing commas cannot be represented in the native perspective annotation',
               code='unsupported_perspective_hierarchy_name')
    cube_file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    cube_doc = minidom.parseString(image.read_logical_file(cube_file))
    cube = _elements(cube_doc, 'Cube')[0]
    annotation_members = _hierarchy_annotation_members(cube)
    has_annotation = bool(members['hierarchies']) or any(n.casefold() != name.casefold() for n in annotation_members)
    prospective_names = [n for n in list_perspectives(image) if n.casefold() != name.casefold()] + [name]
    if has_annotation and any("'" in n for n in prospective_names):
        _error(name, 'Perspective captions containing apostrophes are unsafe when hierarchy annotations exist',
               code='unsupported_perspective_annotation_name')
    dbid = _text(_elements(cube_doc, 'DatabaseID')[0])
    old = existing[0][1].find('.//' + _tag('Perspective')) if existing else None
    perspective_id = old.findtext(_tag('ID')) if old is not None else name
    object_id = old.findtext(_tag('ObjectID')) if old is not None else str(uuid.uuid4()).upper()
    # Renaming a caption intentionally retains its original physical identity.
    # Reusing that caption for a new perspective must not reuse the old file/ID.
    if old is None and any(
        r.findtext('.//' + _tag('Perspective') + '/' + _tag('ID'), '').casefold() == perspective_id.casefold()
        or f.logical_path.replace('\\', '/').rsplit('/', 1)[-1].casefold() == (name + '.0.persp.xml').casefold()
        for f, r in _xml(image)
    ):
        perspective_id = object_id
    timestamp = datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec='milliseconds')
    load = ET.Element(_tag('Load'))
    ET.SubElement(ET.SubElement(load, _tag('ParentObject')), _tag('DatabaseID')).text = dbid
    p = ET.SubElement(ET.SubElement(load, _tag('ObjectDefinition')), _tag('Perspective'))
    ordinal = old.findtext(_tag('Ordinal')) if old is not None else str(max(
        (int(r.findtext('.//' + _tag('Perspective') + '/' + _tag('Ordinal'), '0')) for _, r in _xml(image)), default=-1) + 1)
    for key, value in (('Name', name), ('ID', perspective_id), ('CreatedTimestamp', old.findtext(_tag('CreatedTimestamp')) if old is not None else timestamp),
                       ('LastSchemaUpdate', timestamp), ('ObjectVersion', old.findtext(_tag('ObjectVersion'), '0') if old is not None else '0'),
                       ('ObjectID', object_id), ('Ordinal', ordinal), ('PersistLocation', '0'), ('System', 'false'),
                       ('DataFileList', None), ('Description', description), ('Translations', None)):
        ET.SubElement(p, _tag(key)).text = value
    dimensions = ET.SubElement(p, _tag('Dimensions'))
    for table, (table_id, _) in tables.items():
        cols = [m for m in members['columns'] if m[0] == table]
        if not cols:
            continue
        d = ET.SubElement(dimensions, _tag('Dimension'))
        ET.SubElement(d, _tag('CubeDimensionID')).text = table_id
        ET.SubElement(d, _tag('Hierarchies'))
        attrs = ET.SubElement(d, _tag('Attributes'))
        for _, _, _, col_id in cols:
            a = ET.SubElement(attrs, _tag('Attribute'))
            ET.SubElement(a, _tag('AttributeID')).text = col_id
            ET.SubElement(a, _tag('AttributeHierarchyVisible')).text = 'true'
            ET.SubElement(a, _tag('DefaultMember'))
    ET.SubElement(p, _tag('Actions'))
    calculations = ET.SubElement(p, _tag('Calculations'))
    for member in [m[2] for m in members['measures']] or ['__XL_Count of Models']:
        c = ET.SubElement(calculations, _tag('Calculation'))
        ET.SubElement(c, _tag('Name')).text = '[Measures].[' + member.replace(']', ']]') + ']'
        ET.SubElement(c, _tag('Type')).text = 'Member'
    ET.SubElement(p, _tag('Kpis'))
    groups = ET.SubElement(p, _tag('MeasureGroups'))
    for table_id in dict.fromkeys(m[1] for m in members['measures']):
        g = ET.SubElement(groups, _tag('MeasureGroup'))
        ET.SubElement(g, _tag('MeasureGroupID')).text = table_id
        ET.SubElement(g, _tag('Measures'))
    ET.SubElement(p, _tag('DefaultMeasure'))
    if old is not None:
        # Membership edits must not discard translations or extensions authored
        # by Excel or another model tool.
        edited = {_tag(key) for key in ('Name', 'LastSchemaUpdate', 'Description',
                                        'Dimensions', 'Calculations', 'MeasureGroups')}
        p.attrib.update(old.attrib)
        _preserve_dimension_members(old.find(_tag('Dimensions')), dimensions)
        _preserve_retained_members(old.find(_tag('Calculations')), calculations,
                                  'Calculation', 'Name')
        _preserve_retained_members(old.find(_tag('MeasureGroups')), p.find(_tag('MeasureGroups')),
                                  'MeasureGroup', 'MeasureGroupID')
        for child in old:
            if child.tag not in edited:
                current = p.find(child.tag)
                if current is not None:
                    index = list(p).index(current)
                    p.remove(current)
                    p.insert(index, deepcopy(child))
                else:
                    p.append(deepcopy(child))
        # Preserve wrapper attributes/extensions too, not just the Perspective.
        load = deepcopy(existing[0][1])
        definition = load.find(_tag('ObjectDefinition'))
        previous = definition.find(_tag('Perspective'))
        index = list(definition).index(previous)
        definition.remove(previous)
        definition.insert(index, p)
    _update_hierarchy_annotation(cube_doc, cube, name, members['hierarchies'])
    replacements = {}
    additions, groups = (), ()
    if existing:
        replacements[existing[0][0].storage_path] = ET.tostring(load, encoding='utf-8')
    else:
        filename = perspective_id + '.0.persp.xml'
        listing = _child(cube, 'PerspectiveFileList')
        files = [v for v in _text(listing).split(';') if v] + [filename]
        _set_text(listing, ';'.join(files))
        group = next(g for g in image.backup_log_root.findall('FileGroups/FileGroup') if g.findtext('Class') == '100010')
        base = group.findtext('PersistLocationPath')
        anchor = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
        key = 'perspective-' + object_id
        groups = (VirtualFileGroupAddition(key, anchor.storage_path, perspective_id, name, base, object_id, '0', '100054'),)
        additions = (VirtualFileAddition(uuid.uuid4().hex.upper(), base + '\\' + filename,
                       anchor.storage_path, ET.tostring(load, encoding='utf-8'), b'\0\0\0\0', key),)
    replacements[cube_file.storage_path] = cube_doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements, additions=additions, groups=groups))


def _set_text(node, text):
    while node.firstChild:
        node.removeChild(node.firstChild)
    if text:
        node.appendChild(node.ownerDocument.createTextNode(text))


def _update_hierarchy_annotation(doc, cube, name, hierarchies):
    annotations = _child(cube, 'Annotations')
    found = [a for a in _elements(annotations, 'Annotation') if _text(_child(a, 'Name')) == 'PerspectiveEmptyHierarchyNames']
    root = ET.fromstring(_text(_child(found[0], 'Value'))) if found else ET.Element('Perspectives')
    previous = None
    for item in list(root):
        if item.get('Name', '').casefold() == name.casefold():
            previous = deepcopy(item)
            root.remove(item)
    if hierarchies:
        p = previous if previous is not None else ET.Element('Perspective')
        p.set('Name', name)
        root.append(p)
        t = p.find('Tables')
        if t is None:
            t = ET.SubElement(p, 'Tables')
        old_tables = {item.get('Name'): deepcopy(item) for item in t.findall('Table')}
        for item in list(t.findall('Table')):
            t.remove(item)
        grouped = {}
        for table, _, hierarchy, _ in hierarchies:
            grouped.setdefault(table, []).append(hierarchy)
        # Observed native serializer delimiter is comma; it has no escaping.
        for table in sorted(grouped):
            item = old_tables.get(table, ET.Element('Table'))
            item.set('Name', table)
            item.set('EmptyHierarchyNames', ','.join(grouped[table]))
            t.append(item)
    if not list(root):
        for a in found:
            annotations.removeChild(a)
        return
    if found:
        a = found[0]
    else:
        a = doc.createElementNS(ENGINE, 'Annotation')
        for tag in ('Name', 'Value'):
            a.appendChild(doc.createElementNS(ENGINE, tag))
        _set_text(_child(a, 'Name'), 'PerspectiveEmptyHierarchyNames')
        annotations.appendChild(a)
    _set_text(_child(a, 'Value'), '<?xml version="1.0" encoding="UTF-16"?>' + ET.tostring(root, encoding='unicode'))


def _hierarchy_annotation_members(cube):
    for annotation in _elements(_child(cube, 'Annotations'), 'Annotation'):
        if _text(_child(annotation, 'Name')) == 'PerspectiveEmptyHierarchyNames':
            root = ET.fromstring(_text(_child(annotation, 'Value')))
            return [p.get('Name', '') for p in root.findall('Perspective')]
    return []


def remove_perspective(image, name):
    _validate_name(name)
    match = [(f, r) for f, r in _xml(image) if r.findtext('.//' + _tag('Perspective') + '/' + _tag('Name')).casefold() == name.casefold()]
    if len(match) != 1:
        _error(name, f'Unknown perspective: {name}', kind='lookup')
    file, root = match[0]
    cube_file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(image.read_logical_file(cube_file))
    cube = _elements(doc, 'Cube')[0]
    listing = _child(cube, 'PerspectiveFileList')
    filename = file.logical_path.replace('\\', '/').rsplit('/', 1)[-1]
    _set_text(listing, ';'.join(v for v in _text(listing).split(';') if v and v != filename))
    _update_hierarchy_annotation(doc, cube, name, ())
    return AbfImage(image.assemble({cube_file.storage_path: doc.documentElement.toxml(encoding='utf-8')}, removals=(file.storage_path,)))


def rename_perspective(image, name, new_name):
    """Rename a caption while preserving identity, membership and unknown XML."""
    _validate_name(new_name)
    _validate_name(name)
    matches = [(f, r) for f, r in _xml(image)
               if r.findtext('.//' + _tag('Perspective') + '/' + _tag('Name')).casefold() == name.casefold()]
    if len(matches) != 1:
        _error(name, f'Unknown perspective: {name}', kind='lookup')
    if name.casefold() != new_name.casefold() and any(n.casefold() == new_name.casefold() for n in list_perspectives(image)):
        _error(new_name, 'Perspective already exists')
    file, _ = matches[0]
    doc = minidom.parseString(image.read_logical_file(file))
    _set_text(_child(_elements(doc, 'Perspective')[0], 'Name'), new_name)
    replacements = {file.storage_path: doc.documentElement.toxml(encoding='utf-8')}
    cube_file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    cube_doc = minidom.parseString(image.read_logical_file(cube_file))
    if "'" in new_name and _hierarchy_annotation_members(_elements(cube_doc, 'Cube')[0]):
        _error(new_name, 'Perspective captions containing apostrophes are unsafe when hierarchy annotations exist',
               code='unsupported_perspective_annotation_name')
    for annotation in _elements(cube_doc, 'Annotation'):
        if _text(_child(annotation, 'Name')) != 'PerspectiveEmptyHierarchyNames':
            continue
        value = _child(annotation, 'Value')
        nested = minidom.parseString(_text(value))
        for perspective in nested.getElementsByTagName('Perspective'):
            if perspective.getAttribute('Name').casefold() == name.casefold():
                perspective.setAttribute('Name', new_name)
        _set_text(value, nested.toxml())
    replacements[cube_file.storage_path] = cube_doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(replacements))
