"""Maintain the cube schema view without changing partition data sources."""
import json
import uuid
from xml.dom import minidom
from .abf import AbfImage, UnsupportedModelError, VirtualFileAddition, VirtualFileGroupAddition
from .schema import _elements, _child, _text

_MARKER = 'ExcelModelLab.CubeViewTables'
_XSI = 'http://www.w3.org/2001/XMLSchema-instance'

def _object(image, suffix, kind, identity=None):
    for file in image.files:
        if file.logical_path and file.logical_path.endswith(suffix):
            doc = minidom.parseString(image.read_logical_file(file))
            for obj in _elements(doc, kind):
                if getattr(obj.parentNode, 'localName', None) == 'ObjectDefinition' and (identity is None or _text(_child(obj, 'ID')) == identity):
                    return file, doc, obj
    raise UnsupportedModelError('Cannot resolve cube schema object: ' + str(identity or kind))

def _direct(node, name):
    return next((n for n in node.childNodes if getattr(n, 'localName', None) == name), None)

def _tables(view):
    schema = _child(view, 'Schema')
    choices = _elements(schema, 'choice')
    if len(choices) != 1:
        raise UnsupportedModelError('Cube source view must have one dataset table collection')
    return choices[0]

def _metadata(view):
    annotations = _direct(view, 'Annotations')
    if annotations is not None:
        for annotation in _elements(annotations, 'Annotation'):
            if _text(_child(annotation, 'Name')) == _MARKER:
                return json.loads(_text(_child(annotation, 'Value')))
    return {'owned': False, 'tables': {}}

def _store(view, data):
    from .ole_db import _append
    annotations = _direct(view, 'Annotations')
    if annotations is None:
        annotations = _append(view, 'Annotations')
        view.insertBefore(annotations, _child(view, 'Schema'))
    annotation = next((a for a in _elements(annotations, 'Annotation') if _text(_child(a, 'Name')) == _MARKER), None)
    if annotation is None:
        annotation = _append(annotations, 'Annotation')
        _append(annotation, 'Name', _MARKER)
        _append(annotation, 'Value')
    value = _child(annotation, 'Value')
    while value.firstChild: value.removeChild(value.firstChild)
    value.appendChild(view.ownerDocument.createTextNode(json.dumps(data, sort_keys=True)))

def _set(node, name, value):
    target = _child(node, name)
    while target.firstChild: target.removeChild(target.firstChild)
    target.appendChild(node.ownerDocument.createTextNode(value))

def _merge_owned_table(current, incoming):
    """Update generated schema fields without discarding native extensions."""
    xsd = 'http://www.w3.org/2001/XMLSchema'
    structural = {'element', 'complexType', 'simpleType', 'sequence'}

    def identity(node):
        return node.namespaceURI, node.localName, node.getAttribute('name')

    def merge(target, source):
        source_children = [n for n in source.childNodes if n.nodeType == n.ELEMENT_NODE]
        target_children = [n for n in target.childNodes if n.nodeType == n.ELEMENT_NODE]
        expected = {identity(n) for n in source_children}
        # Binding changes connection/query fields, not the model's column shape.
        # Preserve extension nodes, but reject competing schema declarations.
        if any(n.namespaceURI == xsd and n.localName in structural
               and identity(n) not in expected for n in target_children):
            raise UnsupportedModelError('Existing cube view table requires explicit schema reconciliation')
        for child in source_children:
            matches = [n for n in target_children if identity(n) == identity(child)]
            if len(matches) != 1:
                raise UnsupportedModelError('Existing cube view table requires explicit schema reconciliation')
            merge(matches[0], child)
        for attr in source.attributes.values():
            if attr.namespaceURI:
                target.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
            else:
                target.setAttribute(attr.name, attr.value)

    merged = current.cloneNode(True)
    merge(merged, incoming)
    return merged


def bind_cube_view(image, table_id, view_id, datasource_id):
    from .ole_db import _append, _load, _xml_name
    cube_file, cube_doc, cube = _object(image, '.cub.xml', 'Cube')
    source = _direct(cube, 'Source')
    template_file, template_doc, template = _object(image, '.dsv.xml', 'DataSourceView', view_id)
    additions, groups = [], []
    if source is None:
        identity = 'ExcelModelLabCubeView_' + uuid.uuid4().hex
        instance = str(uuid.uuid4()).upper()
        group = next(g for g in image.backup_log_root.findall('FileGroups/FileGroup') if g.findtext('Class') == '100002')
        doc, view = _load('DataSourceView', group.findtext('ID'), identity, identity, instance)
        view.appendChild(doc.importNode(_child(template, 'Schema'), True))
        _append(view, 'DataSourceID', datasource_id)
        # Start empty; the merge below adds exactly the requested table.
        for node in list(_tables(view).childNodes):
            if getattr(node, 'localName', None) == 'element': _tables(view).removeChild(node)
        metadata = {'owned': True, 'tables': {}}
        source = cube_doc.createElement('Source')
        source.setAttributeNS(_XSI, 'xsi:type', 'DataSourceViewBinding')
        _append(source, 'DataSourceViewID', identity)
        before = _direct(cube, 'Dimensions')
        cube.insertBefore(source, before)
        filename = identity + '.0.dsv.xml'
        key = 'cube-view-' + identity
        anchor = next(f.storage_path for f in image.files if f.logical_path and f.logical_path.endswith('.dim.xml'))
        groups.append(VirtualFileGroupAddition(key, anchor, identity, identity,
                      group.findtext('PersistLocationPath'), instance, '0', '100053'))
        new_file = (uuid.uuid4().hex.upper(), group.findtext('PersistLocationPath') + '\\' + filename, key)
        file = None
    else:
        identity = _text(_child(source, 'DataSourceViewID'))
        file, doc, view = _object(image, '.dsv.xml', 'DataSourceView', identity)
        metadata = _metadata(view)
    encoded = _xml_name(table_id)
    incoming = next((n for n in _tables(template).childNodes if getattr(n, 'localName', None) == 'element' and n.getAttribute('name') == encoded), None)
    if incoming is None: raise UnsupportedModelError('Table missing from its source view: ' + table_id)
    collection = _tables(view)
    current = next((n for n in collection.childNodes if getattr(n, 'localName', None) == 'element' and n.getAttribute('name') == encoded), None)
    if current is not None and (table_id not in metadata['tables'] or table_id in metadata.get('retained', [])):
        # A preexisting view may already describe the table; do not silently replace unknown schema.
        if current.toxml() != incoming.toxml():
            raise UnsupportedModelError('Existing cube view table requires explicit schema reconciliation: ' + table_id)
        if table_id not in metadata.setdefault('retained', []): metadata['retained'].append(table_id)
    merged = (_merge_owned_table(current, incoming) if current is not None
              else doc.importNode(incoming, True))
    if current is not None: collection.removeChild(current)
    merged.setAttribute('msprop:DataSourceID', datasource_id)
    collection.appendChild(merged)
    metadata['tables'][table_id] = datasource_id
    _store(view, metadata)
    replacements = {cube_file.storage_path: cube_doc.documentElement.toxml(encoding='utf-8')}
    payload = doc.documentElement.toxml(encoding='utf-8')
    if file is None:
        storage, logical, key = new_file
        additions.append(VirtualFileAddition(storage, logical, anchor, payload, b'\0'*4, key))
    else: replacements[file.storage_path] = payload
    # Excel discovers model-table queries from one cube DSV. Duplicate
    # per-source DSVs produce empty command text in native refresh.
    for dimension_file in image.files:
        if not dimension_file.logical_path or not dimension_file.logical_path.endswith('.dim.xml'):
            continue
        dimension_doc = minidom.parseString(image.read_logical_file(dimension_file))
        dimension = next((n for n in _elements(dimension_doc, 'Dimension')
                          if getattr(n.parentNode, 'localName', None) == 'ObjectDefinition'
                          and _text(_child(n, 'ID')) == table_id), None)
        if dimension is not None:
            binding = _direct(dimension, 'Source')
            if binding is not None and _direct(binding, 'DataSourceViewID') is not None:
                _set(binding, 'DataSourceViewID', identity)
                replacements[dimension_file.storage_path] = dimension_doc.documentElement.toxml(encoding='utf-8')
    removals = {template_file.storage_path} if view_id != identity else set()
    return AbfImage(image.assemble(replacements, additions=additions, groups=groups, removals=removals))

def detach_cube_view(image, table_id):
    from .ole_db import _xml_name
    cube_file, cube_doc, cube = _object(image, '.cub.xml', 'Cube')
    source = _direct(cube, 'Source')
    if source is None: return image
    file, doc, view = _object(image, '.dsv.xml', 'DataSourceView', _text(_child(source, 'DataSourceViewID')))
    metadata = _metadata(view)
    if table_id not in metadata['tables']: return image
    removed_source = metadata['tables'].pop(table_id)
    collection = _tables(view)
    if table_id not in metadata.get('retained', []):
        for node in list(collection.childNodes):
            if getattr(node, 'localName', None) == 'element' and node.getAttribute('name') == _xml_name(table_id): collection.removeChild(node)
    remaining = [n for n in collection.childNodes if getattr(n, 'localName', None) == 'element']
    if not remaining and metadata['owned']:
        cube.removeChild(source)
        return AbfImage(image.assemble({cube_file.storage_path:cube_doc.documentElement.toxml(encoding='utf-8')}, removals={file.storage_path}))
    if _text(_child(view, 'DataSourceID')) == removed_source and metadata['tables']:
        _set(view, 'DataSourceID', next(iter(metadata['tables'].values())))
    _store(view, metadata)
    return AbfImage(image.assemble({file.storage_path:doc.documentElement.toxml(encoding='utf-8')}))


def cube_view_id(image):
    _, _, cube = _object(image, '.cub.xml', 'Cube')
    return _text(_child(_child(cube, 'Source'), 'DataSourceViewID'))
