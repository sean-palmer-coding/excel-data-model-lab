"""Explicit Excel-native source bindings, separate from Python refresh sources.

Bindings never open a connection or refresh data. Linked-table bindings are based
on Excel's type-102 worksheet connection metadata. Snapshot data remains the
data supplied to the model until the caller deliberately refreshes in Excel.
"""
from dataclasses import dataclass, field
import json
import re
import uuid
from xml.dom import minidom
from zipfile import ZipFile

from .abf import AbfImage, UnsupportedModelError
from .schema import _elements, _child, _text

SHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
ENGINE = 'http://schemas.microsoft.com/analysisservices/2003/engine'
ANNOTATION = 'ExcelModelLab.SourceBindings.v1'
EXTENSION = '{DE250136-89BD-433C-8126-D09CA5730AF9}'
MODEL_EXTENSION = '{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}'


@dataclass(frozen=True)
class LinkedTableSource:
    """An existing worksheet table in the same workbook, identified by name."""
    table_name: str

    def __post_init__(self):
        if not isinstance(self.table_name, str) or not self.table_name.strip():
            raise ValueError('A linked table needs an existing worksheet table name')


@dataclass(frozen=True)
class OleDbSource:
    """Native query configuration; credentials must not be embedded by default."""
    connection_string: str = field(repr=False)
    query: str
    name: str = 'ExcelModelLab Source'

    def __post_init__(self):
        if not isinstance(self.connection_string, str) or not self.connection_string.strip():
            raise ValueError('An OLE DB connection string is required')
        if not isinstance(self.query, str) or not self.query.strip():
            raise ValueError('An OLE DB query is required')
        if re.search(r'(?:^|;)[^;=]*(?:password|pwd|token)[^;=]*=', self.connection_string, re.I):
            raise ValueError('Do not embed credentials in native sources; use integrated security or configure credentials in Excel')
        if not isinstance(self.name, str) or not self.name.strip():
            raise ValueError('An OLE DB source name is required')


@dataclass(frozen=True)
class NativeSourceBinding:
    table: str
    table_id: str
    kind: str
    source_name: str | None = None
    connection_name: str | None = None


def _document(image):
    file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('.cub.xml'))
    doc = minidom.parseString(image.read_logical_file(file))
    cube = next(n for n in _elements(doc, 'Cube') if getattr(n.parentNode, 'localName', None) == 'ObjectDefinition')
    return file, doc, cube


def binding_overrides(image):
    _, _, cube = _document(image)
    for annotation in _elements(_child(cube, 'Annotations'), 'Annotation'):
        if _text(_child(annotation, 'Name')) == ANNOTATION:
            result = json.loads(_text(_child(annotation, 'Value')))
            if not isinstance(result, dict):
                raise UnsupportedModelError('Invalid native source binding metadata')
            return result
    return {}


def _assign(node, value):
    while node.firstChild:
        node.removeChild(node.firstChild)
    node.appendChild(node.ownerDocument.createTextNode(value))


def store_binding(image, table_id, config):
    file, doc, cube = _document(image)
    configs = binding_overrides(image)
    configs[table_id] = config
    annotations = _child(cube, 'Annotations')
    matches = [a for a in _elements(annotations, 'Annotation') if _text(_child(a, 'Name')) == ANNOTATION]
    if matches:
        annotation = matches[0]
    else:
        annotation = doc.createElementNS(ENGINE, 'Annotation')
        for name in ('Name', 'Value'):
            annotation.appendChild(doc.createElementNS(ENGINE, name))
        _assign(_child(annotation, 'Name'), ANNOTATION)
        annotations.appendChild(annotation)
    _assign(_child(annotation, 'Value'), json.dumps(configs, ensure_ascii=False, sort_keys=True))
    return AbfImage(image.assemble({file.storage_path: doc.documentElement.toxml(encoding='utf-8')}))


def worksheet_tables(archive):
    result = {}
    for path in archive.namelist():
        if not path.startswith('xl/tables/') or not path.endswith('.xml'):
            continue
        root = minidom.parseString(archive.read(path)).documentElement
        name = root.getAttribute('displayName') or root.getAttribute('name')
        result[name.casefold()] = (name, tuple(n.getAttribute('name') for n in _elements(root, 'tableColumn')))
    return result


def bind_linked_table(model, table, source):
    table_id = model._dimension_ids[table]
    with ZipFile(model.source) as archive:
        available = worksheet_tables(archive)
    if source.table_name.casefold() not in available:
        raise KeyError('Worksheet table does not exist: ' + source.table_name)
    name, source_columns = available[source.table_name.casefold()]
    calculated = {c.name for c in model.list_calculated_columns() if c.table_id == table_id}
    model_columns = tuple(c.name for c in model.columns
                          if c.table == table and not c.name.startswith('__XL_') and c.name not in calculated)
    if tuple(c.casefold() for c in source_columns) != tuple(c.casefold() for c in model_columns):
        raise ValueError('Linked worksheet and model table columns must match in name and order')
    from .ole_db import detach_table_binding
    image = detach_table_binding(model._image, table_id)
    return store_binding(image, table_id, {'kind': 'linked_table', 'table_name': name})


def list_native_sources(model):
    bindings = {}
    physical = {identity: name for name, identity in model._dimension_ids.items()}
    with ZipFile(model.source) as archive:
        if 'xl/connections.xml' in archive.namelist():
            doc = minidom.parseString(archive.read('xl/connections.xml'))
            workbook = minidom.parseString(archive.read('xl/workbook.xml'))
            connected_tables = {n.getAttribute('connection'): n.getAttribute('id')
                                for n in _elements(workbook, 'modelTable')}
            names = {n.getAttribute('name'): _text(n) for n in _elements(workbook, 'definedName')}
            for connection in doc.documentElement.childNodes:
                if connection.nodeType != connection.ELEMENT_NODE:
                    continue
                details = [n for n in _elements(connection, 'connection') if n is not connection]
                for detail in details:
                    table_id = detail.getAttribute('id')
                    if table_id not in physical:
                        table_id = connected_tables.get(connection.getAttribute('name'))
                    if table_id not in physical:
                        continue
                    ranges = _elements(detail, 'rangePr')
                    kind = 'linked_table' if connection.getAttribute('type') == '102' else 'snapshot' if connection.getAttribute('type') == '104' else 'ole_db'
                    source = names.get(ranges[0].getAttribute('sourceName'), '')[:-2] if ranges else None
                    bindings[table_id] = NativeSourceBinding(physical[table_id], table_id, kind, source, connection.getAttribute('name'))
    for table_id, config in binding_overrides(model._image).items():
        if table_id in physical:
            previous = bindings.get(table_id)
            bindings[table_id] = NativeSourceBinding(physical[table_id], table_id, config['kind'],
                config.get('table_name'), config.get('connection_name') or
                (previous.connection_name if previous is not None else None))
    return tuple(sorted(bindings.values(), key=lambda binding: binding.table))


def binding_package_parts(original, replacements, model_bytes):
    """Apply explicit staged binding changes without modifying worksheet cells."""
    image = AbfImage(model_bytes)
    configs = binding_overrides(image)
    if not configs:
        return {}
    def read(path):
        return replacements.get(path) or original.read(path)
    workbook = minidom.parseString(read('xl/workbook.xml'))
    connections = minidom.parseString(read('xl/connections.xml'))
    tables = {}
    for file in image.files:
        if file.logical_path and file.logical_path.endswith('.dim.xml'):
            doc = minidom.parseString(image.read_logical_file(file))
            for dimension in _elements(doc, 'Dimension'):
                if getattr(dimension.parentNode, 'localName', None) == 'ObjectDefinition':
                    tables[_text(_child(dimension, 'ID'))] = _text(_child(dimension, 'Name'))
    def child(doc, parent, name, namespace=SHEET):
        matches = [n for n in parent.childNodes if getattr(n, 'localName', None) == name.split(':')[-1]]
        return matches[0] if matches else parent.appendChild(doc.createElementNS(namespace, name))
    defined = child(workbook, workbook.documentElement, 'definedNames')
    # definedNames precedes calcPr and extension lists in SpreadsheetML order.
    following = next((n for n in workbook.documentElement.childNodes
                      if getattr(n, 'localName', None) in ('calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'extLst')), None)
    if following and defined.nextSibling is None:
        workbook.documentElement.insertBefore(defined, following)
    extensions = child(workbook, workbook.documentElement, 'extLst')
    model_extension = next((n for n in _elements(extensions, 'ext') if n.getAttribute('uri') == MODEL_EXTENSION), None)
    if model_extension is None:
        model_extension = workbook.createElementNS(SHEET, 'ext')
        model_extension.setAttribute('uri', MODEL_EXTENSION)
        model_extension.setAttribute('xmlns:x15', X15)
        extensions.appendChild(model_extension)
    model_tables = child(workbook, child(workbook, model_extension, 'x15:dataModel', X15), 'x15:modelTables', X15)
    for table_id, config in configs.items():
        if config.get('kind') not in ('linked_table', 'snapshot', 'ole_db'):
            raise UnsupportedModelError('Unsupported native source binding kind')
        previous_names = {n.getAttribute('connection') for n in _elements(model_tables, 'modelTable')
                          if n.getAttribute('id') == table_id}
        existing = [n for n in list(connections.documentElement.childNodes)
                    if n.nodeType == n.ELEMENT_NODE and (n.getAttribute('name') in previous_names or
                        any(d.getAttribute('id') == table_id for d in _elements(n, 'connection') if d is not n))]
        previous_id = existing[0].getAttribute('id') if existing else None
        # A connection may also feed a worksheet query or another model table.
        # Replacing it would silently redirect those consumers, so require the
        # caller to separate that connection first.
        for old in existing:
            if any(n.getAttribute('id') != table_id and
                   n.getAttribute('connection') == old.getAttribute('name')
                   for n in _elements(model_tables, 'modelTable')):
                raise UnsupportedModelError('Cannot replace a connection shared by another model table')
            for part in original.namelist():
                if not part.endswith('.xml') or part in ('xl/connections.xml', 'xl/workbook.xml'):
                    continue
                payload = read(part)
                if b'connectionId' not in payload:
                    continue
                document = minidom.parseString(payload)
                if any(n.getAttribute('connectionId') == old.getAttribute('id')
                       for n in document.getElementsByTagName('*')):
                    raise UnsupportedModelError('Cannot replace a connection referenced by workbook part: ' + part)
        for connection in existing:
            # Defined names can also be referenced by worksheet formulas or
            # other connections. Retain them; they remain valid worksheet
            # references even after this model binding is detached.
            connections.documentElement.removeChild(connection)
        if table_id not in tables:
            for entry in list(_elements(model_tables, 'modelTable')):
                if entry.getAttribute('id') == table_id:
                    model_tables.removeChild(entry)
            continue
        number = previous_id or str(max([int(n.getAttribute('id')) for n in connections.documentElement.childNodes
                                         if n.nodeType == n.ELEMENT_NODE] + [0]) + 1)
        connection_name = config.get('connection_name') or (existing[0].getAttribute('name') if existing else None) or 'ExcelModelLab_' + table_id
        if existing and connection_name != existing[0].getAttribute('name'):
            from .dependencies import report_dependency_parts
            if any(p == 'xl/workbook.xml' or p.startswith('xl/worksheets/')
                   for p in report_dependency_parts(original)):
                raise UnsupportedModelError('Cannot rename a source connection with unresolved CUBE formula bindings')
        if any(n.nodeType == n.ELEMENT_NODE and n.getAttribute('name').casefold() == connection_name.casefold()
               for n in connections.documentElement.childNodes):
            connection_name += '_' + table_id
        connection = connections.createElementNS(SHEET, 'connection')
        if existing:
            old = existing[0]
            for attribute in old.attributes.values():
                connection.setAttribute(attribute.name, attribute.value)
            # Preserve vendor extensions outside the native binding extension.
            for old_ext in _elements(old, 'ext'):
                if old_ext.getAttribute('uri') != EXTENSION:
                    child(connections, connection, 'extLst').appendChild(old_ext.cloneNode(deep=True))
        linked = config['kind'] == 'linked_table'
        ole = config['kind'] == 'ole_db'
        for key, value in {'id': number, 'name': connection_name, 'type': '102' if linked else '100' if ole else '104',
                           'refreshedVersion': '8', 'minRefreshableVersion': '5', 'refreshOnLoad': '0'}.items():
            connection.setAttribute(key, value)
        ext = connections.createElementNS(SHEET, 'ext')
        child(connections, connection, 'extLst').appendChild(ext)
        ext.setAttribute('uri', EXTENSION)
        ext.setAttribute('xmlns:x15', X15)
        detail = child(connections, ext, 'x15:connection', X15)
        detail.setAttribute('id', config['data_source_id'] if ole else table_id)
        if linked:
            hidden_name = '_xlcn.ExcelModelLab_' + uuid.uuid5(uuid.NAMESPACE_URL,
                table_id + '\0' + config['table_name'].casefold()).hex
            range_pr = child(connections, detail, 'x15:rangePr', X15)
            range_pr.setAttribute('sourceName', hidden_name)
            definition = next((n for n in _elements(defined, 'definedName')
                               if n.getAttribute('name') == hidden_name), None)
            if definition is None:
                definition = workbook.createElementNS(SHEET, 'definedName')
                defined.appendChild(definition)
            definition.setAttribute('name', hidden_name)
            definition.setAttribute('hidden', '1')
            _assign(definition, config['table_name'] + '[]')
        connections.documentElement.appendChild(connection)
        entries = [n for n in _elements(model_tables, 'modelTable') if n.getAttribute('id') == table_id]
        entry = entries[0] if entries else model_tables.appendChild(workbook.createElementNS(X15, 'x15:modelTable'))
        for key, value in {'id': table_id, 'name': tables[table_id], 'connection': connection_name}.items():
            entry.setAttribute(key, value)
    return {'xl/workbook.xml': workbook.toxml(encoding='utf-8'),
            'xl/connections.xml': connections.toxml(encoding='utf-8')}
