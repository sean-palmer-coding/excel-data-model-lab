"""Native relational bindings for cached Excel model tables.

The writer declares a schema from the existing snapshot and does not connect to
the database. Excel refresh requires a compatible installed provider and a query
that returns the declared columns in the declared types.
"""
from datetime import datetime, timezone
import re
import uuid
from xml.dom import minidom

from .abf import AbfImage, VirtualFileAddition, VirtualFileGroupAddition, UnsupportedModelError
from .native_sources import ENGINE, binding_overrides, store_binding
from .schema import _child, _elements, _text
from .source_crypto import encrypt_string

XSI = 'http://www.w3.org/2001/XMLSchema-instance'
XS = 'http://www.w3.org/2001/XMLSchema'


def _append(parent, tag_name, value=None, **attributes):
    node = parent.ownerDocument.createElement(tag_name)
    for key, attribute in attributes.items():
        node.setAttribute(key, attribute)
    if value is not None:
        node.appendChild(parent.ownerDocument.createTextNode(str(value)))
    parent.appendChild(node)
    return node


def _xml_name(value):
    return ''.join(c if (c.isalpha() or c == '_' or i > 0 and (c.isdigit() or c in '.-'))
                   and not (c == '_' and re.match(r'_x[0-9a-fA-F]{4}_', value[i:]))
                   else ''.join(f'_x{unit:04X}_' for unit in ([ord(c)] if ord(c) <= 0xffff else
                        [int.from_bytes(c.encode('utf-16le')[i:i+2], 'little') for i in (0, 2)]))
                   for i, c in enumerate(value))


def _load(kind, database_id, name, identifier, instance):
    doc = minidom.parseString(f'<Load xmlns="{ENGINE}" xmlns:xsi="{XSI}" xmlns:xs="{XS}"/>')
    _append(_append(doc.documentElement, 'ParentObject'), 'DatabaseID', database_id)
    obj = _append(_append(doc.documentElement, 'ObjectDefinition'), kind)
    now = datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec='milliseconds')
    for field, value in [('Name', name), ('ID', identifier), ('CreatedTimestamp', now),
                         ('LastSchemaUpdate', now), ('ObjectVersion', '0'), ('ObjectID', instance),
                         ('Ordinal', '0'), ('PersistLocation', '0'), ('System', 'false'),
                         ('DataFileList', ''), ('Description', '')]:
        _append(obj, field, value)
    return doc, obj


def _replace_source(parent, source, before=None):
    for existing in list(parent.childNodes):
        if getattr(existing, 'localName', None) == 'Source':
            parent.removeChild(existing)
    if before is not None:
        parent.insertBefore(source, before)
    else:
        parent.appendChild(source)


def _rebind_datasource(existing, generated):
    """Replace managed connection/auth fields and retain unrelated metadata."""
    document = existing.ownerDocument.cloneNode(True)
    identity = _text(_child(existing, 'ID'))
    matches = [n for n in _elements(document, 'DataSource')
               if getattr(n.parentNode, 'localName', None) == 'ObjectDefinition'
               and _text(_child(n, 'ID')) == identity]
    if len(matches) != 1:
        raise UnsupportedModelError('Cannot reconcile ambiguous data source metadata')
    target = matches[0]

    def children(node, name):
        return [n for n in node.childNodes if getattr(n, 'localName', None) == name
                and n.namespaceURI in (None, ENGINE)]

    order = [n.localName for n in generated.childNodes if n.nodeType == n.ELEMENT_NODE]

    def put(name, replacement):
        old = children(target, name)
        if len(old) > 1:
            raise UnsupportedModelError('Cannot reconcile duplicate data source field: ' + name)
        if old:
            if replacement is None:
                target.removeChild(old[0])
            else:
                target.replaceChild(document.importNode(replacement, True), old[0])
        elif replacement is not None:
            later = set(order[order.index(name) + 1:])
            before = next((n for n in target.childNodes if getattr(n, 'localName', None) in later), None)
            target.insertBefore(document.importNode(replacement, True), before)

    for name in ('Name', 'LastSchemaUpdate', 'ConnectionString', 'ImpersonationInfo'):
        put(name, _child(generated, name))
    # Native save can add a removed-password security marker. It describes the
    # old connection and must not survive a caller-supplied replacement string.
    put('ConnectionStringSecurity', None)
    annotations = children(target, 'Annotations')
    if len(annotations) > 1:
        raise UnsupportedModelError('Cannot reconcile duplicate data source annotations')
    fresh_annotations = _child(generated, 'Annotations')
    if not annotations:
        put('Annotations', fresh_annotations)
    else:
        managed = next(n for n in _elements(fresh_annotations, 'Annotation')
                       if _text(_child(n, 'Name')) == 'ConnectionEditUISource')
        old = [n for n in children(annotations[0], 'Annotation')
               if _text(_child(n, 'Name')) == 'ConnectionEditUISource']
        if len(old) > 1:
            raise UnsupportedModelError('Cannot reconcile duplicate connection UI annotation')
        if old:
            annotations[0].replaceChild(document.importNode(managed, True), old[0])
        else:
            annotations[0].appendChild(document.importNode(managed, True))
    return document, target


def bind_ole_db(model, table, source):
    image, table_id = model._image, model._dimension_ids[table]
    config = binding_overrides(image).get(table_id, {})
    datasource_id = config.get('data_source_id') or str(uuid.uuid4())
    # Build a temporary per-table schema, then merge into the one cube view.
    view_id = 'ExcelModelLabView_' + uuid.uuid4().hex
    db_group = next(g for g in image.backup_log_root.findall('FileGroups/FileGroup') if g.findtext('Class') == '100002')
    database_id, base = db_group.findtext('ID'), db_group.findtext('PersistLocationPath')
    key_file = next(f for f in image.files if f.logical_path and f.logical_path.endswith('CryptKey.bin'))
    key_data = image.read_logical_file(key_file)
    ds_instance, view_instance = str(uuid.uuid4()).upper(), str(uuid.uuid4()).upper()
    old_headers = {}
    old_files = {}
    for file in image.files:
        if file.logical_path and file.logical_path.endswith(('.ds.xml', '.dsv.xml')):
            document = minidom.parseString(image.read_logical_file(file))
            for kind in ('DataSource', 'DataSourceView'):
                for obj in _elements(document, kind):
                    if getattr(obj.parentNode, 'localName', None) == 'ObjectDefinition':
                        old_headers[_text(_child(obj, 'ID'))] = obj
                        old_files[_text(_child(obj, 'ID'))] = file
    if datasource_id in old_headers:
        ds_instance = _text(_child(old_headers[datasource_id], 'ObjectID'))
    if view_id in old_headers:
        view_instance = _text(_child(old_headers[view_id], 'ObjectID'))
    ds_doc, ds = _load('DataSource', database_id, source.name, datasource_id, ds_instance)
    ds.setAttributeNS(XSI, 'xsi:type', 'RelationalDataSource')
    annotations = _append(ds, 'Annotations')
    annotation = _append(annotations, 'Annotation')
    _append(annotation, 'Name', 'ConnectionEditUISource')
    _append(annotation, 'Value', 'SqlServer')
    _append(ds, 'PermissionFileList')
    _append(ds, 'ConnectionString', '[EncryptedDataStart]' + encrypt_string(key_data, source.connection_string) + '[EncryptedDataEnd]')
    _append(ds, 'Timeout', 'P0D')
    impersonation = _append(ds, 'ImpersonationInfo')
    _append(impersonation, 'ImpersonationMode', 'ImpersonateCurrentUser')
    _append(impersonation, 'Account')
    _append(impersonation, 'Password', encrypt_string(key_data, '', flags=0))
    _append(ds, 'MaxActiveConnections', '10')
    _append(ds, 'Isolation', 'ReadCommitted')
    if datasource_id in old_headers:
        ds_doc, ds = _rebind_datasource(old_headers[datasource_id], ds)

    view_doc, view = _load('DataSourceView', database_id, view_id, view_id, view_instance)
    schema = _append(view, 'Schema')
    xsd = _append(schema, 'xs:schema', id='NewDataSet')
    xsd.setAttribute('xmlns', '')
    xsd.setAttribute('xmlns:msdata', 'urn:schemas-microsoft-com:xml-msdata')
    xsd.setAttribute('xmlns:msprop', 'urn:schemas-microsoft-com:xml-msprop')
    dataset = _append(xsd, 'xs:element', name='NewDataSet')
    dataset.setAttribute('msdata:IsDataSet', 'true')
    dataset.setAttribute('msdata:Locale', 'en-US')
    choice = _append(_append(dataset, 'xs:complexType'), 'xs:choice', minOccurs='0', maxOccurs='unbounded')
    schema_table = _append(choice, 'xs:element', name=_xml_name(table_id))
    for key, value in {'msdata:Locale': '', 'msprop:IsLogical': 'True', 'msprop:Description': table,
                       'msprop:QueryDefinition': source.query, 'msprop:FriendlyName': table,
                       'msprop:TableType': 'View', 'msprop:DbTableName': table}.items():
        schema_table.setAttribute(key, value)
    sequence = _append(_append(schema_table, 'xs:complexType'), 'xs:sequence')
    dim_file = next(f for f in image.files if f.storage_path == model._dimension_paths[table])
    dim_doc = minidom.parseString(image.read_logical_file(dim_file))
    dimension = _elements(dim_doc, 'Dimension')[0]
    dimension_source = dim_doc.createElement('Source')
    dimension_source.setAttributeNS(XSI, 'xsi:type', 'DataSourceViewBinding')
    _append(dimension_source, 'DataSourceViewID', view_id)
    _replace_source(dimension, dimension_source, _child(dimension, 'MiningModelID'))
    types = {'BigInt': 'xs:long', 'Integer': 'xs:int', 'Double': 'xs:double',
             'Currency': 'xs:decimal', 'Date': 'xs:dateTime', 'Boolean': 'xs:boolean', 'WChar': 'xs:string'}
    columns = []
    for attribute in _elements(_child(dimension, 'Attributes'), 'Attribute'):
        if _text(_child(attribute, 'Type')) == 'RowNumber' or _elements(attribute, 'Expression'):
            continue
        name, identifier = _text(_child(attribute, 'Name')), _text(_child(attribute, 'ID'))
        datatype = _text(_child(_elements(attribute, 'KeyColumn')[0], 'DataType'))
        if datatype not in types:
            raise UnsupportedModelError('Unsupported OLE DB snapshot type: ' + datatype)
        columns.append(identifier)
        field = _append(sequence, 'xs:element', name=_xml_name(identifier), type=types[datatype], minOccurs='0')
        field.setAttribute('msprop:FriendlyName', name)
        field.setAttribute('msprop:DbColumnName', name)
        for container in _elements(attribute, 'KeyColumn') + _elements(attribute, 'NameColumn'):
            binding = dim_doc.createElement('Source')
            binding.setAttributeNS(XSI, 'xsi:type', 'ColumnBinding')
            _append(binding, 'TableID', table_id)
            _append(binding, 'ColumnID', identifier)
            _replace_source(container, binding)
    empty = _append(schema, 'NewDataSet')
    empty.setAttribute('xmlns', '')
    _append(view, 'DataSourceID', datasource_id)
    for identifier, obj in ((datasource_id, ds), (view_id, view)):
        if identifier in old_headers:
            for field in ('CreatedTimestamp', 'ObjectVersion', 'Ordinal'):
                node = _child(obj, field)
                while node.firstChild:
                    node.removeChild(node.firstChild)
                node.appendChild(obj.ownerDocument.createTextNode(_text(_child(old_headers[identifier], field))))
    replacements = {dim_file.storage_path: dim_doc.documentElement.toxml(encoding='utf-8')}
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.prt.xml', '.det.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        kind = 'Partition' if file.logical_path.endswith('.prt.xml') else 'MeasureGroup'
        obj = next((n for n in _elements(doc, kind) if getattr(n.parentNode, 'localName', None) == 'ObjectDefinition'
                    and _text(_child(n, 'ID')) == table_id), None)
        if obj is None:
            continue
        if kind == 'Partition':
            binding = doc.createElement('Source')
            binding.setAttributeNS(XSI, 'xsi:type', 'QueryBinding')
            _append(binding, 'DataSourceID', datasource_id)
            _append(binding, 'QueryDefinition', source.query)
            _replace_source(obj, binding, _child(obj, 'ProactiveCaching'))
        else:
            for attribute in _elements(obj, 'Attribute'):
                identifier = _text(_child(attribute, 'AttributeID'))
                if identifier not in columns:
                    continue
                for key in _elements(attribute, 'KeyColumn'):
                    binding = doc.createElement('Source')
                    binding.setAttributeNS(XSI, 'xsi:type', 'ColumnBinding')
                    _append(binding, 'TableID', table_id)
                    _append(binding, 'ColumnID', identifier)
                    _replace_source(key, binding)
        replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')

    additions, groups = [], []
    for identifier, instance, xml, suffix, class_id in (
            (datasource_id, ds_instance, ds_doc, '.ds.xml', '100003'),
            (view_id, view_instance, view_doc, '.dsv.xml', '100053')):
        filename = identifier + '.0' + suffix
        existing = old_files.get(identifier)
        data = xml.documentElement.toxml(encoding='utf-8')
        if existing:
            replacements[existing.storage_path] = data
        else:
            group_key = 'source-' + identifier
            groups.append(VirtualFileGroupAddition(group_key, dim_file.storage_path, identifier,
                          source.name if suffix == '.ds.xml' else identifier, base, instance, '0', class_id))
            additions.append(VirtualFileAddition(uuid.uuid4().hex.upper(), base + '\\' + filename,
                             dim_file.storage_path, data, b'\0' * 4, group_key))
    image = AbfImage(image.assemble(replacements, additions=additions, groups=groups))
    from .cube_source_view import bind_cube_view
    image = bind_cube_view(image, table_id, view_id, datasource_id)
    from .cube_source_view import cube_view_id
    view_id = cube_view_id(image)
    return store_binding(image, table_id, {'kind': 'ole_db', 'data_source_id': datasource_id,
                                         'view_id': view_id, 'connection_name': source.name})


def detach_table_binding(image, table_id):
    """Detach query/column refresh bindings while preserving stored values."""
    from .cube_source_view import detach_cube_view
    image = detach_cube_view(image, table_id)
    replacements = {}
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith(('.dim.xml', '.prt.xml', '.det.xml')):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        kind = 'Dimension' if file.logical_path.endswith('.dim.xml') else 'Partition' if file.logical_path.endswith('.prt.xml') else 'MeasureGroup'
        obj = next((n for n in _elements(doc, kind) if getattr(n.parentNode, 'localName', None) == 'ObjectDefinition'
                    and _text(_child(n, 'ID')) == table_id), None)
        if obj is None:
            continue
        changed = False
        for node in list(_elements(obj, 'Source')):
            binding = node.getAttributeNS(XSI, 'type').split(':')[-1]
            if binding in ('DataSourceViewBinding', 'QueryBinding', 'ColumnBinding', 'RowBinding'):
                node.parentNode.removeChild(node)
                changed = True
        if changed:
            replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    config = binding_overrides(image).get(table_id, {})
    owned = {value for key, value in config.items() if key in ('data_source_id', 'view_id')}
    removals = set()
    if owned:
        referenced = set()
        candidates = {}
        candidate_references = {}
        for file in image.files:
            if not file.logical_path or not file.logical_path.endswith('.xml'):
                continue
            doc = minidom.parseString(replacements.get(file.storage_path, image.read_logical_file(file)))
            identity = next((_text(_child(obj, 'ID')) for kind in ('DataSource', 'DataSourceView')
                             for obj in _elements(doc, kind)
                             if getattr(obj.parentNode, 'localName', None) == 'ObjectDefinition'), None)
            links = {_text(node) for tag in ('DataSourceID', 'DataSourceViewID') for node in _elements(doc, tag)}
            links.update(attribute.value for node in doc.getElementsByTagName('*')
                         for attribute in node.attributes.values()
                         if attribute.localName in ('DataSourceID', 'DataSourceViewID'))
            if identity in owned:
                candidates[file.storage_path] = identity
                candidate_references[identity] = links
            else:
                referenced.update(links)
        # A retained shared view may reference an otherwise owned data source.
        # Follow candidate dependencies from live roots before removing anything.
        while True:
            expanded = referenced | {target for owner, links in candidate_references.items()
                                     if owner in referenced for target in links}
            if expanded == referenced:
                break
            referenced = expanded
        # The shared cube view can remain referenced while this table's
        # independent data source is no longer used. Remove independently.
        removals.update(path for path, identifier in candidates.items() if identifier not in referenced)
    return AbfImage(image.assemble(replacements, removals=removals)) if replacements or removals else image
