"""Transfer a stored column from the synthetic authoring backend."""
from xml.dom import minidom
import re
import uuid

from .abf import AbfImage, UnsupportedModelError, VirtualFileAddition
from .schema import _elements, _child, _text


def import_column(image: AbfImage, donor: AbfImage, table_id: str, column: str) -> AbfImage:
    def inventory(source):
        return {f.logical_path.split('\\')[-1]: f for f in source.files if f.logical_path}
    target_files, source_files = inventory(image), inventory(donor)
    def one(files, pattern):
        matches = [f for name, f in files.items() if re.fullmatch(pattern, name)]
        if len(matches) != 1:
            raise UnsupportedModelError(f'Cannot resolve column metadata: {pattern}')
        return matches[0]
    docs, sources, targets = {}, {}, {}
    for kind in ('dim', 'det', 'tbl', 'cub'):
        pattern = re.escape(table_id if kind != 'cub' else 'Model') + r'\.\d+\.' + kind + r'\.xml'
        target, source = one(target_files, pattern), one(source_files, pattern)
        targets[kind], sources[kind] = target, minidom.parseString(donor.read_logical_file(source))
        docs[target.storage_path] = minidom.parseString(image.read_logical_file(target))
    dimension = _elements(docs[targets['dim'].storage_path], 'Dimension')[0]
    source_dimension = _elements(sources['dim'], 'Dimension')[0]
    attrs = _child(dimension, 'Attributes')
    source_attrs = [n for n in _child(source_dimension, 'Attributes').childNodes if n.nodeType == n.ELEMENT_NODE]
    source_attr = next(n for n in source_attrs if _text(_child(n, 'ID')) == column)
    caption = _text(_child(source_attr, 'Name'))
    if any(_text(_child(n, 'Name')).casefold() == caption.casefold()
           or _text(_child(n, 'ID')).casefold() == column.casefold()
           for n in attrs.childNodes if n.nodeType == n.ELEMENT_NODE):
        raise ValueError(f'Column already exists: {caption}')
    cloned = attrs.ownerDocument.importNode(source_attr, True)
    used = {_text(_child(a, 'Value')) for a in _elements(dimension, 'Annotation') if _text(_child(a, 'Name')) == 'ShortColumnId'}
    index = 1
    while True:
        short_id, number = '', index
        while number:
            number, remainder = divmod(number - 1, 26)
            short_id = chr(65 + remainder) + short_id
        if short_id not in used:
            break
        index += 1
    for annotation in _elements(cloned, 'Annotation'):
        if _text(_child(annotation, 'Name')) == 'ShortColumnId':
            _child(annotation, 'Value').firstChild.data = short_id
    attrs.appendChild(cloned)
    target_row = next(a for a in _elements(dimension, 'Attribute')
                      if any(n.localName == 'Type' and _text(n) == 'RowNumber' for n in a.childNodes))
    source_relation = next(a for a in _elements(source_dimension, 'AttributeRelationship') if _text(_child(a, 'AttributeID')) == column)
    relation = attrs.ownerDocument.importNode(source_relation, True)
    used_names = {_text(_child(r, 'Name')).casefold() for r in _elements(dimension, 'AttributeRelationship')}
    property_name = column + '_'
    while property_name.casefold() in used_names:
        property_name += '_'
    _child(relation, 'Name').firstChild.data = property_name
    _child(target_row, 'AttributeRelationships').appendChild(relation)
    for kind in ('cub', 'det'):
        def own_dimension(doc):
            return next(d for d in _elements(doc, 'Dimension') if any(n.nodeType == n.ELEMENT_NODE
                and n.localName in ('DimensionID', 'CubeDimensionID') and _text(n) == table_id for n in d.childNodes))
        target = own_dimension(docs[targets[kind].storage_path])
        source = own_dimension(sources[kind])
        attr = next(a for a in _elements(source, 'Attribute') if _text(_child(a, 'AttributeID')) == column)
        _child(target, 'Attributes').appendChild(target.ownerDocument.importNode(attr, True))
    raw = next(n for n in _elements(sources['tbl'], 'XMObject') if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == column)
    target_raw = next(n for n in _elements(docs[targets['tbl'].storage_path], 'XMObject') if n.getAttribute('class') == 'XMRawColumn')
    target_raw.parentNode.appendChild(target_raw.ownerDocument.importNode(raw, True))
    def info(source, dim):
        prefix = dim.logical_path.rsplit('\\', 1)[0] + '\\' + table_id + '.0.dim\\'
        return next(f for f in source.files if f.logical_path and f.logical_path.startswith(prefix)
                    and re.fullmatch(r'info\.\d+\.xml', f.logical_path.split('\\')[-1]))
    source_dim_file = one(source_files, re.escape(table_id) + r'\.\d+\.dim\.xml')
    target_info = info(image, targets['dim'])
    info_doc = minidom.parseString(image.read_logical_file(target_info))
    source_info = minidom.parseString(donor.read_logical_file(info(donor, source_dim_file)))
    prop = _elements(source_info, 'Property')[source_attrs.index(source_attr)]
    _elements(info_doc, 'Property')[0].parentNode.appendChild(info_doc.importNode(prop, True))
    docs[target_info.storage_path] = info_doc
    copied = [f for name, f in source_files.items() if f'.{table_id}.{column}.' in name or f'H${table_id}${column}.' in name]
    _child(dimension, 'DataFileList').firstChild.data += ';' + ';'.join(f.logical_path.split('\\')[-1] for f in copied)
    target_groups = image.backup_log_root.findall('FileGroups/FileGroup')
    additions = []
    for f in copied:
        group = next(g for g in donor.backup_log_root.findall('FileGroups/FileGroup')
                     if any(n.findtext('StoragePath') == f.storage_path for n in g.findall('FileList/BackupFile')))
        target_group = next(g for g in target_groups if g.findtext('ID') == table_id and g.findtext('Class') == group.findtext('Class'))
        source_base, target_base = group.findtext('PersistLocationPath'), target_group.findtext('PersistLocationPath')
        if not f.logical_path.startswith(source_base):
            raise UnsupportedModelError('Unexpected column storage path')
        path = target_base + f.logical_path[len(source_base):]
        parent = path.rsplit('\\', 1)[0]
        anchor = next(n.findtext('StoragePath') for n in target_group.findall('FileList/BackupFile')
                      if n.findtext('Path').rsplit('\\', 1)[0] == parent)
        additions.append(VirtualFileAddition(uuid.uuid4().hex.upper()[:20], path,
            anchor, donor.read_logical_file(f),
            donor.read_file(f)[-4:] if image.error_code and donor.error_code else None))
    return AbfImage(image.assemble({key: doc.documentElement.toxml(encoding='utf-8') for key, doc in docs.items()}, additions=additions))
