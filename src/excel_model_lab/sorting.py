"""Column ordering metadata and its intrinsic hierarchy indexes."""
from .collation import text_key
from xml.dom import minidom
import pandas as pd

from .abf import AbfImage, UnsupportedModelError
from .schema import _elements, _child, _text
from .lookup import string_lookup_replacements, numeric_lookup_replacements
from .hash_index import encode_numeric_hash_index


def stored_sorts(model):
    result = []
    for table, path in model._dimension_paths.items():
        doc = minidom.parseString(model._image.read_logical_file(path))
        attributes = [a for a in _child(_elements(doc, 'Dimension')[0], 'Attributes').childNodes
                      if getattr(a, 'localName', None) == 'Attribute']
        captions = {_text(_child(a, 'ID')): _text(_child(a, 'Name')) for a in attributes}
        for attribute in attributes:
            sort_nodes = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'OrderByAttributeID']
            if not sort_nodes:
                continue
            target = _text(sort_nodes[0])
            if target:
                result.append((table, _text(_child(attribute, 'Name')), captions.get(target, target)))
    return result


def set_sort_metadata(model, table, column, by):
    physical = model._dimension_ids[table]
    dimension = minidom.parseString(model._image.read_logical_file(model._dimension_paths[table]))
    attributes = [a for a in _child(_elements(dimension, 'Dimension')[0], 'Attributes').childNodes
                  if getattr(a, 'localName', None) == 'Attribute']
    identities = {_text(_child(a, 'Name')): _text(_child(a, 'ID')) for a in attributes}
    column = identities.get(column, column)
    by = identities.get(by, by)
    replacements = {}
    def assign(node, value):
        while node.firstChild:
            node.removeChild(node.firstChild)
        if value:
            node.appendChild(node.ownerDocument.createTextNode(value))
    for file in model._image.files:
        if not file.logical_path:
            continue
        filename = file.logical_path.split('\\')[-1]
        if not filename.startswith(physical + '.') or not filename.endswith(('.dim.xml', '.tbl.xml')):
            continue
        doc = minidom.parseString(model._image.read_logical_file(file))
        if filename.endswith('.dim.xml'):
            attr = next(n for n in _elements(doc, 'Attribute') if any(getattr(c, 'localName', None) == 'ID' and _text(c) == column for c in n.childNodes))
            assign(_child(attr, 'OrderBy'), 'AttributeKey' if by else 'Key')
            assign(_child(attr, 'OrderByAttributeID'), by)
        else:
            raw = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') == column)
            assign(_child(_child(raw, 'Properties'), 'OrderByColumn'), by)
        replacements[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(model._image.assemble(replacements))


def normalized(value):
    if pd.isna(value):
        return (0, '')
    return (1, text_key(value) if isinstance(value, str) else value)


def validate_sort_values(values, keys, table, column):
    mapping = {}
    for value, key in zip(values, keys):
        identity, order = normalized(value), normalized(key)
        if identity in mapping and mapping[identity] != order:
            raise ValueError(f'Each value in {table}.{column} must have exactly one sort value')
        mapping[identity] = order
    return mapping


def rebuild_sort(model, table, column, by):
    from .model import _StringEncoding, _ValueEncoding
    with model._open_reader() as reader:
        frame = reader.get_table(table)
    encoding = model._encodings[(table, column)]
    values = frame[column].tolist()
    keys = frame[by].tolist() if by else values
    mapping = validate_sort_values(values, keys, table, column)
    image = model._image
    if isinstance(encoding, _StringEncoding):
        ids, dictionary = encoding.encode(values)
        if 2 in ids and mapping.get((0, '')) != (0, ''):
            raise UnsupportedModelError('A blank label currently requires a blank sort key')
        ordered = sorted(range(len(dictionary)), key=lambda i: (mapping.get(normalized(dictionary[i]), (2, '')), text_key(dictionary[i])))
        replacements = string_lookup_replacements(image, model._dimension_ids[table], model._column_ids[(table, column)], dictionary,
                                                  has_nulls=2 in ids, ordered=ordered)
    elif isinstance(encoding, _ValueEncoding):
        ids = encoding.encode(values)
        by_id = {data_id: mapping[normalized(value)] for data_id, value in zip(ids, values)}
        ordered = sorted(by_id, key=lambda data_id: (by_id[data_id], data_id))
        replacements = numeric_lookup_replacements(image, model._dimension_ids[table], model._column_ids[(table, column)], ordered, preserve_order=True)
        replacements[encoding.hash_index_path] = encode_numeric_hash_index(image.read_logical_file(encoding.hash_index_path), ordered, preserve_order=True)
    else:
        raise UnsupportedModelError('Column sorting requires string or value-encoded numeric storage')
    return AbfImage(image.assemble(replacements))
