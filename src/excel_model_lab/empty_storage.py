"""Native zero-row storage, derived from a synthetic Power Pivot reference."""
from .reader_adapter import storage_schema, segment_metadata, decode_dictionary
from pathlib import Path
import struct
from xml.dom import minidom

from .abf import AbfImage
from .schema import _elements, _child, _text


def empty_table_storage(image, table_id):
    file = next(f for f in image.files if f.logical_path and
                f.logical_path.split('\\')[-1].startswith(table_id + '.') and
                f.logical_path.endswith('.tbl.xml'))
    doc = minidom.parseString(image.read_logical_file(file))
    native = minidom.parse(str(Path(__file__).parent / 'data' / 'empty-table.xml'))
    prototypes = {n.getAttribute('name'): n for n in _elements(native, 'XMObject')
                  if n.getAttribute('class') == 'XMRawColumn'}
    changes = {}
    for raw in list(_elements(doc, 'XMObject')):
        if raw.getAttribute('class') != 'XMRawColumn':
            continue
        name = raw.getAttribute('name')
        hidden = name in ('RowNumber', '__XL_RowNumber')
        dictionary = next(n for n in _elements(raw, 'XMObject')
                          if 'DataDictionary<' in n.getAttribute('class'))
        is_text = 'XM_String' in dictionary.getAttribute('class')
        clone = doc.importNode(prototypes['__XL_RowNumber' if hidden else 'Label' if is_text else 'ID'], True)
        clone.setAttribute('name', name)
        # Keep semantic flags and data type; reset only the physical empty state.
        clone.replaceChild(doc.importNode(_child(raw, 'Properties'), True), _child(clone, 'Properties'))
        for tag in ('DBType', 'XMType', 'Usage'):
            old, new = _elements(raw, tag), _elements(clone, tag)
            if old and new:
                new[0].firstChild.data = _text(old[0])
        old_data = next(n for n in _elements(raw, 'XMObject')
                        if n.getAttribute('class') == 'XMRawColumnPartitionDataObject')
        filename = old_data.getAttribute('name')
        for node in _elements(clone, 'XMObject'):
            kind = node.getAttribute('class')
            if kind == 'XMRawColumnPartitionDataObject':
                node.setAttribute('name', filename)
            elif kind == 'XMHierarchy':
                node.setAttribute('name', '[Hierarchy for column ' + name + ']')
            elif 'DataDictionary<' in kind and not hidden:
                node.setAttribute('class', dictionary.getAttribute('class').replace('XMValueDataDictionary', 'XMHashDataDictionary'))
        for leaf in _elements(clone, 'DataVersion'):
            leaf.firstChild.data = filename.split('.')[0]
        stream = next(f for f in image.files if f.logical_path and f.logical_path.endswith('\\' + filename))
        # Allocated native blocks: 16 RLE entries, then one zero packed word.
        changes[stream.storage_path] = (struct.pack('<QII', 16, 0xffffffff, 0) + bytes(128)
                                        if hidden else struct.pack('<Q', 16) + bytes(128) + struct.pack('<QQ', 1, 0))
        raw.parentNode.replaceChild(clone, raw)
    for tag in ('Records', 'FirstPartitionRecordCount', 'FirstPartitionSegmentCount', 'SegmentSize'):
        for leaf in _elements(doc, tag):
            leaf.firstChild.data = '0'
    changes[file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return AbfImage(image.assemble(changes))


def populate_empty_table(model, table, frame, *, type_overrides=None):
    """Rebuild missing physical indexes while retaining semantic model objects."""
    from io import BytesIO
    from zipfile import ZipFile
    import uuid
    import pandas as pd
    from .abf import VirtualFileAddition, UnsupportedModelError
    from .schema import clone_table

    from .calculated_columns import list_calculated_columns, resize_calculated_columns
    calculated = [c for c in list_calculated_columns(model._image) if c.table_id == model._dimension_ids[table]]
    calculated_names = {c.name for c in calculated}
    calculated_ids = {c.id for c in calculated}
    info = [c for c in model._columns if c.table == table and not c.name.startswith('__XL_')
            and c.name not in calculated_names]
    typed = frame.copy()
    type_overrides = type_overrides or {}
    with model._open_reader() as reader:
        schema = storage_schema(reader)
        currencies = {r['ColumnName'] for _, r in schema.iterrows()
                      if r['TableName'] == table and r.get('SemanticType') == 'Currency'}
    for column in info:
        values = typed[column.name]
        if column.name in type_overrides:
            continue  # The type-conversion caller already supplied exact dtypes.
        if column.name in currencies:
            from decimal import Decimal
            typed[column.name] = [None if pd.isna(v) else Decimal(str(v)) for v in values]
            continue
        if column.data_type == 'string':
            if any(not isinstance(v, str) for v in values.dropna()):
                raise ValueError('Text columns require string values')
            typed[column.name] = values.astype('string')
        else:
            dtype = {'int64': 'Int64', 'float64': 'Float64', 'bool': 'boolean',
                     'datetime64[ns]': 'datetime64[ns]'}.get(column.data_type)
            if dtype is None:
                raise UnsupportedModelError('Unsupported native empty column type')
            converted = values.astype(dtype)
            if any(a != b for a, b in zip(values, converted) if not pd.isna(a)):
                raise ValueError('Empty-table population would change input values')
            typed[column.name] = converted
    seed = typed.copy()
    for name, kind in type_overrides.items():
        if kind == 'currency' and len(seed) and seed[name].isna().all():
            from decimal import Decimal
            seed.loc[seed.index[0], name] = Decimal(0)
    donor = type(model).create({'EmptySource': seed})
    if not seed.equals(typed):
        donor.replace_table('EmptySource', typed)
    # Native captions need not equal the stable physical column identifiers.
    # Rename donor columns in two phases so swapped captions cannot collide.
    dimension = minidom.parseString(model._image.read_logical_file(model._dimension_paths[table]))
    renamed = []
    for attribute in _elements(dimension, 'Attribute'):
        ids = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'ID']
        names = [n for n in attribute.childNodes if getattr(n, 'localName', None) == 'Name']
        if ids and names and _text(names[0]) in typed and _text(ids[0]) != _text(names[0]):
            temporary = '_StorageColumn_' + uuid.uuid4().hex
            donor.rename_column('EmptySource', _text(names[0]), temporary)
            renamed.append((temporary, _text(ids[0])))
    for temporary, physical in renamed:
        donor.rename_column('EmptySource', temporary, physical)
    donor_id, table_id = donor._dimension_ids['EmptySource'], model._dimension_ids[table]
    with ZipFile(BytesIO(donor.to_bytes())) as archive:
        source = AbfImage(archive.read('xl/model/item.data'))
    source = clone_table(source, donor_id, table_id)
    def table_file(image):
        return next(f for f in image.files if f.logical_path and
                    f.logical_path.split('\\')[-1].startswith(table_id + '.') and
                    f.logical_path.endswith('.tbl.xml'))
    target_file, source_file = table_file(model._image), table_file(source)
    doc = minidom.parseString(model._image.read_logical_file(target_file))
    native = minidom.parseString(source.read_logical_file(source_file))
    source_raw = {n.getAttribute('name'): n for n in _elements(native, 'XMObject') if n.getAttribute('class') == 'XMRawColumn'}
    hidden_name = next(n.getAttribute('name') for n in _elements(doc, 'XMObject')
                       if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') in ('RowNumber', '__XL_RowNumber'))
    stored_ids = {n.getAttribute('name') for n in _elements(doc, 'XMObject')
                  if n.getAttribute('class') == 'XMRawColumn' and n.getAttribute('name') not in calculated_ids}
    for raw in list(_elements(doc, 'XMObject')):
        if raw.getAttribute('class') != 'XMRawColumn':
            continue
        name = raw.getAttribute('name')
        if name in calculated_ids:
            continue
        clone = doc.importNode(source_raw['RowNumber' if name == hidden_name else name], True)
        clone.setAttribute('name', name)
        if name == hidden_name:
            for node in _elements(clone, 'XMObject'):
                if node.getAttribute('class') == 'XMRawColumnPartitionDataObject':
                    node.setAttribute('name', node.getAttribute('name').replace('.RowNumber.', '.' + hidden_name + '.'))
        clone.replaceChild(doc.importNode(_child(raw, 'Properties'), True), _child(clone, 'Properties'))
        raw.parentNode.replaceChild(clone, raw)
    for tag in ('Records', 'FirstPartitionRecordCount'):
        for leaf in _elements(doc, tag):
            leaf.firstChild.data = str(len(frame))
    base = target_file.logical_path.rsplit('\\', 1)[0] + '\\'
    source_base = source_file.logical_path.rsplit('\\', 1)[0] + '\\'
    existing = {f.logical_path: f for f in model._image.files}
    changes, additions, filenames = {}, [], []
    for file in source.files:
        if not file.logical_path or not file.logical_path.startswith(source_base) or file == source_file:
            continue
        filename = file.logical_path[len(source_base):]
        if not (filename.endswith(('.idf', '.dictionary', '.hidx', '.tbl.xml'))):
            continue
        filename = filename.replace('.RowNumber.', '.' + hidden_name + '.')
        path = base + filename
        data = source.read_logical_file(file)
        filenames.append(filename)
        if path in existing:
            changes[existing[path].storage_path] = data
        else:
            additions.append(VirtualFileAddition(uuid.uuid4().hex.upper()[:20], path,
                target_file.storage_path, data,
                source.read_file(file)[-4:] if source.error_code and model._image.error_code else None))
    changes[target_file.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    # A codec change can make old dictionary/hash streams obsolete. Leaving
    # them listed makes independent readers mistake value storage for a hash
    # dictionary and can retain stale native indexes.
    removals, removed_names = [], set()
    for file in model._image.files:
        if not file.logical_path or not file.logical_path.startswith(base):
            continue
        filename = file.logical_path[len(base):]
        if filename in filenames or file == target_file:
            continue
        if any(f'.{table_id}.{column}.' in filename or f'H${table_id}${column}.' in filename
               for column in stored_ids):
            removals.append(file.storage_path)
            removed_names.add(filename)
    dim_file = model._dimension_paths[table]
    dim = minidom.parseString(model._image.read_logical_file(dim_file))
    listing = _elements(dim, 'DataFileList')[0]
    old = [name for name in _text(listing).split(';') if name not in removed_names]
    content = ';'.join(dict.fromkeys(old + filenames))
    while listing.firstChild:
        listing.removeChild(listing.firstChild)
    listing.appendChild(dim.createTextNode(content))
    changes[dim_file] = dim.documentElement.toxml(encoding='utf-8')
    result = AbfImage(model._image.assemble(changes, additions=additions, removals=removals))
    return resize_calculated_columns(result, table_id, len(frame)) if calculated else result
