"""Narrow table rewrite retaining the template's lookup structures.

Excel acceptance has been checked for the 500-to-40-row small fixture with an
appended string and the supplier fixture's two resized dimension tables. This
does not establish support for arbitrary models or relationship recalculation.
"""
from .reader_adapter import storage_schema, segment_metadata, decode_dictionary
from xml.etree import ElementTree as ET
import struct

import pandas as pd
from pbixray import PBIXRay

from .metadata import IMBI, XSI, _member_object, _set, update_string_column_metadata
from .dictionary import encode_uncompressed_string_dictionary
from .vertipaq import encode_pure_bitpacked_ids


def preserve_table(model, table):
    """Return replacements for an eligible subset rewrite, otherwise None."""
    from .model import _ValueEncoding, _StringEncoding, _NumericDictionaryEncoding

    columns = {key[1]: value for key, value in model._encodings.items() if key[0] == table}
    # Retaining unused dictionary members produces phantom DISTINCTCOUNT values.
    # The ordinary save path compacts dictionaries and rebuilds dependent indexes.
    if any(isinstance(e, (_StringEncoding, _NumericDictionaryEncoding)) for e in columns.values()):
        return None
    count = model._resized_tables[table]
    if not columns or count < 1 or any(
        not isinstance(e, (_ValueEncoding, _StringEncoding, _NumericDictionaryEncoding))
        or len(e.segments) != 1 or count > e.row_count for e in columns.values()
    ):
        return None
    metadata_path = next(iter(columns.values())).metadata_path
    root = ET.fromstring(model._image.read_logical_file(metadata_path))
    maps = [x for x in root.iter(f'{{{IMBI}}}XMObject') if x.get('class') == 'XMSegment1Map']
    if len(maps) != 1:
        return None
    changes = {}
    string_updates = []
    with model._open_reader() as reader:
        original = reader.get_table(table)
        for column, encoding in columns.items():
            physical_column = model._column_ids[table, column]
            values = model._pending[(table, column)]
            if isinstance(encoding, _StringEncoding):
                dictionary = decode_dictionary(reader, 
                    model._image.read_logical_file(encoding.dictionary_path), 3
                )
                inverse = {v: k for k, v in dictionary.values.items()}
                if any(not pd.isna(v) and str(v) not in inverse for v in values):
                    ids, entries = encoding.encode(values, allow_resize=True)
                    changes[encoding.dictionary_path] = encode_uncompressed_string_dictionary(
                        model._image.read_logical_file(encoding.dictionary_path), entries
                    )
                    string_updates.append((physical_column, ids, len(entries)))
                else:
                    ids = [2 if pd.isna(v) else inverse[str(v)] for v in values]
            elif isinstance(encoding, _NumericDictionaryEncoding):
                ids, entries = encoding.encode(values, allow_resize=True)
                if len(entries) != len(encoding.original_values):
                    return None
            else:
                if encoding.semantic_type == 'Boolean':
                    return None
                ids = encoding.encode(values, allow_resize=True)
                if any(not pd.isna(v) and i < 3 for v, i in zip(values, ids)):
                    return None
                existing = set(encoding.encode(original[column].tolist()))
                if not set(ids).issubset(existing):
                    return None
            segment = encoding.segments[0]
            width = int(segment['bit_width'])
            minimum = 2 if segment.get('has_nulls') else int(segment['min_data_id'])
            if width > 32 or any(not 0 <= i - minimum < 1 << width for i in ids):
                return None
            changes[encoding.storage_path] = encode_pure_bitpacked_ids(
                ids, bit_width=width, min_data_id=minimum
            )
            target = next(x for x in root.iter(f'{{{IMBI}}}XMObject')
                          if x.get('class') == 'XMRawColumn' and x.get('name') == physical_column)
            stats = _member_object(target, 'ColumnStats')
            _set(stats.find(f'{{{IMBI}}}Properties'), 'RowCount', count)
            for obj in target.iter(f'{{{IMBI}}}XMObject'):
                props = obj.find(f'{{{IMBI}}}Properties')
                if obj.get('class') == 'XMColumnSegment':
                    _set(props, 'Records', count)
                elif obj.get('class') == 'XMColumnSegmentStats':
                    if props.findtext(f'{{{IMBI}}}RowCount') != '0':
                        _set(props, 'RowCount', count)
                        _set(props, 'RLERuns', 0)
                        _set(props, 'OthersRLERuns', 1)
                elif obj.get('class') == 'XMRLECompressionInfo':
                    _set(props, 'StorageUsedSize', 4)
    _set(maps[0].find(f'{{{IMBI}}}Properties'), 'Records', count)
    # The engine's sequential row identifier also needs the new record count.
    hidden = [x for x in root.iter(f'{{{IMBI}}}XMObject')
              if x.get('class') == 'XMRawColumn' and x.get('name') in ('__XL_RowNumber', 'RowNumber')
              and x.get('name') not in columns]
    if len(hidden) != 1:
        return None
    original_count = next(iter(columns.values())).row_count
    for x in hidden[0].iter():
        name = x.tag.split('}')[-1]
        if name in ('RowCount', 'Records', 'DistinctDataIDs', 'DistinctStates') and x.text == str(original_count):
            x.text = str(count)
        elif name == 'MaxDataID' and x.text == str(original_count + 2):
            x.text = str(count + 2)
    suffix = f"{model._dimension_ids[table]}.{hidden[0].get('name')}.0.idf"
    files = [f for f in model._image.files if f.logical_path and
             f.logical_path.replace('\\', '/').split('/')[-1].endswith(suffix)]
    if len(files) != 1:
        return None
    data = bytearray(model._image.read_logical_file(files[0]))
    if len(data) < 16 or struct.unpack_from('<I', data, 8)[0] != 0xFFFFFFFF:
        return None
    struct.pack_into('<I', data, 12, count)
    changes[files[0].storage_path] = bytes(data)
    ET.register_namespace('', IMBI)
    ET.register_namespace('xsi', XSI)
    changes[metadata_path] = ET.tostring(root, encoding='utf-8').replace(
        b'>', b' xmlns:xsd="http://www.w3.org/2001/XMLSchema">', 1
    )
    for column, ids, cardinality in string_updates:
        changes[metadata_path] = update_string_column_metadata(
            changes[metadata_path], column=column, data_ids=ids, distinct_count=cardinality
        )
    return changes
