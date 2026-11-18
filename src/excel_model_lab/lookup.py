"""Numeric intrinsic-hierarchy streams associated with value-encoded columns."""
import struct
from xml.dom import minidom

from .abf import UnsupportedModelError
from .schema import _elements, _child, _text


def numeric_lookup_replacements(image, table_id, column, data_ids, *, preserve_order=False):
    """Rebuild the numeric position-to-ID mapping using sorted distinct IDs.

    HIDX keys are positions in this mapping, so both streams must be replaced
    together. Excel's 32-bit mapping dialect includes three sentinel records.
    """
    stem = f'H${table_id}${column}'
    matches = [f for f in image.files if f.logical_path
               and f.logical_path.split('\\')[-1].startswith(stem + '.')
               and f.logical_path.endswith('.tbl.xml')]
    if len(matches) != 1:
        raise UnsupportedModelError('Cannot resolve numeric intrinsic hierarchy')
    meta = matches[0]
    doc = minidom.parseString(image.read_logical_file(meta))
    maps = [n for n in _elements(doc, 'XMObject')
            if n.getAttribute('class').startswith('XMSegmentEqualMapEx<')]
    columns = [n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn']
    if len(maps) != 1 or len(columns) != 1 or columns[0].getAttribute('name') != 'POS_TO_ID':
        raise UnsupportedModelError('Unsupported numeric hierarchy layout')
    segments = [n for n in _elements(columns[0], 'XMObject')
                if n.getAttribute('class') == 'XMColumnSegment']
    codecs = [n for n in _elements(columns[0], 'XMObject')
              if n.getAttribute('class').startswith('XMRENoSplitCompressionInfo')]
    if not segments or len(codecs) != len(segments) or any(
        n.getAttribute('class') != 'XMRENoSplitCompressionInfo<32>' for n in codecs
    ):
        raise UnsupportedModelError('Unsupported numeric mapping encoding')
    ids = list(dict.fromkeys(data_ids)) if preserve_order else sorted(set(data_ids))
    if not ids or min(ids) < 0 or max(ids) > 0xffffffff:
        raise UnsupportedModelError('Numeric mapping requires nonempty unsigned 32-bit IDs')
    def set_text(node, local, value):
        leaf = _child(_child(node, 'Properties'), local)
        while leaf.firstChild:
            leaf.removeChild(leaf.firstChild)
        leaf.appendChild(doc.createTextNode(str(value)))
    set_text(maps[0], 'Records', len(ids) + 3)
    set_text(maps[0], 'RecordsPerSegment', len(ids))
    # Equal-map segments may never exceed RecordsPerSegment. For one or
    # two distinct values, the three sentinel records need additional segments.
    values = ids + [0, 0, 0]
    chunks = [values[start:start + len(ids)] for start in range(0, len(values), len(ids))]
    parent = segments[0].parentNode
    prototype = segments[0].cloneNode(True)
    for segment in segments:
        parent.removeChild(segment)
    for chunk in chunks:
        segment = prototype.cloneNode(True)
        set_text(segment, 'Records', len(chunk))
        parent.appendChild(segment)
    set_text(maps[0], 'Segments', len(chunks))
    objects = [n for n in _elements(columns[0], 'XMObject')
               if n.getAttribute('class') == 'XMRawColumnPartitionDataObject']
    if len(objects) != 1:
        raise UnsupportedModelError('Cannot resolve numeric mapping data file')
    set_text(objects[0], 'SegmentCount', len(chunks))
    filename = objects[0].getAttribute('name')
    files = [f for f in image.files if f.logical_path
             and f.logical_path.rsplit('\\', 1)[0] == meta.logical_path.rsplit('\\', 1)[0]
             and f.logical_path.split('\\')[-1] == filename]
    if len(files) != 1:
        raise UnsupportedModelError('Cannot resolve numeric mapping data file')
    stream = b''
    for chunk in chunks:
        words = len(chunk) // 2 + 1
        payload = struct.pack('<' + 'I' * len(chunk), *chunk).ljust(words * 8, b'\0')
        stream += struct.pack('<Q', words) + payload
    return {meta.storage_path: doc.documentElement.toxml(encoding='utf-8'),
            files[0].storage_path: stream}


def string_lookup_replacements(image, table_id, column, dictionary, *, has_nulls=False, ordered=None):
    """Rebuild intrinsic string lookups with validated model collation."""
    from .collation import model_string_order
    natural_order = model_string_order(image, table_id, dictionary)
    ordered = natural_order if ordered is None else list(ordered)
    return dictionary_lookup_replacements(image, table_id, column, dictionary,
                                         has_nulls=has_nulls, ordered=ordered)


def dictionary_lookup_replacements(image, table_id, column, dictionary, *, has_nulls=False, ordered):
    """Rebuild the two lookup directions shared by string and numeric dictionaries."""
    stem = f'H${table_id}${column}'
    matches = [f for f in image.files if f.logical_path and
               f.logical_path.split('\\')[-1].startswith(stem + '.') and f.logical_path.endswith('.tbl.xml')]
    if len(matches) != 1:
        raise UnsupportedModelError('Cannot resolve string intrinsic hierarchy')
    meta = matches[0]
    doc = minidom.parseString(image.read_logical_file(meta))
    objects = _elements(doc, 'XMObject')
    maps = [n for n in objects if n.getAttribute('class').startswith('XMSegmentEqualMapEx<')]
    columns = [n for n in objects if n.getAttribute('class') == 'XMRawColumn']
    if len(maps) != 1 or {n.getAttribute('name') for n in columns} != {'POS_TO_ID', 'ID_TO_POS'}:
        raise UnsupportedModelError('Unsupported string lookup layout')
    count = len(dictionary)
    ordered = list(ordered)
    if sorted(ordered) != list(range(count)):
        raise ValueError('String sort order must include each dictionary entry exactly once')
    position = {index: rank + int(has_nulls) for rank, index in enumerate(ordered)}
    forward = ([2] if has_nulls else []) + [index + 3 for index in ordered]
    forward += [0] * (count + 3 - len(forward))
    inverse = [0, count + int(has_nulls), 0] + [position[index] for index in range(count)]
    def set_value(node, key, value):
        leaf = _child(_child(node, 'Properties'), key)
        leaf.firstChild.data = str(value)
    set_value(maps[0], 'Records', count + 3)
    # The blank member occupies a lookup position even though it has no
    # dictionary payload. Omitting it here makes Excel's DISTINCTCOUNT skip it.
    records_per_segment = max(count + int(has_nulls), 1)
    set_value(maps[0], 'RecordsPerSegment', records_per_segment)
    set_value(maps[0], 'Segments', (count + 3 + records_per_segment - 1) // records_per_segment)
    replacements = {}
    for raw in columns:
        segments = [n for n in _elements(raw, 'XMObject') if n.getAttribute('class') == 'XMColumnSegment']
        codecs = [n for n in _elements(raw, 'XMObject') if n.getAttribute('class').startswith('XMRENoSplitCompressionInfo')]
        if not segments or len(codecs) != len(segments) or any(n.getAttribute('class') != 'XMRENoSplitCompressionInfo<32>' for n in codecs):
            raise UnsupportedModelError('Unsupported string lookup encoding')
        values = forward if raw.getAttribute('name') == 'POS_TO_ID' else inverse
        chunks = [values[start:start + records_per_segment]
                  for start in range(0, len(values), records_per_segment)]
        parent = segments[0].parentNode
        prototype = segments[0].cloneNode(True)
        for segment in segments:
            parent.removeChild(segment)
        stream = b''
        for entries in chunks:
            segment = prototype.cloneNode(True)
            set_value(segment, 'Records', len(entries))
            parent.appendChild(segment)
            words = len(entries) // 2 + 1
            stream += struct.pack('<Q', words) + struct.pack('<' + 'I'*len(entries), *entries).ljust(words*8, b'\0')
        data = [n for n in _elements(raw, 'XMObject') if n.getAttribute('class') == 'XMRawColumnPartitionDataObject']
        if len(data) != 1:
            raise UnsupportedModelError('Cannot resolve string lookup stream')
        set_value(data[0], 'SegmentCount', len(chunks))
        path = meta.logical_path.rsplit('\\', 1)[0] + '\\' + data[0].getAttribute('name')
        files = [f for f in image.files if f.logical_path == path]
        if len(files) != 1:
            raise UnsupportedModelError('Cannot resolve string lookup stream')
        replacements[files[0].storage_path] = stream
    replacements[meta.storage_path] = doc.documentElement.toxml(encoding='utf-8')
    return replacements
