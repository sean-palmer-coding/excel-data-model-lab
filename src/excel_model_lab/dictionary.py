from __future__ import annotations

import struct
from collections.abc import Sequence
from io import BytesIO

from .abf import InvalidModelError, UnsupportedModelError


def write_uncompressed_string_dictionary(output, template: bytes,
                                         values: Sequence[str], *,
                                         page_size_bytes: int = 1 << 20) -> int:
    """Write bounded-size UTF-16 pages and stable dictionary record handles.

    MS-XLDM 2.3.2.1.3.2 permits at most 524,288 pages, each at most 4 GiB.
    Individual strings never cross a page boundary. A string larger than the
    target gets its own page, subject to the format's hard page-size limit.
    Scratch space is bounded by the largest string, page descriptors, and the
    references used for uniqueness checking; no complete text buffer is joined.
    """
    from pbixray.column_data.dictionary import ColumnDataDictionary
    import operator
    try:
        page_size_bytes = operator.index(page_size_bytes)
    except TypeError as exc:
        raise ValueError('Page size must be an integer') from exc
    if not 60 <= page_size_bytes <= 1 << 32:
        raise ValueError('Page size must be between 60 bytes and 4 GiB')
    try:
        parsed = ColumnDataDictionary.from_io(BytesIO(template))
    except Exception as exc:
        raise InvalidModelError(f"Cannot parse template string dictionary: {exc}") from exc
    if int(parsed.dictionary_type) != 2:
        raise ValueError("Template is not a string dictionary")
    seen, pages = set(), []
    longest = page_units = first = 0
    for index, value in enumerate(values):
        if not isinstance(value, str):
            raise TypeError('Dictionary values must be strings')
        if value == '':
            raise UnsupportedModelError(
                f'Empty stored text at dictionary position {index} is not supported by the Excel model writer; '
                'choose None explicitly for BLANK through the model API or supply nonempty text')
        if value in seen:
            raise ValueError('Dictionary values must be unique')
        seen.add(value)
        if '\x00' in value:
            raise ValueError('Dictionary strings cannot contain NUL characters')
        units = len(value.encode('utf-16-le')) // 2
        longest = max(longest, units)
        if (units + 1) * 2 + 58 > 1 << 32:
            raise ValueError('One string exceeds the 4 GiB dictionary page limit')
        if page_units and (page_units + units + 1) * 2 + 58 > page_size_bytes:
            pages.append((first, index, page_units))
            first, page_units = index, 0
        page_units += units + 1
    pages.append((first, len(values), page_units))
    if len(pages) > 524288:
        raise ValueError('String dictionary exceeds the 524288-page limit')
    del seen
    count = len(values)
    header = (struct.pack('<i', 2) +
              struct.pack('<6i', *parsed.hash_information.hash_elements) +
              # Retain the historically accepted single-page header byte.
              struct.pack('<qbqq', count, int(len(pages) == 1), longest, len(pages)))
    output.write(header)
    written = len(header)
    for first, end, units in pages:
        output.write(struct.pack('<QBQQB', 0, 0, first, end - first, 0) +
                     b'\xdd\xcc\xbb\xaa' + struct.pack('<QQQ', 0, units, units * 2))
        for index in range(first, end):
            output.write(values[index].encode('utf-16-le'))
            output.write(b'\0\0')
        output.write(b'\xcd\xab\xcd\xab')
        written += 58 + units * 2
    output.write(struct.pack('<QI', count, 8))
    handles = bytearray()
    for page_id, (first, end, _) in enumerate(pages):
        offset = 0
        for index in range(first, end):
            handles += struct.pack('<II', offset, page_id)
            offset += len(values[index].encode('utf-16-le')) // 2 + 1
            if len(handles) >= 65536:
                output.write(handles)
                handles.clear()
    output.write(handles)
    return written + 12 + count * 8


def encode_uncompressed_string_dictionary(template: bytes,
                                          values: Sequence[str], *,
                                          page_size_bytes: int = 1 << 20) -> bytes:
    """Build a paged string dictionary; BLANK is ID 2, values start at ID 3."""
    output = BytesIO()
    write_uncompressed_string_dictionary(output, template, values,
                                         page_size_bytes=page_size_bytes)
    return output.getvalue()


def encode_numeric_dictionary(template: bytes, values: Sequence[int | float]) -> bytes:
    """Build a VertiPaq long or real vector dictionary from unique values."""
    from pbixray.column_data.dictionary import ColumnDataDictionary

    try:
        parsed = ColumnDataDictionary.from_io(BytesIO(template))
    except Exception as exc:
        raise InvalidModelError(f"Cannot parse template numeric dictionary: {exc}") from exc
    dictionary_type = int(parsed.dictionary_type)
    if dictionary_type not in {0, 1}:
        raise ValueError("Template is not a numeric dictionary")
    element_size = int(parsed.data.vector_of_vectors_info.element_size)
    if dictionary_type == 0 and element_size not in {4, 8}:
        raise ValueError(f"Unsupported integer dictionary element size {element_size}")
    if dictionary_type == 1 and element_size != 8:
        raise ValueError(f"Unsupported real dictionary element size {element_size}")
    import math
    converted = []
    for value in values:
        try:
            numeric = float(value) if dictionary_type == 1 else int(value)
        except (OverflowError, ValueError, TypeError) as exc:
            raise ValueError('Numeric dictionary values must be finite numbers') from exc
        if dictionary_type == 1 and not math.isfinite(numeric):
            raise ValueError('Numeric dictionary values must be finite')
        if numeric != value:
            raise ValueError('Numeric dictionary conversion would lose precision')
        if dictionary_type == 0 and not -(1 << (element_size * 8 - 1)) <= numeric < (1 << (element_size * 8 - 1)):
            raise ValueError('Numeric dictionary value exceeds its signed storage range')
        converted.append(numeric)
    if len(set(converted)) != len(converted):
        raise ValueError('Numeric dictionary values must be unique')
    values = converted
    output = bytearray()
    output += struct.pack("<i", dictionary_type)
    output += struct.pack("<6i", *parsed.hash_information.hash_elements)
    output += struct.pack("<QI", len(values), element_size)
    if dictionary_type == 1:
        output += b"".join(struct.pack("<d", float(value)) for value in values)
    elif element_size == 4:
        output += b"".join(struct.pack("<i", int(value)) for value in values)
    else:
        output += b"".join(struct.pack("<q", int(value)) for value in values)
    return bytes(output)
