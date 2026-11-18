from __future__ import annotations

import struct
from collections.abc import Sequence

from .abf import InvalidModelError, UnsupportedModelError


# MS-XLDM XMRENoSplitCompressionInfo specializations, not arbitrary widths.
SUPPORTED_BIT_WIDTHS = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 21, 32)


def select_bit_width(span: int) -> int:
    if span < 0:
        raise ValueError("Encoded span must be nonnegative")
    required = max(1, span.bit_length())
    for width in SUPPORTED_BIT_WIDTHS:
        if width >= required:
            return width
    raise UnsupportedModelError("Encoded span exceeds 32-bit compression")


def patch_single_segment_bitpacked_id(
    idf: bytes,
    *,
    position: int,
    bit_width: int,
    new_data_id: int,
    min_data_id: int,
) -> bytes:
    """Patch one data ID in the narrow single-segment bit-packed IDF form.

    This is the first mutation primitive, not a general IDF encoder. It supports
    a single segment whose first primary entry points to the bit-packed stream
    and whose remaining primary entries are padding.
    """
    if not 1 <= bit_width <= 63:
        raise ValueError("bit_width must be between 1 and 63")
    if position < 0:
        raise ValueError("position must be nonnegative")
    if len(idf) < 17:
        raise InvalidModelError("IDF is too short")
    primary_count = struct.unpack_from("<Q", idf, 0)[0]
    primary_bytes = primary_count * 8
    sub_size_offset = 8 + primary_bytes
    if sub_size_offset + 8 > len(idf):
        raise InvalidModelError("IDF primary segment exceeds the file")
    first_value, first_repeat = struct.unpack_from("<II", idf, 8)
    if first_value != 0xFFFFFFFF:
        raise UnsupportedModelError("IDF is not the expected bit-packed-only form")
    if position >= first_repeat:
        raise IndexError(f"position {position} is outside {first_repeat} records")
    for offset in range(16, sub_size_offset, 8):
        value, repeat = struct.unpack_from("<II", idf, offset)
        if value or repeat:
            raise UnsupportedModelError("IDF contains additional RLE entries")
    sub_word_count = struct.unpack_from("<Q", idf, sub_size_offset)[0]
    sub_offset = sub_size_offset + 8
    if sub_offset + sub_word_count * 8 != len(idf):
        raise UnsupportedModelError("IDF contains multiple segments or trailing data")

    encoded = new_data_id - min_data_id
    limit = 1 << bit_width
    if not 0 <= encoded < limit:
        raise ValueError(
            f"data ID {new_data_id} cannot be represented with min {min_data_id} "
            f"and width {bit_width}"
        )
    values_per_word = 64 // bit_width
    word_index, slot = divmod(position, values_per_word)
    if word_index >= sub_word_count:
        raise InvalidModelError("Bit-packed stream is shorter than its record count")
    shift = slot * bit_width
    mask = ((1 << bit_width) - 1) << shift
    result = bytearray(idf)
    word = struct.unpack_from("<Q", result, sub_offset + word_index * 8)[0]
    word = (word & ~mask) | (encoded << shift)
    struct.pack_into("<Q", result, sub_offset + word_index * 8, word)
    return bytes(result)


def encode_single_segment_bitpacked_ids(
    template_idf: bytes,
    *,
    data_ids: Sequence[int],
    bit_width: int,
    min_data_id: int,
) -> bytes:
    """Regenerate an entire supported single-segment IDF from data IDs."""
    if not 1 <= bit_width <= 63:
        raise ValueError("bit_width must be between 1 and 63")
    if len(template_idf) < 17:
        raise InvalidModelError("IDF is too short")
    primary_count = struct.unpack_from("<Q", template_idf, 0)[0]
    if primary_count < 1:
        raise UnsupportedModelError("IDF has no primary entries")
    sub_size_offset = 8 + primary_count * 8
    if sub_size_offset + 8 > len(template_idf):
        raise InvalidModelError("IDF primary segment exceeds the file")
    first_value, original_records = struct.unpack_from("<II", template_idf, 8)
    if first_value != 0xFFFFFFFF:
        raise UnsupportedModelError("IDF is not the expected bit-packed-only form")
    if len(data_ids) != original_records:
        raise UnsupportedModelError(
            f"Row count must remain {original_records}; received {len(data_ids)}"
        )
    for offset in range(16, sub_size_offset, 8):
        value, repeat = struct.unpack_from("<II", template_idf, offset)
        if value or repeat:
            raise UnsupportedModelError("IDF contains additional RLE entries")

    limit = 1 << bit_width
    encoded = []
    for data_id in data_ids:
        code = int(data_id) - min_data_id
        if not 0 <= code < limit:
            raise ValueError(
                f"data ID {data_id} cannot be represented with min {min_data_id} "
                f"and width {bit_width}"
            )
        encoded.append(code)
    values_per_word = 64 // bit_width
    word_count = (len(encoded) + values_per_word - 1) // values_per_word
    words = []
    for base in range(0, len(encoded), values_per_word):
        word = 0
        for slot, code in enumerate(encoded[base:base + values_per_word]):
            word |= code << (slot * bit_width)
        words.append(word)
    prefix = bytearray(template_idf[:sub_size_offset])
    struct.pack_into("<II", prefix, 8, 0xFFFFFFFF, len(encoded))
    return bytes(prefix) + struct.pack("<Q", word_count) + b"".join(
        struct.pack("<Q", word) for word in words
    )


def write_pure_bitpacked_ids(
    output,
    data_ids,
    *,
    record_count: int,
    bit_width: int,
    min_data_id: int,
    primary_count: int = 16,
) -> int:
    """Write a one-segment IDF to a binary stream using bounded scratch memory.

    ``data_ids`` may be a one-shot iterator. The caller supplies its exact length
    so neither seeking nor retaining all values is necessary. On an invalid
    value/count the destination may contain a partial segment; callers should
    stage output in a temporary file before publishing it.
    """
    import operator
    record_count = operator.index(record_count)
    if not 0 <= record_count <= 0xFFFFFFFF:
        raise ValueError("record_count must fit an unsigned 32-bit segment")
    if not 1 <= bit_width <= 32:
        raise ValueError("bit_width must be between 1 and 32")
    if primary_count < 1:
        raise ValueError("primary_count must be positive")
    minimum = operator.index(min_data_id)
    if not 0 <= minimum <= 0xFFFFFFFF:
        raise ValueError("min_data_id must fit an unsigned 32-bit data ID")
    iterator = iter(data_ids)
    if record_count == 0:
        sentinel = object()
        if next(iterator, sentinel) is not sentinel:
            raise ValueError("Received more data IDs than record_count")
        output.write(struct.pack('<QQ', 0, 0))
        return 16
    per_word = 64 // bit_width
    word_count = record_count // per_word + 1
    output.write(struct.pack('<QII', primary_count, 0xFFFFFFFF, record_count))
    output.write(bytes((primary_count - 1) * 8))
    output.write(struct.pack('<Q', word_count))
    limit = 1 << bit_width
    word = 0
    slot = 0
    buffer = bytearray()
    for position in range(record_count):
        try:
            value = operator.index(next(iterator))
        except StopIteration as exc:
            raise ValueError(f"Expected {record_count} data IDs; received {position}") from exc
        code = value - minimum
        if not 0 <= value <= 0xFFFFFFFF or not 0 <= code < limit:
            raise ValueError(
                f"data ID {value} cannot be represented with min {minimum} "
                f"and width {bit_width}"
            )
        word |= code << (slot * bit_width)
        slot += 1
        if slot == per_word:
            buffer += struct.pack('<Q', word)
            word = slot = 0
            if len(buffer) >= 65536:
                output.write(buffer)
                buffer.clear()
    sentinel = object()
    if next(iterator, sentinel) is not sentinel:
        raise ValueError("Received more data IDs than record_count")
    # Native readers may inspect one word after an exact packed-word boundary.
    buffer += struct.pack('<Q', word)
    output.write(buffer)
    return 16 + primary_count * 8 + word_count * 8


def encode_pure_bitpacked_ids(
    data_ids: Sequence[int],
    *,
    bit_width: int,
    min_data_id: int,
    primary_count: int = 16,
) -> bytes:
    """Create a one-segment IDF without intermediate value or word lists."""
    from io import BytesIO
    output = BytesIO()
    write_pure_bitpacked_ids(output, data_ids, record_count=len(data_ids),
                            bit_width=bit_width, min_data_id=min_data_id,
                            primary_count=primary_count)
    return output.getvalue()


def encode_single_segment_preserving_layout(
    template_idf: bytes,
    *,
    data_ids: Sequence[int],
    bit_width: int,
    min_data_id: int,
) -> bytes:
    """Encode values while retaining a segment's existing RLE/bit-pack topology.

    Retaining the topology also retains the record counts declared in the table
    metadata. Values landing in an existing RLE run must remain uniform; callers
    receive a precise error when a rewrite would require changing that metadata.
    """
    if not 1 <= bit_width <= 63:
        raise ValueError("bit_width must be between 1 and 63")
    if len(template_idf) < 24:
        raise InvalidModelError("IDF is too short")
    primary_count = struct.unpack_from("<Q", template_idf, 0)[0]
    sub_size_offset = 8 + primary_count * 8
    if sub_size_offset + 8 > len(template_idf):
        raise InvalidModelError("IDF primary segment exceeds the file")

    primary: list[tuple[int, int]] = []
    bitpacked: list[int] = []
    row_offset = 0
    bitpacked_offset = 0
    for entry_index in range(primary_count):
        data_value, repeat = struct.unpack_from("<II", template_idf, 8 + entry_index * 8)
        if repeat == 0 or row_offset >= len(data_ids):
            primary.append((0, 0))
            continue
        repeat = min(repeat, len(data_ids) - row_offset)
        values = [int(value) for value in data_ids[row_offset:row_offset + repeat]]
        is_bitpacked = (data_value + bitpacked_offset) & 0xFFFFFFFF == 0xFFFFFFFF
        if is_bitpacked:
            primary.append(((0xFFFFFFFF - bitpacked_offset) & 0xFFFFFFFF, repeat))
            bitpacked.extend(values)
            bitpacked_offset += repeat
        else:
            if any(value != values[0] for value in values[1:]):
                raise UnsupportedModelError(
                    f"Replacement breaks the existing RLE run at rows "
                    f"{row_offset}:{row_offset + repeat}"
                )
            primary.append((values[0], repeat))
        row_offset += repeat
    if row_offset != len(data_ids):
        raise UnsupportedModelError(
            f"IDF layout covers {row_offset} rows but received {len(data_ids)}"
        )

    limit = 1 << bit_width
    codes: list[int] = []
    for data_id in bitpacked:
        code = data_id - min_data_id
        if not 0 <= code < limit:
            raise ValueError(
                f"data ID {data_id} cannot be represented with min {min_data_id} "
                f"and width {bit_width}"
            )
        codes.append(code)
    values_per_word = 64 // bit_width
    words = []
    for base in range(0, len(codes), values_per_word):
        word = 0
        for slot, code in enumerate(codes[base:base + values_per_word]):
            word |= code << (slot * bit_width)
        words.append(word)

    return (
        struct.pack("<Q", primary_count)
        + b"".join(struct.pack("<II", value, repeat) for value, repeat in primary)
        + struct.pack("<Q", len(words))
        + b"".join(struct.pack("<Q", word) for word in words)
    )


def encode_idf_preserving_layout(
    template_idf: bytes,
    *,
    data_ids: Sequence[int],
    segments: Sequence[dict],
) -> bytes:
    """Encode every segment in an IDF while retaining each segment topology."""
    output: list[bytes] = []
    byte_offset = 0
    row_offset = 0
    for index, metadata in enumerate(segments):
        if byte_offset + 8 > len(template_idf):
            raise InvalidModelError(f"IDF ends before segment {index}")
        primary_count = struct.unpack_from("<Q", template_idf, byte_offset)[0]
        sub_size_offset = byte_offset + 8 + primary_count * 8
        if sub_size_offset + 8 > len(template_idf):
            raise InvalidModelError(f"Segment {index} primary data exceeds the IDF")
        word_count = struct.unpack_from("<Q", template_idf, sub_size_offset)[0]
        segment_end = sub_size_offset + 8 + word_count * 8
        if segment_end > len(template_idf):
            raise InvalidModelError(f"Segment {index} bit-packed data exceeds the IDF")
        records = int(metadata.get("records", 0))
        segment_values = data_ids[row_offset:row_offset + records]
        output.append(
            encode_single_segment_preserving_layout(
                template_idf[byte_offset:segment_end],
                data_ids=segment_values,
                bit_width=int(metadata["bit_width"]),
                min_data_id=(2 if metadata.get("has_nulls") else int(metadata["min_data_id"])),
            )
        )
        row_offset += records
        byte_offset = segment_end
    if byte_offset != len(template_idf):
        raise InvalidModelError("IDF contains more segments than its metadata")
    if row_offset != len(data_ids):
        raise ValueError(f"Segment metadata covers {row_offset} rows; received {len(data_ids)}")
    return b"".join(output)
