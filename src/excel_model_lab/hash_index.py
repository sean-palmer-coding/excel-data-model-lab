from __future__ import annotations

import struct
from collections import Counter
from io import BytesIO

from .abf import InvalidModelError, UnsupportedModelError

# Recovered from Excel-produced HIDX fixtures. Bucket selection uses the high
# bits of uint32(data_id * multiplier).
HASH_MULTIPLIER = 314_160_805


def encode_numeric_hash_index(template: bytes, data_ids: list[int], *, preserve_order=False) -> bytes:
    """Regenerate a value-encoded numeric HIDX from its distinct data IDs."""
    from pbixray.column_data.hidx import ColumnDataHidx

    try:
        parsed = ColumnDataHidx.from_io(BytesIO(template))
    except Exception as exc:
        raise InvalidModelError(f"Cannot parse template HIDX: {exc}") from exc
    bins = int(parsed.c_bins)
    if bins <= 0 or bins & (bins - 1):
        raise UnsupportedModelError("HIDX bin count must be a power of two")
    if parsed.hash_entry_size != 8 or parsed.hash_bin_size != 16 + parsed.local_entry_count * 8:
        raise UnsupportedModelError("Unsupported HIDX entry or bin layout")

    distinct = list(dict.fromkeys(int(value) for value in data_ids)) if preserve_order else sorted(set(int(value) for value in data_ids))
    local_count = int(parsed.local_entry_count)
    # Native files store zero chain pointers; overflow entries follow bucket
    # order and are reconstructed from each bucket's count when loaded.
    while len(distinct) > bins * 3:
        bins *= 2
    shift = 32 - (bins.bit_length() - 1)
    buckets: list[list[tuple[int, int]]] = [[] for _ in range(bins)]
    for key, value in enumerate(distinct):
        bucket = ((value * HASH_MULTIPLIER) & 0xFFFFFFFF) >> shift
        buckets[bucket].append((value, key))
    fast = sum(min(local_count, len(bucket)) for bucket in buckets)
    overflow = [entry for bucket in buckets for entry in bucket[local_count:]]
    histogram_size = (
        int(parsed.hash_statistics.element_count) if parsed.hash_stats else 10
    )
    occupancy = Counter(len(bucket) for bucket in buckets if bucket)
    histogram = [occupancy.get(index, 0) for index in range(1, histogram_size + 1)]

    output = bytearray()
    output += struct.pack(
        "<iIIIQqqB",
        int(parsed.hash_algorithm),
        int(parsed.hash_entry_size),
        int(parsed.hash_bin_size),
        local_count,
        bins,
        len(distinct),
        bins - 1,
        int(parsed.hash_stats),
    )
    if parsed.hash_stats:
        output += struct.pack(
            "<7QI",
            len(distinct),
            bins,
            sum(bool(bucket) for bucket in buckets),
            fast,
            local_count,
            max((len(bucket) for bucket in buckets), default=0),
            histogram_size,
            int(parsed.hash_statistics.element_size),
        )
        element_size = int(parsed.hash_statistics.element_size)
        for value in histogram:
            output += struct.pack("<I" if element_size == 4 else "<Q", value)
    for bucket in buckets:
        output += struct.pack("<QI", 0, len(bucket))
        local = bucket[:local_count]
        local += [(0, 0)] * (local_count - len(local))
        output += b"".join(struct.pack("<II", value, key) for value, key in local)
        output += struct.pack("<I", 0)
    output += struct.pack("<Q", len(overflow))
    output += b"".join(struct.pack("<II", value, key) for value, key in overflow)
    return bytes(output)
