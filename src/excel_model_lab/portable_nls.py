"""Portable assembly of pinned default-English Windows NLS sort weights.

The data is generated from Windows API outputs, not from ICU or Unicode casefold.
See data/nls-en-v1.json for its exact scope, version and generation evidence.
"""
from array import array
from bisect import bisect_left
from functools import lru_cache
import gzip
import json
from importlib.resources import files
import struct
import sys

from .abf import UnsupportedModelError


@lru_cache(maxsize=1)
def _weights():
    raw = gzip.decompress(files('excel_model_lab').joinpath('data/nls-en-v1.bin.gz').read_bytes())
    if raw[:8] != b'EMLNLS02':
        raise UnsupportedModelError('Invalid portable collation weight resource')
    count = struct.unpack_from('<I', raw, 8)[0]
    points = array('I')
    points.frombytes(raw[12:12 + count * 4])
    offsets = array('I')
    offsets.frombytes(raw[12 + count * 4:12 + count * 8])
    if sys.byteorder != 'little':
        points.byteswap()
        offsets.byteswap()
    return points, offsets, raw[12 + count * 8:]


@lru_cache(maxsize=4096)
def _character_weights(character):
    points, offsets, raw = _weights()
    code = ord(character)
    index = bisect_left(points, code)
    if index == len(points) or points[index] != code:
        raise UnsupportedModelError(
            f'Portable en-US/en-GB collation does not support U+{code:04X}; '
            'use Windows NLS for this contextual or unassigned character')
    offset = offsets[index]
    sizes = raw[offset:offset + 6]
    offset += 6
    result = []
    for size in sizes:
        result.append(raw[offset:offset + size])
        offset += size
    return result


# Native contextual repeat/prolongation weights under the pinned English sort.
_REPEATS = {0x3005: 6, 0x3031: 2, 0x3032: 5, 0x3033: 2, 0x3034: 5,
            0x3035: 2, 0x303b: 6, 0x309d: 2, 0x309e: 5, 0x30a0: 0,
            0x30fc: 2, 0x30fd: 2, 0x30fe: 5, 0xa015: 7, 0xff70: 2}


@lru_cache(maxsize=1)
def _jamo_states():
    return {int(k): v for k, v in json.loads(files('excel_model_lab').joinpath(
        'data/nls-jamo-v1.json').read_text()).items()}


def _jamo_class(code):
    if code <= 0x115f or 0xa960 <= code <= 0xa97f:
        return 0
    return 1 if code <= 0x11a7 or 0xd7b0 <= code <= 0xd7c6 else 2


def _jamo_cluster(value, index):
    states = _jamo_states()
    segment = [ord(value[index])]
    end = index + 1
    while end < len(value) and ord(value[end]) in states and _jamo_class(ord(value[end])) > _jamo_class(segment[-1]):
        segment.append(ord(value[end]))
        end += 1
    classes = [_jamo_class(c) for c in segment]
    if classes[0] == 0 and 1 in classes and not any(states[c][0] for c in segment):
        segment, classes, end = segment[:1], classes[:1], index + 1
    filler = 0 not in classes or 1 not in classes or 0x115f in segment or 0x1160 in segment
    include = segment + ([] if 0 in classes else [0x115f]) + ([] if 1 in classes else [0x1160])
    indices = [max(states[c][1][j] for c in include) for j in range(3)]
    weights = [max([states[c][2] for c in include if _jamo_class(c) == j] or [2]) for j in range(3)]
    modern = 0xac00 + (indices[0] * 21 + indices[1]) * 28 + indices[2] - filler
    primary = b'\xc0\x02' if modern < 0xac00 else _character_weights(chr(modern))[0]
    if filler:
        weights[0] += 128
    primary += bytes([255, weights[0], 255, weights[1], 255, weights[2]])
    return primary, bytes.fromhex(states[segment[-1]][3]), end


def portable_text_key(value, locale='en-US'):
    """Return an NLS-compatible key for characters in the pinned weight set."""
    if locale not in {'en-US', 'en-GB'}:
        raise UnsupportedModelError('Portable collation supports en-US and en-GB only')
    if not isinstance(value, str):
        raise TypeError('Collation values must be strings')
    if '\x00' in value:
        raise UnsupportedModelError('Embedded NUL characters are not supported in model strings')
    primary, diacritics, case, special = (bytearray() for _ in range(4))
    extras = [bytearray() for _ in range(4)]
    repeat_context = (b'\xff\xff', b'')
    index = 0
    while index < len(value):
        character = value[index]
        code = ord(character)
        index += 1
        if code in _jamo_states():
            pw, repeat, index = _jamo_cluster(value, index - 1)
            dw = cw = b'\x02'
            sw = extra = b''
            repeat_context = (repeat, b'')
        elif code in _REPEATS:
            pw, previous_extra = repeat_context
            extra = b''
            if pw == b'\xff\xff':
                dw = cw = b''
            else:
                dw, cw = bytes([_REPEATS[code]]), b'\x02'
                if previous_extra:
                    extra = bytearray(previous_extra)
                    extra[1] = 4
                    if code in {0x30fc, 0xff70}:
                        pw = pw[:-1] + bytes([pw[-1] & 0x87])
                        extra[1] = 5
                        extra[3] = 0xc4 if code == 0xff70 else 0xc5
            sw = b''
        else:
            pw, dw, cw, sw, extra, repeat = _character_weights(character)
            if pw or sw or repeat != b'\xff\xff':
                repeat_context = (repeat, extra)
        if not pw and dw and not sw:
            # A nonspacing accent augments the preceding diacritic weight.
            if diacritics:
                diacritics[-1] = (diacritics[-1] + sum(dw)) & 0xff
            else:
                diacritics.extend(dw)
        else:
            diacritics.extend(dw)
        case.extend(cw)
        for offset in range(0, len(sw), 4):
            position = (int.from_bytes(sw[offset:offset + 2], 'big') - len(primary) // 2) & 0xffff
            special.extend(position.to_bytes(2, 'big'))
            special.extend(sw[offset + 2:offset + 4])
        primary.extend(pw)
        for group, weight in zip(extras, extra):
            group.append(weight)
    while diacritics and diacritics[-1] <= 2:
        diacritics.pop()
    while case and case[-1] <= 2:
        case.pop()
    extra_key = b''
    if extras[0]:
        extra_key = (extras[0].rstrip(b'\xc6') + b'\xff' +
                     extras[1].rstrip(b'\x03') + b'\x02' +
                     extras[2].rstrip(b'\xe4') + b'\xff' +
                     extras[3].rstrip(b'\xc5') + b'\xff')
    return bytes(primary + b'\x01' + diacritics + b'\x01' + case +
                 b'\x01' + extra_key + b'\x01' + special + b'\x00')
