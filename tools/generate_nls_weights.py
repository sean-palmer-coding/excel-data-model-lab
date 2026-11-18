"""Regenerate pinned portable sort weights on Windows; no Excel is required.

Run with the project installed. Every retained character is checked against
both en-US and en-GB LCMapStringEx, including prefix/suffix contexts.
"""
import ctypes
import gzip
import hashlib
import json
import platform
from pathlib import Path
import struct
import unicodedata

from excel_model_lab.collation import _native_text_key


def levels(value, locale='en-US'):
    return _native_text_key(value, locale)[:-1].rsplit(b'\x01', 4)


def assemble(value, weights):
    # Use the production assembler with an in-memory table during generation.
    import excel_model_lab.portable_nls as portable
    original = portable._character_weights
    try:
        portable._character_weights = weights.__getitem__
        return portable.portable_text_key(value)
    finally:
        portable._character_weights = original


def generate(destination):
    weights = {}
    rejected = []
    for code in range(1, 0x110000):
        character = chr(code)
        if unicodedata.category(character) in {'Cn', 'Cs', 'Co'}:
            continue
        if 0x1100 <= code <= 0x11ff or 0xa960 <= code <= 0xa97f or 0xd7b0 <= code <= 0xd7ff:
            rejected.append(code)
            continue
        base = levels(character)
        accent = levels(character + '\u00e9')
        case = levels(character + '\U0001f600')
        if any(len(v) != 5 for v in (base, accent, case)):
            rejected.append(code)
            continue
        extra = b''
        if base[3]:
            try:
                small, middle, width, tail = base[3].split(b'\xff')
                repeat, kana = middle.split(b'\x02')
                if tail or any(len(v) > 1 for v in (small, repeat, kana, width)):
                    raise ValueError()
                extra = (small or b'\xc6') + (repeat or b'\x03') + (kana or b'\xe4') + (width or b'\xc5')
            except ValueError:
                rejected.append(code)
                continue
        if not accent[1].endswith(b'\x0e') or not case[2].endswith(b'\x27\x27'):
            rejected.append(code)
            continue
        repeat_key = levels(character + '\u309d')
        if len(repeat_key) != 5 or not repeat_key[0].startswith(base[0]):
            rejected.append(code)
            continue
        repeat = repeat_key[0][len(base[0]):]
        weights[character] = (base[0], accent[1][:-1], case[2][:-2], base[4], extra, repeat)
    # ASCII anchors are always present. Failed context characters are removed
    # only after checking, so the result cannot depend on traversal order.
    bad = []
    for character in weights:
        for value in (character, 'a' + character, '\u00e9' + character,
                      character + '\u00e9', 'a' + character + '\u00e9',
                      character + '\u309d', character + '\u30fc', character + '\u3005'):
            actual = assemble(value, weights)
            if any(actual != _native_text_key(value, locale) for locale in ('en-US', 'en-GB')):
                bad.append(ord(character))
                break
    for code in bad:
        del weights[chr(code)]
    rejected.extend(bad)
    points, offsets, payload = [], [], bytearray()
    for character, entry in weights.items():
        points.append(ord(character))
        offsets.append(len(payload))
        payload.extend(bytes(map(len, entry)))
        for part in entry:
            payload.extend(part)
    raw = (b'EMLNLS02' + struct.pack('<I', len(points)) +
           struct.pack('<' + 'I' * len(points), *points) +
           struct.pack('<' + 'I' * len(offsets), *offsets) + payload)
    compressed = gzip.compress(raw, mtime=0)
    class Version(ctypes.Structure):
        _fields_ = [('size', ctypes.c_uint32), ('nls', ctypes.c_uint32),
                    ('defined', ctypes.c_uint32), ('lcid', ctypes.c_uint32),
                    ('guid', ctypes.c_ubyte * 16)]
    version = Version()
    version.size = ctypes.sizeof(version)
    get_version = ctypes.WinDLL('kernel32', use_last_error=True).GetNLSVersionEx
    if not get_version(1, ctypes.c_wchar_p('en-US'), ctypes.byref(version)):
        raise ctypes.WinError(ctypes.get_last_error())
    provenance = {
        'format': 'EMLNLS02', 'flags': 'LCMAP_SORTKEY | NORM_IGNORECASE',
        'source': 'Generated LCMapStringEx API results for synthetic character strings',
        'generator': 'tools/generate_nls_weights.py',
        'windows_version': platform.version(), 'nls_version': version.nls,
        'defined_version': version.defined, 'effective_lcid': version.lcid,
        'custom_version_guid': bytes(version.guid).hex(),
        'unicode_enumeration_version': unicodedata.unidata_version,
        'locales': ['en-US', 'en-GB'], 'character_count': len(weights),
        'sha256': hashlib.sha256(compressed).hexdigest(),
        'excluded_categories': ['unassigned', 'surrogates', 'private-use'],
        'contextual_resource': 'nls-jamo-v1.json',
        'contextual_generator': 'tools/generate_nls_jamo.py',
        'verification': 'Every retained character and eight contexts byte-match native NLS for both locales.',
        'scope': 'Pinned captured NLS version, not a claim of compatibility with every past or future NLS release.'}
    contextual = destination / 'nls-jamo-v1.json'
    if contextual.exists():
        provenance['contextual_character_count'] = len(json.loads(contextual.read_text()))
        provenance['contextual_sha256'] = hashlib.sha256(contextual.read_bytes()).hexdigest()
    destination.mkdir(parents=True, exist_ok=True)
    (destination / 'nls-en-v1.bin.gz').write_bytes(compressed)
    (destination / 'nls-en-v1.json').write_text(json.dumps(provenance, indent=2) + '\n')
    print(json.dumps({key: provenance[key] for key in ('character_count', 'nls_version', 'sha256')}))


if __name__ == '__main__':
    generate(Path(__file__).resolve().parents[1] / 'src/excel_model_lab/data')
