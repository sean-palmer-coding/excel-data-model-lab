"""Case-insensitive, accent-sensitive default en-US/en-GB model collation.

Windows uses NLS. Other operating systems use pinned NLS weight data, never an
ICU/casefold approximation. Unsupported contextual characters fail explicitly.
"""
import os
from functools import lru_cache
from xml.dom import minidom
from .abf import UnsupportedModelError
from .schema import _elements, _child, _text

_LOCALES = {'1033': 'en-US', '2057': 'en-GB', 'en-US': 'en-US', 'en-GB': 'en-GB'}


def _locale(value):
    try:
        return _LOCALES[str(value)]
    except KeyError as exc:
        raise UnsupportedModelError('String writing supports default en-US and en-GB collation only') from exc


@lru_cache(maxsize=1)
def _mapper():
    import ctypes
    mapper = ctypes.WinDLL('kernel32', use_last_error=True).LCMapStringEx
    mapper.argtypes = [ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_wchar_p,
        ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p,
        ctypes.c_void_p, ctypes.c_longlong]
    mapper.restype = ctypes.c_int
    return mapper


def _native_text_key(value, locale='en-US'):
    import ctypes
    mapper = _mapper()
    flags = 0x400 | 1  # LCMAP_SORTKEY | NORM_IGNORECASE
    size = mapper(locale, flags, value, -1, None, 0, None, None, 0)
    if not size:
        raise ctypes.WinError(ctypes.get_last_error())
    buffer = ctypes.create_string_buffer(size)
    if not mapper(locale, flags, value, -1, buffer, size, None, None, 0):
        raise ctypes.WinError(ctypes.get_last_error())
    return buffer.raw


@lru_cache(maxsize=8192)
def text_key(value, locale='en-US'):
    """Return a sortable/hashable key under the selected model collation."""
    locale = _locale(locale)
    if not isinstance(value, str):
        raise TypeError('Collation values must be strings')
    if '\x00' in value:
        raise UnsupportedModelError('Embedded NUL characters are not supported in model strings')
    if os.name == 'nt':
        return _native_text_key(value, locale)
    from .portable_nls import portable_text_key
    return portable_text_key(value, locale)


def string_order(values, locale='en-US'):
    locale = _locale(locale)
    keys = [text_key(value, locale) for value in values]
    order = sorted(range(len(values)), key=keys.__getitem__)
    for a, b in zip(order, order[1:]):
        if keys[a] == keys[b] and values[a] != values[b]:
            raise UnsupportedModelError(
                f'Distinct strings compare equal under model collation: {values[a]!r}, {values[b]!r}; '
                'normalize equivalent spellings before writing')
    return order


def model_string_order(image, table_id, values):
    for file in image.files:
        if not file.logical_path or not file.logical_path.endswith('.dim.xml'):
            continue
        doc = minidom.parseString(image.read_logical_file(file))
        for dimension in _elements(doc, 'Dimension'):
            if _text(_child(dimension, 'ID')) != table_id:
                continue
            language = _text(_child(dimension, 'Language'))
            if any(_text(n) for n in _elements(dimension, 'Collation')):
                raise UnsupportedModelError('Explicit custom model collations are not supported')
            return string_order(values, language)
    raise UnsupportedModelError('Cannot resolve model collation for string dictionary')
