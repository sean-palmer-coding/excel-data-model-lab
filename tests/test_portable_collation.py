import json
import ctypes
import os
import random
from importlib.resources import files

import pytest

from excel_model_lab.abf import UnsupportedModelError
from excel_model_lab.collation import _native_text_key, string_order
from excel_model_lab.portable_nls import portable_text_key, _weights


@pytest.mark.parametrize('value,hex_key', [
    ('a-b', '0e020e0901010101fffe821200'),
    ('caf\u00e9', '0e0a0e020e230e21010202020e01010100'),
    ('\u6771\u4eac', 'fdff560ffdff3d180101010100'),
    ('\U0001f600', 'b03fb6fc013f3f012727010100'),
    ('\u00df', '0e910e910101010100'),
    ('\u3042\u30a2', '22022202010101ff02e4c4ffff0100'),
    # NLS 394498 changed this weight; the portable resource pins 394243.
    ('\U00030ab9', 'b7b7013f0127010100'),
])
def test_captured_native_keys(value, hex_key):
    assert portable_text_key(value).hex() == hex_key
    assert portable_text_key(value, 'en-GB').hex() == hex_key


def test_canonical_accents_and_ligatures():
    assert portable_text_key('CAF\u00c9') == portable_text_key('cafe\u0301')
    assert portable_text_key('Stra\u00dfe') == portable_text_key('strasse')
    assert portable_text_key('bill-s') > portable_text_key('bills')
    assert portable_text_key('cafe') != portable_text_key('caf\u00e9')


@pytest.mark.parametrize('value', ['\x00', '\ud800', '\ue000', '\u0378'])
def test_unsupported_characters_fail_without_approximation(value):
    with pytest.raises(UnsupportedModelError):
        portable_text_key(value)


def test_provenance_covers_packaged_data():
    import hashlib
    root = files('excel_model_lab').joinpath('data')
    metadata = json.loads(root.joinpath('nls-en-v1.json').read_text())
    assert hashlib.sha256(root.joinpath('nls-en-v1.bin.gz').read_bytes()).hexdigest() == metadata['sha256']
    assert metadata['locales'] == ['en-US', 'en-GB']
    assert metadata['character_count'] == len(_weights()[0])
    contextual = root.joinpath(metadata['contextual_resource']).read_bytes()
    assert hashlib.sha256(contextual).hexdigest() == metadata['contextual_sha256']
    assert len(json.loads(contextual)) == metadata['contextual_character_count']


def _native_nls_identity(locale):
    class Version(ctypes.Structure):
        _fields_ = [('size', ctypes.c_uint32), ('nls', ctypes.c_uint32),
                    ('defined', ctypes.c_uint32), ('lcid', ctypes.c_uint32),
                    ('guid', ctypes.c_ubyte * 16)]

    value = Version()
    value.size = ctypes.sizeof(value)
    api = ctypes.WinDLL('kernel32', use_last_error=True).GetNLSVersionEx
    api.argtypes = [ctypes.c_int, ctypes.c_wchar_p, ctypes.POINTER(Version)]
    api.restype = ctypes.c_bool
    if not api(1, locale, ctypes.byref(value)):
        raise ctypes.WinError(ctypes.get_last_error())
    # An unreadable identity is a failure, never permission to skip parity.
    assert value.nls and value.defined and value.lcid
    return dict(nls_version=value.nls, defined_version=value.defined,
                effective_lcid=value.lcid, custom_version_guid=bytes(value.guid).hex())


def _require_pinned_native_version():
    metadata = json.loads(files('excel_model_lab').joinpath('data/nls-en-v1.json').read_text())
    for locale, lcid in [('en-US', 1033), ('en-GB', 2057)]:
        actual = _native_nls_identity(locale)
        expected = {key: metadata[key] for key in actual}
        expected['effective_lcid'] = lcid
        if actual != expected:
            pytest.skip(f'Full Unicode parity requires pinned NLS identity {expected}; '
                        f'{locale} host identity is {actual}. Stable/context corpora remain tested.')


def test_native_version_gate_reports_drift_and_preserves_errors(monkeypatch):
    metadata = json.loads(files('excel_model_lab').joinpath('data/nls-en-v1.json').read_text())

    def identity(locale):
        return dict(nls_version=metadata['nls_version'],
                    defined_version=metadata['defined_version'],
                    effective_lcid={'en-US': 1033, 'en-GB': 2057}[locale],
                    custom_version_guid=metadata['custom_version_guid'])

    monkeypatch.setitem(_require_pinned_native_version.__globals__, '_native_nls_identity', identity)
    _require_pinned_native_version()

    def changed_identity(locale):
        return dict(identity(locale), nls_version=394498)

    monkeypatch.setitem(_require_pinned_native_version.__globals__, '_native_nls_identity', changed_identity)
    with pytest.raises(pytest.skip.Exception, match='394498'):
        _require_pinned_native_version()

    def unreadable_identity(locale):
        raise OSError('Native identity unavailable')

    monkeypatch.setitem(_require_pinned_native_version.__globals__, '_native_nls_identity', unreadable_identity)
    with pytest.raises(OSError, match='Native identity unavailable'):
        _require_pinned_native_version()


@pytest.mark.skipif(os.name != 'nt', reason='Native NLS comparison requires Windows')
@pytest.mark.parametrize('corpus_index', range(4),
                         ids=['pinned-full-unicode', 'basic-and-combining',
                              'mixed-scripts', 'contextual-kana'])
def test_randomized_native_parity_for_both_locales(corpus_index):
    if corpus_index == 0:
        _require_pinned_native_version()
    rng = random.Random(34827)
    points = _weights()[0]
    corpora = [list(map(chr, points)),
               [chr(p) for p in points if p < 0x3100],
               list('Ab-caf\u00e9\u0301\u6771\u4eac\u3042\u3041\u30a2\u30ac\U0001f600'),
               list('\u3005\u3031\u3032\u309d\u309e\u30fc\u30fd\u30fe\uff70\u3041\u3042\u30a2\u30ac\u0301\U000e0100abc-')]
    for index, characters in enumerate(corpora):
        for _ in range(1500):
            value = ''.join(rng.choices(characters, k=rng.randrange(1, 25)))
            if index != corpus_index:
                continue
            expected = portable_text_key(value)
            for locale in ('en-US', 'en-GB'):
                assert expected == _native_text_key(value, locale), (locale, ascii(value))


def test_en_gb_unicode_is_supported():
    assert string_order(['\u6771\u4eac', 'caf\u00e9'], 'en-GB') == [1, 0]


@pytest.mark.parametrize('value,hex_key', [
    ('\u1100', 'c002ff83ff03ff020101010100'),
    ('\u1113\u1161\u11a8', 'c6faff07ff03ff030101010100'),
])
def test_captured_jamo_composition(value, hex_key):
    assert portable_text_key(value).hex() == hex_key


@pytest.mark.skipif(os.name != 'nt', reason='Native NLS comparison requires Windows')
def test_jamo_context_native_parity():
    from excel_model_lab.portable_nls import _jamo_states
    rng = random.Random(44)
    chars = list(map(chr, _jamo_states())) + list('a-\u00e9\u0301\u30a2\u309d\u30fc\U0001f600')
    for _ in range(5000):
        value = ''.join(rng.choices(chars, k=rng.randrange(1, 25)))
        expected = portable_text_key(value)
        assert expected == _native_text_key(value, 'en-US')
        assert expected == _native_text_key(value, 'en-GB')
