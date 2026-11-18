"""Report native NLS identity and the first bounded portable/native mismatch."""
import ctypes
import json
import platform
import random
from importlib.resources import files
from excel_model_lab.collation import _native_text_key
from excel_model_lab.portable_nls import portable_text_key, _weights

class Version(ctypes.Structure):
    _fields_ = [('size', ctypes.c_uint32), ('nls', ctypes.c_uint32),
                ('defined', ctypes.c_uint32), ('lcid', ctypes.c_uint32),
                ('guid', ctypes.c_ubyte * 16)]

def version(locale):
    value = Version(); value.size = ctypes.sizeof(value)
    api = ctypes.WinDLL('kernel32', use_last_error=True).GetNLSVersionEx
    api.argtypes = [ctypes.c_int, ctypes.c_wchar_p, ctypes.POINTER(Version)]
    api.restype = ctypes.c_bool
    if not api(1, locale, ctypes.byref(value)):
        raise ctypes.WinError(ctypes.get_last_error())
    return dict(nls_version=value.nls, defined_version=value.defined,
                effective_lcid=value.lcid, custom_version_guid=bytes(value.guid).hex())

metadata = json.loads(files('excel_model_lab').joinpath('data/nls-en-v1.json').read_text())
print(json.dumps(dict(os=platform.platform(), pinned=metadata), ensure_ascii=True), flush=True)
points = _weights()[0]
corpora = [list(map(chr, points)), [chr(p) for p in points if p < 0x3100],
           list('Ab-caf\u00e9\u0301\u6771\u4eac\u3042\u3041\u30a2\u30ac\U0001f600'),
           list('\u3005\u3031\u3032\u309d\u309e\u30fc\u30fd\u30fe\uff70\u3041\u3042\u30a2\u30ac\u0301\U000e0100abc-')]
for locale in ('en-US', 'en-GB'):
    print(json.dumps(dict(locale=locale, native=version(locale))), flush=True)
    rng = random.Random(34827); mismatch = None
    for corpus_index, characters in enumerate(corpora):
        for sample_index in range(1500):
            value = ''.join(rng.choices(characters, k=rng.randrange(1, 25)))
            expected, actual = portable_text_key(value, locale), _native_text_key(value, locale)
            if expected != actual:
                mismatch = dict(locale=locale, corpus=corpus_index, sample=sample_index,
                    value=value.encode('unicode_escape').decode('ascii'),
                    portable=expected.hex(), native=actual.hex(), characters=[])
                for character in dict.fromkeys(value):
                    portable, native = portable_text_key(character, locale), _native_text_key(character, locale)
                    mismatch['characters'].append(dict(codepoint=f'U+{ord(character):04X}',
                        portable=portable.hex(), native=native.hex(), equal=portable==native))
                break
        if mismatch: break
    print(json.dumps(mismatch or dict(locale=locale, parity='all 6000 cases passed'), ensure_ascii=True), flush=True)
