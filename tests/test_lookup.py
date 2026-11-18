from pathlib import Path
import struct

from excel_model_lab.abf import AbfImage
from excel_model_lab.lookup import numeric_lookup_replacements


FIXTURE = Path(__file__).parent / 'data' / 'native-authoring-base.xlsx'


def test_mapping_uses_sorted_distinct_ids_and_word_padding():
    image = AbfImage.from_workbook(FIXTURE)
    changes = numeric_lookup_replacements(image, 'Table1', 'ID', [32, 12, 22, 12])
    stream = next(value for key, value in changes.items()
                  if next(f.logical_path for f in image.files if f.storage_path == key).endswith('.idf'))
    assert struct.unpack('<12I', stream) == (2, 0, 12, 22, 32, 0, 2, 0, 0, 0, 0, 0)


def test_mapping_matches_native_binary_for_unchanged_ids():
    image = AbfImage.from_workbook(FIXTURE)
    changes = numeric_lookup_replacements(image, 'Table1', 'ID', [3, 4, 5])
    for key, value in changes.items():
        if next(f.logical_path for f in image.files if f.storage_path == key).endswith('.idf'):
            assert value == image.read_logical_file(key)



def test_string_lookup_reproduces_native_seed_bytes():
    from excel_model_lab.model import ExcelModelTemplate
    from excel_model_lab.lookup import string_lookup_replacements
    from pathlib import Path
    model = ExcelModelTemplate(Path(__file__).parent / 'data' / 'native-authoring-base.xlsx')
    image = model._image
    replacements = string_lookup_replacements(image, 'Table1', 'Label', ['Alpha', 'Beta', 'Gamma'])
    for key, value in replacements.items():
        if next(f for f in image.files if f.storage_path == key).logical_path.endswith('.idf'):
            assert value == image.read_logical_file(key)


def test_small_numeric_lookup_matches_native_and_can_grow_again():
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _child, _text

    image = AbfImage.from_workbook(FIXTURE)
    for ids, expected_counts in [([3, 4], [2, 2, 1]), ([3], [1, 1, 1, 1]),
                                 ([3, 4, 5, 6], [4, 3])]:
        changes = numeric_lookup_replacements(image, 'Table1', 'ID', ids)
        meta = next(value for value in changes.values() if value.startswith(b'<'))
        doc = minidom.parseString(meta)
        segments = [n for n in _elements(doc, 'XMObject')
                    if n.getAttribute('class') == 'XMColumnSegment']
        counts = [int(_text(_child(_child(n, 'Properties'), 'Records'))) for n in segments]
        assert counts == expected_counts
        assert int(_text(_elements(doc, 'Segments')[0])) == len(counts)
        assert int(_text(_elements(doc, 'SegmentCount')[0])) == len(counts)
        stream = next(value for value in changes.values() if not value.startswith(b'<'))
        if ids == [3, 4]:
            assert stream == (FIXTURE.parent / 'native-two-values-pos.idf').read_bytes()
        image = AbfImage(image.assemble(changes))

