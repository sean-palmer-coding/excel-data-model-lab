import struct
import pandas as pd
import pytest

from excel_model_lab import ExcelModelTemplate
from excel_model_lab.sorting import stored_sorts


def positions(model, column):
    file = next(f for f in model._image.files if f.logical_path and f'${column}.POS_TO_ID.' in f.logical_path)
    payload = model._image.read_logical_file(file)
    return struct.unpack('<3I', payload[8:20])


def test_text_sort_rebuild_clear_and_save(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'Label': ['Alpha', 'Beta', 'Gamma'], 'Order': [3, 2, 1]})})
    model.set_sort_column('T', 'Label', 'Order')
    model = ExcelModelTemplate(model.save(tmp_path / 'sorted.xlsx'))
    assert stored_sorts(model) == [('T', 'Label', 'Order')]
    assert positions(model, 'Label') == (5, 4, 3)
    model.replace_column('T', 'Order', [2, 1, 3])
    model = ExcelModelTemplate(model.save(tmp_path / 'resorted.xlsx'))
    assert positions(model, 'Label') == (4, 3, 5)
    model.set_sort_column('T', 'Label', None)
    model = ExcelModelTemplate(model.save(tmp_path / 'cleared.xlsx'))
    assert stored_sorts(model) == []
    assert positions(model, 'Label') == (3, 4, 5)


def test_numeric_sort_hash_and_lookup_agree(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2, 3], 'Order': [3, 1, 2]})})
    model.set_sort_column('T', 'ID', 'Order')
    model = ExcelModelTemplate(model.save(tmp_path / 'numeric-sort.xlsx'))
    assert positions(model, 'ID') == (4, 5, 3)
    with model._open_reader() as reader:
        assert reader.get_table('T')['ID'].tolist() == [1, 2, 3]


def test_sort_ambiguity_and_cycles_leave_model_intact(tmp_path):
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'Label': ['A', 'A', 'B'], 'Order': [1, 2, 3]})})
    before = model._image.data
    with pytest.raises(ValueError, match='exactly one'):
        model.set_sort_column('T', 'Label', 'Order')
    assert model._image.data == before
    model = ExcelModelTemplate.create({'T': pd.DataFrame({'ID': [1, 2, 3], 'Order': [3, 2, 1]})})
    model.set_sort_column('T', 'ID', 'Order')
    with pytest.raises(ValueError, match='circular'):
        model.set_sort_column('T', 'Order', 'ID')
