from decimal import Decimal

import pandas as pd
import pytest

from excel_model_lab import WorkbookModel


@pytest.mark.parametrize('values', [
    ['one'] * 4, ['one', 'two', 'one', 'two'],
    [None, None, None, None], ['one', None, 'one', None],
    ['one', 'two', 'three', 'four'],
    [2**60, 2**60, 2**60, 2**60],
    [2**60, 2**60+1, 2**60, 2**60+1],
    [Decimal('1000000000000.0001')] * 4,
])
def test_dictionary_contains_only_present_values(tmp_path, values):
    data = pd.DataFrame({'Value': pd.Series(values, dtype='string')}) if all(v is None for v in values) else pd.DataFrame({'Value': values})
    model = WorkbookModel.create({'T': data})
    model = WorkbookModel(model.save(tmp_path / 'created.xlsx'))
    encoding = model._encodings['T', 'Value']
    assert len(encoding.original_values) == len(set(v for v in values if v is not None))
    actual = model.get_table('T').Value.tolist()
    assert [pd.isna(v) for v in actual] == [v is None for v in values]
    assert [v for v in actual if not pd.isna(v)] == [v for v in values if v is not None]
    replacement = [values[0]] * len(values)
    model.replace_column('T', 'Value', replacement)
    reopened = WorkbookModel(model.save(tmp_path / 'replaced.xlsx'))
    assert len(reopened._encodings['T', 'Value'].original_values) == int(values[0] is not None)


def test_compaction_rebuilds_relationship_and_hierarchy_indexes(tmp_path):
    model = WorkbookModel.create({'Keys': pd.DataFrame({'Key': ['unused', 'b', 'c'], 'Rank': [1, 2, 3]}),
                                  'Facts': pd.DataFrame({'Key': ['b', 'b', 'c'], 'Amount': [10, 20, 30]})})
    model.add_relationship('Facts', 'Key', 'Keys', 'Key')
    model.add_hierarchy('Keys', 'Labels', ['Key'])
    model = WorkbookModel(model.save(tmp_path / 'before.xlsx'))
    model.replace_tables({'Keys': pd.DataFrame({'Key': ['b', 'c'], 'Rank': [2, 3]}),
                          'Facts': pd.DataFrame({'Key': ['b', 'b'], 'Amount': [10, 20]})})
    result = WorkbookModel(model.save(tmp_path / 'after.xlsx'))
    assert result._encodings['Keys', 'Key'].original_values == ('b', 'c')
    assert result._encodings['Facts', 'Key'].original_values == ('b',)
    assert result.get_table('Facts').Amount.tolist() == [10, 20]


def test_all_null_text_can_be_populated_repeatedly(tmp_path):
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': pd.Series([None] * 3, dtype='string')})})
    for index, values in enumerate(([None] * 3, ['a', None, 'b'], [None] * 3, ['c'] * 3)):
        model.replace_column('T', 'Value', values)
        model = WorkbookModel(model.save(tmp_path / f'stage-{index}.xlsx'))
        actual = model.get_table('T').Value.tolist()
        assert [None if pd.isna(v) else v for v in actual] == values
        encoding = model._encodings['T', 'Value']
        assert (encoding.dictionary_path is None) == all(v is None for v in values)


def test_native_blank_lookup_segment_capacity(tmp_path):
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _text
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': ['a', None, 'b']})})
    model = WorkbookModel(model.save(tmp_path / 'mixed.xlsx'))
    file = next(f for f in model._image.files if f.logical_path and '$Value.' in f.logical_path
                and f.logical_path.endswith('.tbl.xml'))
    doc = minidom.parseString(model._image.read_logical_file(file))
    assert [_text(n) for n in _elements(doc, 'RecordsPerSegment')] == ['3']
    assert model.get_table('T').Value.dropna().tolist() == ['a', 'b']


def test_numeric_dictionary_native_usage(tmp_path):
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _text
    values = [2**60, 2**60+1] * 3
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': values})})
    model = WorkbookModel(model.save(tmp_path / 'numeric-two.xlsx'))
    enc = model._encodings['T', 'Value']
    doc = minidom.parseString(model._image.read_logical_file(enc.metadata_path))
    raw = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn'
               and n.getAttribute('name') == 'Value')
    assert [_text(n) for n in _elements(raw, 'Usage')] == ['3']
    assert model.get_table('T').Value.tolist() == values

@pytest.mark.parametrize('kind', ['integer', 'currency', 'decimal'])
def test_nullable_numeric_dictionary_clear_and_repopulate(tmp_path, kind):
    from decimal import Decimal
    from xml.dom import minidom
    from excel_model_lab.schema import _elements, _text
    values = {'integer': [2**60, None, 2**60+1],
              'currency': [Decimal('1.25'), None, Decimal('100000000000.5')],
              'decimal': [1e100, None, 1.25]}[kind]
    dtype = {'integer': 'Int64', 'currency': object, 'decimal': 'Float64'}[kind]
    model = WorkbookModel.create({'T': pd.DataFrame({'Value': pd.Series(values, dtype=dtype)})})
    for index, replacement in enumerate((values, [None]*3, values, [None]*3)):
        model.replace_column('T', 'Value', replacement)
        model = WorkbookModel(model.save(tmp_path / f'{kind}-{index}.xlsx'))
        actual = [None if pd.isna(v) else v for v in model.get_table('T').Value]
        assert actual == replacement
        if kind == 'currency':
            assert all(isinstance(v, Decimal) for v in actual if v is not None)
        encoding = model._encodings['T', 'Value']
        assert (encoding.dictionary_path is None) == all(v is None for v in replacement)
        if kind != 'decimal':
            doc = minidom.parseString(model._image.read_logical_file(encoding.metadata_path))
            raw = next(n for n in _elements(doc, 'XMObject') if n.getAttribute('class') == 'XMRawColumn'
                       and n.getAttribute('name') == 'Value')
            assert [_text(n) for n in _elements(raw, 'OperatingOn32')] == ['false']
